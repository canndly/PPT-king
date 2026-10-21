# ppt-king

PPTX 与 Markdown 的双向转换：**结构抽取与版式渲染是确定性的，格式映射交给大模型**。

- `parse`：PPTX → Markdown + 图片（模型负责元素到 Markdown 的映射）。
- `render`：**16×9 网格 Markdown → PPTX**（纯本地，不调模型）。
- `layout`：让模型给普通 Markdown 补上网格坐标，再交给 `render`。
- `build`：导出目录 → 版式 JSON → PPTX（另一种写法，见文末）。

## 安装

```bash
uv sync
cp .env.example .env   # 只有 parse / layout / build 需要 Key
```

## 16×9 网格（推荐）

16:9 画布按 16 列 9 行等分，格子恰好是正方形（762000 EMU ≈ 0.83 英寸）。
坐标从 1 开始，`1,1` 是左上角。

### 语法只有三条

```markdown
# 演示文稿标题

## 一、题目解析

@ 1,2 8x6
本设计要求基于 LM5180 设计隔离型电源系统。

反激电路适合中小功率隔离电源场合。

@ 9,3 7x5
![设计目标](assets/slide02_img03.png)

@ 1,6 16x4
| 输入电压 | 输出电压 | 效率 |
| --- | --- | --- |
| 24 V | 5.036 V | 81.24% |
```

| 写法 | 含义 |
| --- | --- |
| `# 标题` | 演示文稿标题（可省略） |
| `## 页标题` | 开始新的一页；标题固定渲染在第 1 行 |
| `##` | 无标题页（封面、Thanks 页） |
| `@ 列,行` | 位置，自动延伸到右边界/底边界 |
| `@ 列,行 宽x高` | 位置 + 跨度 |
| `@ 列,行 宽x高 center` | 再加对齐（`left`/`center`/`right`） |
| `@ 列,行 宽x高 center 34` | 再加字号（pt） |

块类型由 Markdown 自身决定：段落、列表、图片、表格、引用、代码。
**空行不切块**——一个 `@` 到下一个 `@`/`##` 之间的全部内容属于同一个内容区，
内部保留段落、列表、`###` 小标题的层次。

### 不写 `@` 也能出片

没有坐标的块由渲染器按内容形态自动组合：

| 页面内容 | 自动版式 |
| --- | --- |
| 只有文字 | 单栏铺满 |
| 1 段文字 + 1 张图 | 左文右图 |
| 只有 1 张图 | 居中放大 |
| 只有 2 张图 | 左右并排 |
| 只有 ≥3 张图 | 两列网格 |
| 只有表格 | 整宽表格 |
| 文字 + 表格/多图 | 文字在上，其余依次向下 |

## 用法

```bash
# 直接渲染（不需要 API Key）
uv run ppt-king render export/某报告/某报告.md
uv run ppt-king render deck.md -o out.pptx --accent 2E5C8A
uv run ppt-king render deck.md --check          # 只看解析结果，不渲染

# PPTX -> Markdown
uv run ppt-king parse example/某报告.pptx
uv run ppt-king parse example/*.pptx -o out --concurrency 8
uv run ppt-king parse example/某报告.pptx --stdout
uv run ppt-king parse example/某报告.pptx --drop-repeated-images
uv run ppt-king parse example/某报告.pptx --debug-dump prompts

# 让模型补坐标，再渲染
uv run ppt-king layout export/某报告/某报告.md --render
```

产物：

```
export/<ppt 名>/
├── <ppt 名>.md
├── <ppt 名>.pptx     # render 的默认输出
├── assets/           # 图片
└── media/            # 音视频、嵌入对象
```

## 主题

`GridTheme`（`grid_render.py`）：白底、深灰标题、正文 17pt、表头 `#2E5C8A` 白字、隔行浅灰。
`--accent` 改强调色与表头色。文字超出框时按 1pt 递降，最低 9pt；仍溢出会在命令行报告。

## 配置（.env）

| 变量 | 说明 | 默认 |
| --- | --- | --- |
| `PPT_KING_API_KEY` | 接口密钥（必填） | — |
| `PPT_KING_BASE_URL` | OpenAI 兼容地址 | `https://api.deepseek.com` |
| `PPT_KING_MODEL` | 模型名 | `deepseek-flash` |
| `PPT_KING_TEMPERATURE` | 采样温度，转换任务建议 0 | `0` |
| `PPT_KING_MAX_TOKENS` | 单页输出上限 | `8000` |
| `PPT_KING_TIMEOUT` | 单次请求超时（秒） | `180` |
| `PPT_KING_MAX_RETRIES` | 429/5xx/超时重试次数 | `3` |
| `PPT_KING_CONCURRENCY` | 并发页数 | `4` |
| `PPT_KING_CACHE_DIR` | 页级缓存目录 | 系统临时目录 |

无 Key 或调用失败会直接报错，不做静默降级。

## 映射规则

| PPT 元素 | Markdown |
| --- | --- |
| 幻灯片 | `## N. 标题`（无标题为 `## 幻灯片 N`） |
| 正文/列表 | 段落或 `-` / `1.` 列表，按 `level` 每级缩进 2 空格 |
| 粗体/斜体/等宽 | `**…**` / `*…*` / 反引号 |
| 表格 | GFM 表格，首行为表头，合并单元格保留主值 |
| 图片 | `![alt](assets/slideNN_imgNN.ext)` |
| 图表 | `### 图表：标题` + 分类×系列数据表 |
| 演讲者备注 | `> **备注**：…` |
| 音视频/嵌入对象 | `[名称](media/…)` |

不支持：旧版 `.ppt`（请先转 `.pptx`）、SmartArt、动画、渲染预览。

## 还原（build）的版式契约

模型每页输出一个 JSON（`spec.json` 可存档、可手改后重新渲染）：

```json
{
  "layout": "two_column",
  "subtitle": "",
  "blocks": [
    {"type": "heading", "text": "设计要求"},
    {"type": "paragraph", "text": "输入 24 V，输出 5 V/1 A。"},
    {"type": "bullets", "items": [{"text": "隔离型电源", "level": 0}]},
    {"type": "image", "ref": "A1", "caption": "原理图"},
    {"type": "table", "header": ["分类", "效率"], "rows": [["0.2A", "80.1"]]},
    {"type": "quote", "text": "说明"}
  ],
  "notes": ""
}
```

| layout | 用途 |
| --- | --- |
| `title` | 封面：大标题 + 副标题 + 底部图片条 |
| `section` | 章节过渡页 |
| `title_content` | 单列纵向流式排版（默认） |
| `two_column` / `image_left` | 文图分栏 |
| `gallery` | 图片网格（≥3 张图） |
| `table` | 以表格为主 |

渲染器为 16:9（12192000×6858000 EMU），统一字体、标题带、强调色、表头底色与页码；
文字高度按字符宽度估算，必要时自动降字号以避免溢出。

## 开发

```bash
uv run pytest
```

模块划分：

- 网格方向：`grid`（语法解析）、`grid_render`（确定性渲染）、`grid_ai`（模型补坐标）。
- 解析方向：`reader`（pptx→IR）、`serialize`/`prompt`（IR→载荷/提示词）、`markdown`（清洗、校验、占位符替换）、`exporter`（编排）。
- 还原方向：`composer`（md→版式 JSON）、`spec`（JSON 契约与校验）、`builder`（确定性渲染）、`rebuild`（编排）。
- 公共：`config`、`llm`、`cache`、`assets`、`cli`。

注：`composer`/`spec`/`builder`/`rebuild` 是早期的「7 种预设版式 + JSON 契约」路线，
与新网格语法并存；网格语法更简单，日常优先使用它。
