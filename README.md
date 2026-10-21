# ppt-king

从资料文件夹生成 PPT：**AI 写带坐标的 Markdown，渲染器确定性地出片**。

- `generate`：资料目录 → 大纲 → 逐页网格 Markdown → PPTX（一次跑完）。
- `render`：16×9 网格 Markdown → PPTX，纯本地，不调用模型。
- `layout`：给普通 Markdown 补网格坐标，再交给 `render`。
- `parse`：PPTX → Markdown + 图片，供重排或二次生成。

## 安装

```bash
uv sync
cp .env.example .env   # generate / layout / parse 需要 Key
```

也可以每次用 `--api-key` 传入，不写 `.env`。

## 从资料生成 PPT

### 最小用法

```bash
uv run ppt-king generate ./资料 --api-key sk-xxx
# 默认输出：out/资料/资料.pptx
```

资料目录约定：

```text
资料/
├── outline.md        # 可选，也可以叫 大纲.md
├── 项目说明.md
├── 实验数据.txt
└── 架构图.png
```

- 受众、汇报场景、语气等要求直接写进 `outline.md`，不再通过命令行参数指定。
- 大纲优先使用 `--outline` 指定的文件。
- 没有指定时，只在资料目录根目录查找 `outline.md`、`大纲.md`、`outline.txt`、`大纲.txt`。
- 子目录里的大纲不会自动生效；需要时用 `--outline` 显式指定。
- 都没有时，由 AI 先对资料做摘要，再归纳大纲。
- 大纲格式：

```markdown
# 演示文稿标题

## 封面

- 副标题
- 汇报人 / 日期

## 研究背景

- 行业现状
- 现有方案的不足

## 总结

- 主要结论
```

### 常用参数

```bash
uv run ppt-king generate ./资料 \
  --api-key sk-xxx \
  --outline ./资料/大纲.md \
  --out ./out/我的汇报 \
  --slides 12-18 \
  --accent 2E5C8A
```

| 参数 | 说明 |
| --- | --- |
| `--outline PATH` | 显式指定大纲文件 |
| `-o, --out DIR` | 输出目录，默认 `out/<资料目录名>` |
| `--slides 12` / `12-18` | 页数约束，默认 `8-20`；只约束 AI 生成的大纲 |
| `--language` | 正文语言，默认 `zh` |
| `--accent` | 强调色 / 表头色，6 位十六进制 |
| `--logo PATH` | 覆盖默认 Logo（默认 `static/page_logo.png`，固定左上角 8×1） |
| `--no-logo` | 不显示默认 Logo |
| `--header TEXT` | 可选全局页眉；页面里也可以用 `@ 9,1-16,1` 单独填写右上角槽位 |
| `--resume` | 复用已有大纲和逐页结果 |
| `--force` | 忽略已有中间产物，全部重做 |
| `--strict` | 任意一页失败即整体失败 |
| `--no-cache` | 禁用页级响应缓存 |
| `--debug-dump DIR` | 保存每页请求与模型输出，便于排查 |
| `--model` / `--base-url` / `--api-key` | 覆盖 `.env` 配置 |

### 产物

```text
out/<资料目录名>/
├── outline.md          # 使用或生成的大纲，可手改
├── pages/
│   ├── p01.md          # 每页独立网格 Markdown
│   └── p02.md
├── slides.grid.md      # 合并后的完整网格 Markdown，可手改
├── assets/             # 从资料复制、去重后的图片
├── generate.json       # 参数、页数、失败页、token 统计
└── <资料目录名>.pptx
```

`--resume` 的行为：

- 已有 `slides.grid.md` 时直接复用它渲染，不再调用模型。
- 没有完整网格时，复用 `pages/pNN.md`，只重写缺失或失败的页。
- 手改 `outline.md` 后想重新生成正文，请加 `--force`。

单页失败时会写入占位内容并继续，命令行以非零退出码提示；修复资料或提示词后加 `--resume` 重跑即可。需要整体失败时用 `--strict`。

### 内部流程

```text
资料目录
  │ sources.scan_sources：读取 md/txt，复制图片，发现/校验大纲
  ▼
SourceBundle
  │ outline.generate_outline：逐份摘要 → 归纳大纲（或直接解析人工大纲）
  ▼
Outline
  │ writing.select_evidence：按标题与要点关键词挑选资料段落
  │ writing.write_page：逐页调用模型，输出 @ 坐标块
  │ parse_grid 校验坐标、块类型、图片白名单；失败自动纠错一次
  ▼
slides.grid.md
  │ grid_render.render_deck：确定性渲染
  ▼
.pptx
```

设计要点：

- **逐页生成**：每页独立调用、独立落盘，可并发、可重试、可断点续跑。
- **上下文受控**：先摘要再写页，写页只带相关段落，不把整个资料库塞给模型。
- **自动封面**：大纲缺少封面页时程序自动补一页；封面不渲染左上角的“封面”小标题。
- **顶栏槽位**：默认 Logo `static/page_logo.png` 占左上角 8×1，`--logo` 可覆盖，`--no-logo` 可关闭；右上角 8×1 可在页面里用 `@ 9,1-16,1` 填写章节名或一句结论，也可用 `--header` 统一填写。页面标题在第 2 行，正文从第 3 行开始。
- **底部安全区**：内容底边最多到第 8 行底部，第 9 行留空，避免正文贴边。
- **图片白名单**：模型只能引用程序复制到 `assets/` 的图片，编造路径会被拒绝。
- **表格优先**：同类指标达到 3 项时优先输出表格；同一张表跨页重复会被自动去重。
- **渲染确定性**：坐标、字号、主题、表格样式都由本地代码计算，模型不直接画图。

## 16×9 网格语法

16:9 画布按 16 列 9 行等分，格子是正方形（762000 EMU ≈ 0.83 英寸）。坐标从 1 开始，`1,1` 是左上角。

```markdown
# 演示文稿标题

## 一、题目解析

@ 1,2-8,7
本设计要求基于 LM5180 设计隔离型电源系统。

反激电路适合中小功率隔离电源场合。

@ 9,3-15,7
![设计目标](assets/slide02_img03.png)

@ 1,6-16,8
| 输入电压 | 输出电压 | 效率 |
| --- | --- | --- |
| 24 V | 5.036 V | 81.24% |
```

| 写法 | 含义 |
| --- | --- |
| `# 标题` | 演示文稿标题（可省略） |
| `## 页标题` | 开始新的一页；标题固定渲染在第 1 行 |
| `##` | 无标题页（封面、Thanks 页） |
| `@ 左,上-右,下` | 左上角到右下角，两个角都包含在块内（推荐） |
| `@ 左,上-右,下 center` | 再加对齐（`left`/`center`/`right`） |
| `@ 左,上-右,下 center 34` | 再加字号（pt） |
| `@ 列,行 宽x高` | 旧写法，仍然兼容 |

例如 `@ 1,6-16,8` 表示从第 1 列第 6 行到第 16 列第 8 行。

块类型由 Markdown 自身决定：段落、列表、图片、表格、引用、代码。空行不切块，一个 `@` 到下一个 `@` / `##` 之间的内容属于同一个内容区。

行内格式：

- `**加粗**`
- `<red>红色强调</red>`，用于关键结论或数字
- 行内反引号会被自动去掉，公式直接写纯文本，例如 `m·v·ω ≤ μmg`

公式建议单独成块，放在两段文字之间，渲染为居中斜体：

```markdown
@ 1,5-16,5 center 24
$$ m·v·ω ≤ μmg  ⇒  v·ω ≤ μg $$
```

不写 `@` 也能出片，渲染器会按内容形态自动组合：

| 页面内容 | 自动版式 |
| --- | --- |
| 只有文字 | 单栏铺满 |
| 1 段文字 + 1 张图 | 左文右图 |
| 只有 1 张图 | 居中放大 |
| 只有 2 张图 | 左右并排 |
| 只有 ≥3 张图 | 两列网格 |
| 只有表格 | 整宽表格 |
| 文字 + 表格/多图 | 文字在上，其余依次向下 |

## 直接渲染

```bash
uv run ppt-king render out/我的汇报/slides.grid.md
uv run ppt-king render deck.md -o out.pptx --accent 2E5C8A
uv run ppt-king render deck.md --check      # 只看解析结果，不渲染
```

图片路径不存在时不会中断渲染，会打印 `图片不存在：P03 assets/xxx.png` 警告，
并在原位置显示 `[图片缺失：assets/xxx.png]` 占位。`generate` 和
`layout --render` 同样会输出该警告。

## 给普通 Markdown 补坐标

```bash
uv run ppt-king layout report.md            # 产出 report.grid.md
uv run ppt-king layout report.md --render   # 产出后直接渲染 PPTX
```

需要 Key。模型只输出块和坐标，页标题由程序统一拼接。

## 解析现有 PPTX

```bash
uv run ppt-king parse example/某报告.pptx
uv run ppt-king parse example/*.pptx -o out --concurrency 8
uv run ppt-king parse example/某报告.pptx --stdout
```

产物：

```text
export/<ppt 名>/
├── <ppt 名>.md
├── assets/
└── media/
```

映射规则：

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

## 主题

`GridTheme`（`grid_render.py`）：白底、Times New Roman、黑色标题与正文、正文默认 20pt、图注 18pt、表头 `#2E5C8A` 白字。`<red>` 强调会自动加粗为纯红 `#FF0000`。`--accent` 可覆盖强调色与表头色。文字超出框时从 20pt 按 1pt 递降，最低 15pt；表格会先按单元格换行计算每行高度，再决定字号，仍溢出会在命令行报告。

## 配置（.env）

| 变量 | 说明 | 默认 |
| --- | --- | --- |
| `PPT_KING_API_KEY` | 接口密钥（必填） | — |
| `PPT_KING_BASE_URL` | OpenAI 兼容地址 | `https://api.deepseek.com` |
| `PPT_KING_MODEL` | 模型名 | `deepseek-flash` |
| `PPT_KING_TEMPERATURE` | 采样温度 | `0` |
| `PPT_KING_MAX_TOKENS` | 单页输出上限 | `8000` |
| `PPT_KING_TIMEOUT` | 单次请求超时（秒） | `180` |
| `PPT_KING_MAX_RETRIES` | 429/5xx/超时重试次数 | `3` |
| `PPT_KING_CONCURRENCY` | 并发页数（generate 默认使用，无需命令行指定） | `4` |
| `PPT_KING_CACHE_DIR` | 页级缓存目录 | 系统临时目录 |

无 Key 或调用失败会直接报错，不做静默降级。

## 当前范围与限制

- 资料格式：第一版支持 `.md` / `.markdown` / `.txt` 与常见图片（png/jpg/jpeg/gif/bmp/webp/svg）。
- PDF、DOCX、PPTX 资料暂不内置解析；可通过 `sources.register_extractor()` 扩展。
- 扫描版 PDF 即使后续支持抽取，也需要 OCR，本项目不提供。
- 不生成图片，只使用资料目录中已有的图片。
- 不支持旧版 `.ppt`、SmartArt、动画与渲染预览。
- 输出目录不能与资料目录相同，避免覆盖原始资料。

## 开发

```bash
uv run pytest
```

模块划分：

- 生成方向：`sources`（资料摄入）、`outline`（大纲）、`writing`（逐页写作）、`generate`（编排）。
- 网格方向：`grid`（语法解析）、`grid_render`（确定性渲染）、`grid_ai`（模型补坐标）。
- 解析方向：`reader`（pptx→IR）、`serialize`/`prompt`（IR→载荷/提示词）、`markdown`（清洗、校验、占位符替换）、`exporter`（编排）。
- 公共：`config`、`llm`、`cache`、`assets`、`cli`。
