# AGENTS.md — 给 AI 编码 Agent 的项目说明

## 这个项目是什么

`ppt-king` 是一个“资料 → PPT”的工具：

```text
资料/ 中的 PDF / Word / PowerPoint / Excel / HTML / 图片 / Markdown
  → AI 生成大纲
  → AI 逐页写带 16×9 坐标的 Markdown
  → 本地确定性渲染为 .pptx
```

用户通常会这样说：

> 阅读项目，根据“资料文件夹”中的内容给我生成 ppt。

## 你的工作流程

### 1. 先读项目

- 阅读 [`README.md`](README.md) 了解完整用法。
- 阅读 [`static/prompt.md`](static/prompt.md) 了解页面风格偏好；除非用户明确要求，不要修改它。
- 不要为了“完成任务”而修改源码；用户要的是成品 PPT。

### 2. 运行自检

```bash
uv run ppt-king doctor
```

它会检查：

- `.env` 是否存在
- `PPT_KING_API_KEY` 是否已填写
- `资料/` 是否存在、是否有可用文件
- 页面模板与默认 Logo 是否就绪

根据结果处理：

| 自检结果 | 你的动作 |
| --- | --- |
| `uv` 命令不存在 | 请用户安装 [uv](https://docs.astral.sh/uv/)，不要改用 pip 破坏环境 |
| `.env` 不存在 | 执行 `cp .env.example .env` |
| API Key 未配置 | **停下来**，请用户自己编辑 `.env` 填入 `PPT_KING_API_KEY`；不要向用户索要 Key，也不要把 Key 写进聊天记录 |
| `资料/` 不存在或为空 | 请用户把资料放进 `资料/`，或询问资料所在路径 |
| 提示“可以开始” | 进入下一步 |

### 3. 直接生成

```bash
uv run ppt-king generate
```

默认读取 `资料/`，输出到 `out/资料/资料.pptx`。

如果用户指定了其他资料目录：

```bash
uv run ppt-king generate ./某个目录
```

如果用户明确要求页数、主题色、无 Logo 等，再追加对应参数：

```bash
uv run ppt-king generate --slides 12-18 --accent 2E5C8A
```

用户没有要求时，不要自作主张加参数。

### 4. 处理失败

- 单页失败时，程序会写入占位内容并返回非零退出码。先看终端里的失败页编号与原因。
- 优先复用已有结果重跑：

```bash
uv run ppt-king generate --resume
```

- 最多自动重试 2 次。仍失败时，向用户说明具体原因（资料不足、PDF 是扫描件、API 超时等）。
- **不要手改 `out/**/pages/*.md` 或 `slides.grid.md` 来掩盖问题**，除非用户明确要求手工修。
- 不要删除用户的资料文件。

### 5. 汇报结果

成功后告诉用户：

- PPTX 路径（`out/资料/资料.pptx`）
- 页数
- 是否有失败页或缺失图片
- 中间产物位置（`outline.md`、`slides.grid.md`、`pages/`），方便用户微调

## 安全与边界

- **不要读取、打印或提交 `.env` 内容**；只确认 Key 是否存在。
- 不要执行 `git add .env`、`git commit`、`git push`，除非用户明确要求。
- 不要把用户资料上传到配置的 LLM API 之外的任何服务。
- 不要修改 `资料/` 中的原始文件。
- 不要声称能做 OCR：扫描版 PDF 没有文字层时，请用户先 OCR。

## 常用命令

```bash
uv sync                          # 安装依赖
uv run ppt-king doctor           # 检查环境
uv run ppt-king generate         # 默认 资料/ -> out/资料/资料.pptx
uv run ppt-king generate --resume
uv run ppt-king render x.md      # 只渲染网格 Markdown
uv run pytest                    # 跑测试
```

## 项目结构速查

- [`documents.py`](src/ppt_king/documents.py)：多格式资料 → Markdown + 图片
- [`sources.py`](src/ppt_king/sources.py)：扫描资料、落盘图片、发现大纲
- [`outline.py`](src/ppt_king/outline.py)：解析或生成大纲
- [`writing.py`](src/ppt_king/writing.py)：逐页生成带坐标 Markdown
- [`grid_render.py`](src/ppt_king/grid_render.py)：确定性渲染 PPTX
- [`static/prompt.md`](static/prompt.md)：页面风格偏好
- [`static/templates/`](static/templates)：页面模板库
