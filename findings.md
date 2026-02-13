# Findings & Decisions

## Requirements
- 在现有前端模式外新增 `local_high_precision` 模式
- 文本必须可编辑、图片必须独立对象
- 图标要“矢量优先”，失败时回退独立图片
- 输出可追踪报告（每个图标的处理结果）

## Research Findings
- 当前前端链路 `extractImages()` 为占位实现，无法完成图标级分离。
- 默认 `fidelity` 会铺整页图，虽保真但会弱化可编辑体验。
- 本机具备 `Python 3.10`、`fastapi`、`uvicorn`、`python-pptx`，适合引入本地服务。
- `python-pptx` 的 freeform 能力可用（`build_freeform` 存在），可用于路径离散化写回。
- 转换器烟测（临时 PDF）可生成有效 PPTX，并统计到 `vector_icons_ok`。

## Technical Decisions
| Decision | Rationale |
|----------|-----------|
| 新增 `backend/` FastAPI job API | 支持高精度长任务与进度回传 |
| 对象提取使用 PyMuPDF | 可直接读取文本/图像/绘图对象层 |
| 图标聚类策略：邻近 + 尺寸过滤 | 避免把大背景误判为图标 |
| 矢量优先，失败回退图片 | 满足“可用优先 + 可追踪” |
| 前端模式分流 | 不破坏现有纯前端链路 |

## Issues Encountered
| Issue | Resolution |
|-------|------------|
| PowerShell 不支持 bash 风格 here-doc (`<<`) | 改用 `python -c` |
| PowerShell 字符串中 `$var:` 解析失败 | 使用 `${var}` 形式 |
| `fastapi.testclient` 在当前依赖组合下初始化报 `Client.__init__` 参数错误 | 改用转换器级烟测与模块导入验证 |

## Resources
- `backend/main.py`
- `backend/converter.py`
- `backend/requirements.txt`
- `app.js`
- `index.html`
- `README.md`

## Visual/Browser Findings
- 现有 UI 可直接扩展模式选项，无需大改布局。
- 前端已有进度和消息区域，适配 job 轮询成本低。
