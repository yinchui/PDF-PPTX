# PDF 转 PPTX 本地高精度设计（2026-02-13）

## 1. 设计目标

- 文字与图片分离，文本可编辑
- 图标尽量逐个分离，矢量优先
- 不能矢量化时回退为独立图片并可追踪

## 2. 总体架构

### 前端
- 保留原有三模式
- 新增 `local_high_precision`
- 负责上传、轮询、下载、展示报告摘要

### 本地服务（FastAPI）
- Job API + 状态管理
- 对象层提取与写回
- 产出 `output.pptx`、`report.json`、`page_graph.json`

## 3. 核心数据模型

- `ExtractedText`: `text, bbox_pt, font_name, font_size_pt, color`
- `ExtractedImage`: `id, bbox_pt, mime, bytes`
- `VectorPath`: `id, bbox_pt, stroke, fill, width, items`
- `IconCandidate`: `id, bbox_pt, paths[], classify_result`
- `ConversionReport`:
  - `vector_icons_ok`
  - `vector_icons_fallback`
  - `text_count`
  - `image_count`
  - `warnings`
  - `icons[]`

## 4. 图标分离策略

1. 来源：`page.get_drawings()`  
2. 过滤：
- 面积占比 > 35% 的路径当背景过滤
- 聚类后尺寸必须在 8pt~220pt
3. 聚类：
- bbox 邻近阈值 6pt，连通域聚类
4. 写回：
- primitive 优先（矩形等）
- 复杂路径离散化后 freeform
- 异常时回退 clip 栅格图

## 5. 坐标映射

- 输入坐标：PDF pt（PyMuPDF 页面坐标）
- 输出坐标：PPT inch（13.333 x 7.5）
- 统一使用 bbox 比例映射，避免局部硬编码

## 6. 可观测性与容错

- Job 状态：`queued/running/done/failed`
- 进度与阶段文本实时返回
- 错误信息和 traceback 存在 job 状态
- 报告记录每个 icon 结果，不允许静默失败

## 7. 已知限制

- 扫描件/扁平化 PDF 的图标矢量化能力有限
- 复杂渐变、遮罩、裁剪路径多回退为图片
- “矢量可编辑率”依赖源 PDF 的对象质量
