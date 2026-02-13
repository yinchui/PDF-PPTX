# PDF 转 PPTX 本地高精度实施计划（2026-02-13）

## 目标

在保留前端三模式（`fidelity` / `balanced` / `editable`）的前提下，新增 `local_high_precision`：  
实现文字可编辑、图片独立对象、图标矢量优先分离并输出追踪报告。

## 范围

### In Scope
- 本地 FastAPI 服务（job 生命周期）
- PyMuPDF 对象层提取（文本/图像/矢量）
- python-pptx 写回（文本、图片、矢量/回退）
- 前端本地模式接入与轮询下载
- 报告输出（`report.json` / `page_graph.json`）

### Out of Scope
- 云端处理
- 承诺 100% 图标矢量化
- 旧浏览器兼容增强

## 里程碑

### M1：Backend Skeleton
- `POST /api/v1/jobs`
- `GET /api/v1/jobs/{jobId}`
- `GET /api/v1/jobs/{jobId}/download`
- `GET /api/v1/jobs/{jobId}/report`
- `GET /api/v1/jobs/{jobId}/page-graph`

### M2：Extraction + Write-back
- 提取文本、图片、矢量路径
- 图标候选聚类（邻近 + 尺寸过滤）
- 矢量优先写回（primitive/freeform）
- 失败回退图片并记录原因

### M3：Frontend Integration
- 模式新增 `local_high_precision`
- 创建任务 + 轮询状态 + 下载结果 + 报告提示

### M4：Docs + Validation
- README 增加本地服务启动和 API 说明
- 执行静态检查（JS/Python）

## 验收标准

1. PPTX 可被 PowerPoint/WPS 打开，页数一致。  
2. 文本可编辑、图片可独立选中。  
3. 图标对象结果可追踪（`vector` / `fallback_image`）。  
4. 转换失败不静默，错误可通过 job 状态或 report 定位。  

## 当前状态

- 代码实现：完成
- 静态检查：完成
- 端到端人工样例验收：待执行
