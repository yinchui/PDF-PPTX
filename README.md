# PDF转PPTX（前端 + 本地高精度服务）

该项目提供两条转换链路：

1. 纯前端链路（无需服务）  
2. 本地高精度链路（`local_high_precision`，用于“文字可编辑 + 图片分离 + 图标级分离”）

不会上传到云端，所有处理均在本机完成。

## 功能概览

- PDF 拖拽/点击上传与多页预览
- 转换模式：
  - `fidelity`：视觉保真优先（整页图 + 文本）
  - `balanced`：平衡模式（文本不足自动图像兜底）
  - `editable`：仅文本对象（轻量可编辑）
  - `local_high_precision`：本地服务高精度分离（矢量优先、回退图片）
- 分段进度与阶段提示
- 调试日志开关
- 一键重置流程

## 快速开始（纯前端）

1. 直接打开 `index.html`
2. 选择 PDF
3. 选择模式（默认 `fidelity`）
4. 点击“转换为PPTX”

## 本地高精度模式（推荐用于图标分离）

### 1. 安装依赖

```powershell
cd backend
pip install -r requirements.txt
```

### 2. 启动本地服务

```powershell
cd backend
uvicorn main:app --host 127.0.0.1 --port 8000 --reload
```

### 3. 前端使用

1. 打开 `index.html`
2. 选择 `高精度分离（本地服务）`
3. 点击转换
4. 前端会轮询：
   - `POST /api/v1/jobs`
   - `GET /api/v1/jobs/{jobId}`
   - `GET /api/v1/jobs/{jobId}/download`
   - `GET /api/v1/jobs/{jobId}/report`

## 高精度模式输出策略

- 文字：写入可编辑文本框
- 图片：写入独立图片对象
- 图标：
  - 优先矢量写回（primitive / freeform）
  - 失败自动回退为独立图片对象
- 生成报告：记录每个图标是 `vector` 还是 `fallback_image`

## API 概览

### `POST /api/v1/jobs`

- `multipart/form-data`
  - `file`: PDF
  - `options`: JSON 字符串（可选）
- 返回：`{ "jobId": "..." }`

### `GET /api/v1/jobs/{jobId}`

- 返回任务状态、进度、阶段、指标、警告

### `GET /api/v1/jobs/{jobId}/download`

- 返回 `pptx` 文件

### `GET /api/v1/jobs/{jobId}/report`

- 返回分离报告（矢量成功数、回退数、失败原因等）

### `GET /api/v1/jobs/{jobId}/page-graph`

- 返回对象层调试图谱（文本/图片/矢量/图标候选）

## 已知限制

- 不是所有 PDF 图标都能 100% 转成 PPT 原生可编辑矢量
- 扫描件/扁平化 PDF 的矢量可编辑率会下降
- 复杂渐变、遮罩、裁剪路径会触发回退图片
- 大文件 + 高精度分离会更慢（精度优先策略）

## 故障排查

1. 前端提示无法连接本地服务  
   - 确认已启动：`uvicorn main:app --host 127.0.0.1 --port 8000`

2. 高精度模式失败  
   - 查看 `GET /api/v1/jobs/{jobId}` 返回的 `error`
   - 查看 `report` 的 `warnings`

3. 图标未矢量化  
   - 属于预期回退场景，检查 `report.icons[].result`
   - `fallback_image` 表示已分离但非矢量

## 验证清单

- [x] `node --check app.js`
- [x] `python -m py_compile backend/main.py backend/converter.py`
- [ ] 用当前 6 页样例执行 `local_high_precision` 并检查报告
- [ ] PowerPoint/WPS 打开后抽检文字编辑与图标分离
