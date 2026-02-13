# Task Plan: 本地高精度“文字可编辑 + 图片分离 + 图标级分离”

## Goal
在保留纯前端转换链路的基础上，新增本地高精度服务模式，实现文字可编辑、图片独立对象、图标矢量优先分离并支持回退与报告追踪。

## Current Phase
Phase 6 (Manual Acceptance Pending)

## Phases
### Phase 1: Baseline Check
- [x] 审查现有前端链路与已实现模式
- [x] 确认新增本地后端不会破坏现有 `fidelity/balanced/editable`
- **Status:** complete

### Phase 2: Backend Service Skeleton
- [x] 新增 `backend/` 目录
- [x] 提供 FastAPI + job 生命周期 API
- [x] 提供下载、报告、page-graph 接口
- **Status:** complete

### Phase 3: Object Extraction Layer
- [x] 提取文本对象（span + bbox + 字体）
- [x] 提取图片对象（xref + bbox）
- [x] 提取矢量路径（drawings + style）
- [x] 图标候选聚类（邻近 + 尺寸过滤）
- **Status:** complete

### Phase 4: PPTX Write-back Layer
- [x] 文本写回可编辑文本框
- [x] 图片写回独立对象
- [x] 图标矢量优先写回（primitive/freeform）
- [x] 失败回退图片并记录原因
- **Status:** complete

### Phase 5: Frontend Integration
- [x] 新增 `local_high_precision` 模式入口
- [x] 接入创建任务、轮询状态、下载与报告
- [x] 增加本地服务不可达提示
- **Status:** complete

### Phase 6: Docs & Verification
- [x] 更新 README（本地服务启动与接口）
- [x] 执行语法/编译检查
- [ ] 用真实样例完成端到端人工验收
- **Status:** in_progress

## Key Questions
1. 图标是否必须全矢量？  
已锁定：分层验收，矢量优先，失败回退图片并可追踪。
2. 如何保证本地模式不影响原有模式？  
已锁定：前端按模式分流，现有模式仍走原链路。
3. 如何让失败可追踪？  
已锁定：`report.json` 记录每个 icon 的处理结果与原因。

## Decisions Made
| Decision | Rationale |
|----------|-----------|
| 新增本地服务而非替换前端链路 | 兼顾可用性与高精度能力 |
| 图标处理采用“矢量优先 + 回退图片” | 在精度与稳定性之间可交付 |
| API 使用 job 模式 | 便于轮询进度和错误可观测 |
| 输出 report/page-graph | 支持调参与问题追踪 |

## Errors Encountered
| Error | Attempt | Resolution |
|-------|---------|------------|
| PowerShell here-doc 语法不兼容 | 1 | 改用 `python -c` 探测依赖 |
| PowerShell 变量插值包含 `:` 报错 | 1 | 使用 `${var}` 包裹变量名 |
| `fastapi.testclient` 在当前依赖组合下初始化异常 | 1 | 用模块导入 + 转换器烟测替代接口内嵌测试 |

## Notes
- 剩余工作是人工端到端验收，不再是代码实现阻塞。
