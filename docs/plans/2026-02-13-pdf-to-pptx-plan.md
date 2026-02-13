# PDF转PPTX网页应用实施计划

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 创建一个纯前端的网页应用，将PDF格式的PPT转换为可编辑的PPTX文件

**Architecture:** 单页应用(SPA)，浏览器端完成所有PDF解析和PPTX生成，无需后端服务器

**Tech Stack:**
- HTML5 + CSS3 + Vanilla JavaScript
- pdf.js (PDF解析)
- pptxgenjs (PPTX生成)
- Tailwind CSS (样式)

---

## 任务概览

| 任务 | 描述 |
|------|------|
| Task 1 | 初始化项目结构 |
| Task 2 | 实现PDF上传组件 |
| Task 3 | 实现PDF预览渲染 |
| Task 4 | 实现PDF内容提取 |
| Task 5 | 实现PPTX生成器 |
| Task 6 | 实现UI交互和下载 |
| Task 7 | 测试和验证 |

---

### Task 1: 初始化项目结构

**Files:**
- Create: `index.html`
- Create: `styles.css`
- Create: `app.js`
- Create: `README.md`

**Step 1: 创建 index.html**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF转PPTX</title>
    <link rel="stylesheet" href="styles.css">
</head>
<body>
    <div class="container">
        <header>
            <h1>PDF转PPTX</h1>
            <p>将PDF格式的PPT转换为可编辑的PPTX文件</p>
        </header>

        <main>
            <div id="upload-area" class="upload-area">
                <div class="upload-content">
                    <svg class="upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                        <polyline points="17 8 12 3 7 8"></polyline>
                        <line x1="12" y1="3" x2="12" y2="15"></line>
                    </svg>
                    <p>拖拽PDF文件到这里，或点击选择</p>
                    <input type="file" id="file-input" accept=".pdf" hidden>
                    <button class="btn-primary" onclick="document.getElementById('file-input').click()">选择文件</button>
                </div>
            </div>

            <div id="preview-area" class="preview-area hidden">
                <h2>预览</h2>
                <div id="pages-container" class="pages-container"></div>

                <div class="action-bar">
                    <button id="convert-btn" class="btn-primary">转换为PPTX</button>
                    <div id="progress-container" class="progress-container hidden">
                        <div id="progress-bar" class="progress-bar"></div>
                        <span id="progress-text">0%</span>
                    </div>
                </div>

                <div id="download-area" class="download-area hidden">
                    <a id="download-btn" class="btn-success" download>下载PPTX</a>
                </div>
            </div>
        </main>

        <footer>
            <p>文件仅在浏览器中处理，不会上传到服务器</p>
        </footer>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>
    <script src="app.js"></script>
</body>
</html>
```

**Step 2: 创建 styles.css**

```css
:root {
    --primary: #4F46E5;
    --primary-hover: #4338CA;
    --success: #10B981;
    --bg: #F9FAFB;
    --card: #FFFFFF;
    --text: #1F2937;
    --text-light: #6B7280;
    --border: #E5E7EB;
}

* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
}

.container {
    max-width: 800px;
    margin: 0 auto;
    padding: 2rem;
}

header {
    text-align: center;
    margin-bottom: 2rem;
}

header h1 {
    font-size: 2rem;
    font-weight: 700;
    color: var(--text);
}

header p {
    color: var(--text-light);
    margin-top: 0.5rem;
}

.upload-area {
    background: var(--card);
    border: 2px dashed var(--border);
    border-radius: 1rem;
    padding: 3rem;
    text-align: center;
    transition: border-color 0.2s, background 0.2s;
    cursor: pointer;
}

.upload-area:hover, .upload-area.dragover {
    border-color: var(--primary);
    background: #EEF2FF;
}

.upload-content {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 1rem;
}

.upload-icon {
    width: 48px;
    height: 48px;
    color: var(--text-light);
}

.btn-primary {
    background: var(--primary);
    color: white;
    border: none;
    padding: 0.75rem 1.5rem;
    border-radius: 0.5rem;
    font-size: 1rem;
    cursor: pointer;
    transition: background 0.2s;
}

.btn-primary:hover {
    background: var(--primary-hover);
}

.btn-success {
    background: var(--success);
    color: white;
    border: none;
    padding: 0.75rem 1.5rem;
    border-radius: 0.5rem;
    font-size: 1rem;
    cursor: pointer;
    text-decoration: none;
    display: inline-block;
}

.btn-success:hover {
    background: #059669;
}

.hidden {
    display: none !important;
}

.preview-area {
    margin-top: 2rem;
}

.preview-area h2 {
    font-size: 1.25rem;
    margin-bottom: 1rem;
}

.pages-container {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
    gap: 1rem;
    margin-bottom: 1.5rem;
}

.page-preview {
    background: white;
    border: 1px solid var(--border);
    border-radius: 0.5rem;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}

.page-preview canvas {
    width: 100%;
    display: block;
}

.page-preview .page-number {
    text-align: center;
    padding: 0.5rem;
    background: var(--bg);
    font-size: 0.875rem;
    color: var(--text-light);
}

.action-bar {
    display: flex;
    align-items: center;
    gap: 1rem;
    flex-wrap: wrap;
}

.progress-container {
    flex: 1;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

.progress-bar {
    flex: 1;
    height: 8px;
    background: var(--border);
    border-radius: 4px;
    overflow: hidden;
}

.progress-bar::after {
    content: '';
    display: block;
    height: 100%;
    background: var(--primary);
    width: var(--progress, 0%);
    transition: width 0.3s;
}

.download-area {
    margin-top: 1.5rem;
    text-align: center;
}

footer {
    text-align: center;
    margin-top: 3rem;
    color: var(--text-light);
    font-size: 0.875rem;
}
```

**Step 3: 创建 app.js (骨架)**

```javascript
// PDF.js worker 配置
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const state = {
    pdfDoc: null,
    fileName: '',
    pages: []
};

// DOM 元素
const uploadArea = document.getElementById('upload-area');
const fileInput = document.getElementById('file-input');
const previewArea = document.getElementById('preview-area');
const pagesContainer = document.getElementById('pages-container');
const convertBtn = document.getElementById('convert-btn');
const progressContainer = document.getElementById('progress-container');
const progressBar = document.getElementById('progress-bar');
const progressText = document.getElementById('progress-text');
const downloadArea = document.getElementById('download-area');
const downloadBtn = document.getElementById('download-btn');

// 事件监听
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', handleDragOver);
uploadArea.addEventListener('dragleave', handleDragLeave);
uploadArea.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);
convertBtn.addEventListener('click', convertToPptx);

function handleDragOver(e) {
    e.preventDefault();
    uploadArea.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

async function processFile(file) {
    if (file.type !== 'application/pdf') {
        alert('请上传PDF文件');
        return;
    }

    state.fileName = file.name.replace('.pdf', '');

    // TODO: 加载PDF
    // TODO: 渲染预览
}

async function loadPDF(file) {
    // TODO: 实现
}

async function renderPreview() {
    // TODO: 实现
}

async function convertToPptx() {
    // TODO: 实现
}
```

**Step 4: 验证项目结构**

Run: `ls -la`
Expected: 看到 index.html, styles.css, app.js 文件

**Step 5: Commit**

```bash
git init
git add .
git commit -m "feat: 初始化项目结构"
```

---

### Task 2: 实现PDF上传和预览

**Files:**
- Modify: `app.js`

**Step 1: 更新 processFile 函数**

```javascript
async function processFile(file) {
    if (file.type !== 'application/pdf') {
        alert('请上传PDF文件');
        return;
    }

    state.fileName = file.name.replace('.pdf', '');

    try {
        const arrayBuffer = await file.arrayBuffer();
        state.pdfDoc = await pdfjsLib.getDocument(arrayBuffer).promise;

        renderPreview();
        uploadArea.classList.add('hidden');
        previewArea.classList.remove('hidden');
    } catch (error) {
        console.error('PDF加载失败:', error);
        alert('PDF文件加载失败，请确保文件未损坏');
    }
}
```

**Step 2: 更新 renderPreview 函数**

```javascript
async function renderPreview() {
    pagesContainer.innerHTML = '';
    state.pages = [];

    const totalPages = state.pdfDoc.numPages;

    for (let i = 1; i <= totalPages; i++) {
        const page = await state.pdfDoc.getPage(i);
        const viewport = page.getViewport({ scale: 0.5 });

        // 创建canvas
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        // 渲染页面
        await page.render({
            canvasContext: context,
            viewport: viewport
        }).promise;

        // 保存页面信息
        state.pages.push({
            pageNum: i,
            page: page,
            viewport: viewport
        });

        // 创建预览元素
        const previewDiv = document.createElement('div');
        previewDiv.className = 'page-preview';
        previewDiv.innerHTML = `
            <canvas width="${viewport.width}" height="${viewport.height}"></canvas>
            <div class="page-number">第 ${i} 页</div>
        `';

        // 将canvas内容复制过去
        const previewCanvas = previewDiv.querySelector('canvas');
        previewCanvas.getContext('2d').drawImage(canvas, 0, 0);

        pagesContainer.appendChild(previewDiv);
    }
}
```

**Step 3: 测试上传和预览**

1. 打开 index.html
2. 拖拽一个PDF文件到上传区域
3. 应该看到PDF页面预览

**Step 4: Commit**

```bash
git add app.js
git commit -feat: 实现PDF上传和预览功能
```

---

### Task 3: 实现PDF内容提取

**Files:**
- Modify: `app.js`

**Step 1: 添加 PDF 提取器**

```javascript
class PDFExtractor {
    constructor(pdfDoc) {
        this.pdfDoc = pdfDoc;
    }

    async extractPage(pageNum) {
        const page = await this.pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale: 1.0 });

        // 提取文本内容
        const textContent = await page.getTextContent();
        const texts = this.extractText(textContent, viewport);

        // 提取图片
        const images = await this.extractImages(page);

        // 提取背景色（简化版）
        const bgColor = this.estimateBackgroundColor(viewport);

        return {
            pageNum,
            width: viewport.width,
            height: viewport.height,
            texts,
            images,
            bgColor
        };
    }

    extractText(textContent, viewport) {
        const items = textContent.items;
        const texts = [];

        for (const item of items) {
            if (item.str.trim()) {
                // 转换坐标（PDF坐标需要转换）
                const tx = pdfjsLib.Util.transform(
                    viewport.transform,
                    item.transform
                );

                texts.push({
                    text: item.str,
                    x: tx[4],
                    y: viewport.height - tx[5], // 翻转Y坐标
                    fontSize: item.height || 12,
                    fontName: item.fontName || 'Arial',
                    width: item.width
                });
            }
        }

        return texts;
    }

    async extractImages(page) {
        const images = [];
        const ops = await page.getOperatorList();

        // 简化处理：暂不提取复杂图片
        // 实际生产中需要更复杂的处理

        return images;
    }

    estimateBackgroundColor(viewport) {
        // 默认白色背景
        return 'FFFFFF';
    }

    async extractAll() {
        const results = [];
        const total = this.pdfDoc.numPages;

        for (let i = 1; i <= total; i++) {
            const pageData = await this.extractPage(i);
            results.push(pageData);
        }

        return results;
    }
}
```

**Step 2: 更新 convertToPptx 函数使用提取器**

```javascript
async function convertToPptx() {
    convertBtn.disabled = true;
    progressContainer.classList.remove('hidden');
    downloadArea.classList.add('hidden');

    try {
        const extractor = new PDFExtractor(state.pdfDoc);

        const pageDataList = [];
        const total = state.pdfDoc.numPages;

        for (let i = 0; i < total; i++) {
            const progress = Math.round(((i + 1) / total) * 100);
            progressBar.style.setProperty('--progress', `${progress}%`);
            progressText.textContent = `${progress}%`;

            const pageData = await extractor.extractPage(i + 1);
            pageDataList.push(pageData);
        }

        // 生成PPTX
        const pptxData = generatePPTX(pageDataList);

        // 创建下载链接
        const blob = new Blob([pptxData], {
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        });
        const url = URL.createObjectURL(blob);

        downloadBtn.href = url;
        downloadBtn.download = `${state.fileName}.pptx`;
        downloadArea.classList.remove('hidden');

    } catch (error) {
        console.error('转换失败:', error);
        alert('转换失败，请重试');
    } finally {
        convertBtn.disabled = false;
        progressContainer.classList.add('hidden');
    }
}
```

**Step 3: Commit**

```bash
git add app.js
git commit -m "feat: 实现PDF内容提取器"
```

---

### Task 4: 实现PPTX生成器

**Files:**
- Modify: `app.js`

**Step 1: 添加 generatePPTX 函数**

```javascript
function generatePPTX(pageDataList) {
    const pptx = new PptxGenJS();

    // 设置幻灯片尺寸为A4横向
    pptx.layout = 'LAYOUT_16x9';

    for (const pageData of pageDataList) {
        const slide = pptx.addSlide();

        // 设置背景色
        if (pageData.bgColor && pageData.bgColor !== 'FFFFFF') {
            slide.background = { color: pageData.bgColor };
        }

        // 添加文字
        for (const text of pageData.texts) {
            // 计算字体大小（转换为磅）
            const fontSize = Math.max(text.fontSize * 0.75, 8);

            slide.addText(text.text, {
                x: text.x / pageData.width * 10,  // 转换为百分比
                y: text.y / pageData.height * 5.625,  // 16:9比例
                w: text.width ? (text.width / pageData.width * 10) : 4,
                h: (fontSize / 72) * 1.2,  // 转换为英寸
                fontSize: fontSize,
                fontFace: 'Arial',
                color: '363636'
            });
        }
    }

    return pptx.write({ blobType: 'blob' });
}
```

**Step 2: 测试转换功能**

1. 上传PDF
2. 点击"转换为PPTX"
3. 应该能看到进度条
4. 完成后显示下载按钮

**Step 3: Commit**

```bash
git add app.js
git commit -m "feat: 实现PPTX生成器"
```

---

### Task 5: 优化和完善

**Files:**
- Modify: `styles.css`
- Modify: `app.js`

**Step 1: 添加加载状态**

```css
.loading {
    opacity: 0.5;
    pointer-events: none;
}

.upload-area.loading::after {
    content: '加载中...';
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
}
```

**Step 2: 优化文字定位**

```javascript
function generatePPTX(pageDataList) {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    // PPTX的尺寸（英寸）
    const slideWidth = 10;
    const slideHeight = 5.625;

    for (const pageData of pageDataList) {
        const slide = pptx.addSlide();

        // 设置背景
        if (pageData.bgColor) {
            slide.background = { color: pageData.bgColor };
        }

        // 按Y坐标排序文字
        const sortedTexts = [...pageData.texts].sort((a, b) => b.y - a.y);

        for (const text of sortedTexts) {
            // 坐标转换：PDF坐标 -> PPTX坐标
            // PDF: 左下角为原点，向上为正
            // PPTX: 左上角为原点，向下为正

            const x = (text.x / pageData.width) * slideWidth;
            const y = slideHeight - (text.y / pageData.height) * slideHeight - 0.3;
            const fontSize = Math.max(text.fontSize * 0.75, 8);
            const w = text.width ? (text.width / pageData.width * slideWidth) : 4;

            slide.addText(text.text, {
                x: Math.max(0.1, x),
                y: Math.max(0.1, y),
                w: Math.min(w, slideWidth - x - 0.1),
                h: (fontSize / 72) * 1.2,
                fontSize: fontSize,
                fontFace: 'Arial',
                color: '363636'
            });
        }
    }

    return pptx.write({ blobType: 'blob' });
}
```

**Step 3: 添加错误处理UI**

```javascript
function showError(message) {
    alert(message);
}

function showSuccess(message) {
    // 可以用更优雅的方式
    console.log(message);
}
```

**Step 4: Commit**

```bash
git add .
git commit -m "feat: 优化转换逻辑和用户体验"
```

---

### Task 6: 最终测试和验证

**Files:**
- 测试所有功能

**Step 1: 测试完整流程**

1. 打开 index.html
2. 上传一个PDF文件
3. 确认预览显示正常
4. 点击转换
5. 下载PPTX
6. 用PowerPoint/WPS打开确认可编辑

**Step 2: 验证验收标准**

- [x] 可以上传PDF文件并显示预览
- [x] 可以将PDF转换为PPTX并下载
- [x] 转换后的PPTX可以用PowerPoint/WPS打开
- [x] 文字基本可编辑

**Step 3: Commit**

```bash
git add .
git commit -m "feat: 完成PDF转PPTX应用开发"
```

---

## 执行选项

**Plan complete and saved to `docs/plans/2026-02-13-pdf-to-pptx-plan.md`. Two execution options:**

**1. Subagent-Driven (this session)** - I dispatch fresh subagent per task, review between tasks, fast iteration

**2. Parallel Session (separate)** - Open new session with executing-plans, batch execution with checkpoints

**Which approach?**
