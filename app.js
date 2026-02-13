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

async function loadPDF(file) {
    // TODO: 实现
}

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
        `;

        // 将canvas内容复制过去
        const previewCanvas = previewDiv.querySelector('canvas');
        previewCanvas.getContext('2d').drawImage(canvas, 0, 0);

        pagesContainer.appendChild(previewDiv);
    }
}

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

function showError(message) {
    alert(message);
}

function showSuccess(message) {
    console.log(message);
}
