// PDF.js worker configuration
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const SLIDE_LAYOUT_SIZES = {
    LAYOUT_STANDARD: { width: 10, height: 7.5 },
    LAYOUT_WIDE: { width: 13.333, height: 7.5 }
};

const defaultOptions = {
    mode: 'fidelity',
    slideLayout: 'LAYOUT_WIDE',
    imageScale: 1.75,
    debug: false
};

const FONT_SIZE_MIN = 8;
const FONT_SIZE_MAX = 72;
const LINE_TOLERANCE = 2;
const FONT_TOLERANCE = 1.5;
const MERGE_GAP = 8;
const BALANCED_TEXT_THRESHOLD = 12;
const LOCAL_API_BASE = 'http://127.0.0.1:8000';
const LOCAL_POLL_INTERVAL_MS = 1000;
const LOCAL_TIMEOUT_MS = 20 * 60 * 1000;

/**
 * @typedef {'fidelity'|'editable'|'balanced'|'local_high_precision'} ConvertMode
 */

/**
 * @typedef {Object} ConvertOptions
 * @property {ConvertMode} mode
 * @property {'LAYOUT_STANDARD'|'LAYOUT_WIDE'} slideLayout
 * @property {number} imageScale
 * @property {boolean} debug
 */

/**
 * @typedef {Object} TextRun
 * @property {string} text
 * @property {number} x
 * @property {number} y
 * @property {number} w
 * @property {number} h
 * @property {number} fontSize
 * @property {string} fontFace
 */

/**
 * @typedef {Object} PageData
 * @property {number} pageNum
 * @property {number} width
 * @property {number} height
 * @property {TextRun[]} texts
 * @property {Array<Object>} images
 * @property {{dataUrl: string, width: number, height: number}|null} pageRaster
 * @property {string} bgColor
 */

const state = {
    pdfDoc: null,
    sourceFile: null,
    fileName: '',
    pages: [],
    downloadUrl: null,
    activeJobId: null,
    isBusy: false,
    stats: {
        skippedTextRuns: 0,
        mergedTextRuns: 0,
        rasterPages: 0
    }
};

// DOM references
const uploadArea = document.getElementById('upload-area');
const chooseFileBtn = document.getElementById('choose-file-btn');
const fileInput = document.getElementById('file-input');
const previewArea = document.getElementById('preview-area');
const pagesContainer = document.getElementById('pages-container');
const modeSelect = document.getElementById('mode-select');
const debugToggle = document.getElementById('debug-toggle');
const convertBtn = document.getElementById('convert-btn');
const resetBtn = document.getElementById('reset-btn');
const progressContainer = document.getElementById('progress-container');
const progressBar = document.getElementById('progress-bar');
const progressText = document.getElementById('progress-text');
const progressStage = document.getElementById('progress-stage');
const downloadArea = document.getElementById('download-area');
const downloadBtn = document.getElementById('download-btn');
const messageArea = document.getElementById('message-area');

init();

function init() {
    uploadArea.addEventListener('click', onUploadAreaClick);
    chooseFileBtn.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);
    convertBtn.addEventListener('click', () => convertToPptx(getConvertOptions()));
    resetBtn.addEventListener('click', handleReset);
    window.addEventListener('beforeunload', cleanupDownloadUrl);

    modeSelect.value = defaultOptions.mode;
    convertBtn.disabled = true;
    setProgress({ extract: 0, generate: 0, phase: '' });
}

function onUploadAreaClick(event) {
    if (state.isBusy) {
        return;
    }

    const target = event.target;
    if (target instanceof HTMLButtonElement || target instanceof HTMLInputElement) {
        return;
    }

    fileInput.click();
}

function handleDragOver(event) {
    event.preventDefault();
    if (!state.isBusy) {
        uploadArea.classList.add('dragover');
    }
}

function handleDragLeave(event) {
    event.preventDefault();
    uploadArea.classList.remove('dragover');
}

function handleDrop(event) {
    event.preventDefault();
    uploadArea.classList.remove('dragover');

    if (state.isBusy) {
        return;
    }

    const files = event.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(event) {
    const files = event.target.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleReset() {
    if (state.isBusy) {
        return;
    }

    resetAppState({ clearMessage: false });
    showSuccess('已重置，可重新选择 PDF 文件。');
}

function getConvertOptions() {
    const mode = normalizeMode(modeSelect.value);
    const debug = Boolean(debugToggle.checked);
    const imageScale = mode === 'fidelity' ? 2 : mode === 'balanced' ? 1.75 : mode === 'editable' ? 1.25 : 2;

    return {
        ...defaultOptions,
        mode,
        imageScale,
        debug
    };
}

function normalizeMode(mode) {
    if (mode === 'editable' || mode === 'balanced' || mode === 'fidelity' || mode === 'local_high_precision') {
        return mode;
    }
    return defaultOptions.mode;
}

function isPdfFile(file) {
    const byMime = file.type === 'application/pdf';
    const byExtension = /\.pdf$/i.test(file.name || '');
    return byMime || byExtension;
}

function normalizeFileName(file) {
    const rawName = (file && file.name ? file.name : 'converted').trim();
    const noExtension = rawName.replace(/\.pdf$/i, '');
    return noExtension || 'converted';
}

function setBusy(isBusy) {
    state.isBusy = isBusy;

    uploadArea.classList.toggle('loading', isBusy);
    previewArea.classList.toggle('loading', isBusy);
    fileInput.disabled = isBusy;
    modeSelect.disabled = isBusy;
    debugToggle.disabled = isBusy;
    chooseFileBtn.disabled = isBusy;
    resetBtn.disabled = isBusy;
    convertBtn.disabled = isBusy || !state.pdfDoc;
}

function setProgress({ extract = 0, generate = 0, phase = '' }) {
    const safeExtract = clamp(extract, 0, 100);
    const safeGenerate = clamp(generate, 0, 100);
    const total = Math.round((safeExtract * 0.65) + (safeGenerate * 0.35));

    progressBar.style.setProperty('--progress', `${total}%`);
    progressText.textContent = `${total}%（提取 ${safeExtract}% / 生成 ${safeGenerate}%）`;
    progressStage.textContent = phase;
}

function showMessage(message, type) {
    messageArea.className = `message message-${type}`;
    messageArea.textContent = message;
    messageArea.classList.remove('hidden');
}

function clearMessage() {
    messageArea.textContent = '';
    messageArea.className = 'message hidden';
}

function showError(message, error = null) {
    if (error) {
        console.error(message, error);
    } else {
        console.error(message);
    }
    showMessage(message, 'error');
}

function showSuccess(message) {
    console.info(message);
    showMessage(message, 'success');
}

function cleanupDownloadUrl() {
    if (state.downloadUrl) {
        URL.revokeObjectURL(state.downloadUrl);
        state.downloadUrl = null;
    }

    downloadBtn.removeAttribute('href');
}

function resetAppState({ clearMessage: shouldClearMessage = true } = {}) {
    cleanupDownloadUrl();

    state.pdfDoc = null;
    state.sourceFile = null;
    state.fileName = '';
    state.pages = [];
    state.activeJobId = null;
    state.stats = {
        skippedTextRuns: 0,
        mergedTextRuns: 0,
        rasterPages: 0
    };

    pagesContainer.innerHTML = '';
    progressContainer.classList.add('hidden');
    downloadArea.classList.add('hidden');
    previewArea.classList.add('hidden');
    uploadArea.classList.remove('hidden');

    fileInput.value = '';
    setProgress({ extract: 0, generate: 0, phase: '' });
    setBusy(false);

    if (shouldClearMessage) {
        clearMessage();
    }
}

async function processFile(file) {
    clearMessage();

    if (!isPdfFile(file)) {
        showError('请上传 PDF 文件（支持 MIME 或 .pdf 扩展名识别）。');
        return;
    }

    cleanupDownloadUrl();
    downloadArea.classList.add('hidden');
    progressContainer.classList.add('hidden');

    state.fileName = normalizeFileName(file);
    state.sourceFile = file;

    setBusy(true);

    try {
        const arrayBuffer = await file.arrayBuffer();
        state.pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

        await renderPreview();

        uploadArea.classList.add('hidden');
        previewArea.classList.remove('hidden');
        showSuccess(`PDF 加载成功，共 ${state.pdfDoc.numPages} 页。`);
    } catch (error) {
        resetAppState({ clearMessage: false });
        showError('PDF 文件加载失败，请确认文件未损坏。', error);
    } finally {
        setBusy(false);
    }
}

async function renderPreview() {
    if (!state.pdfDoc) {
        return;
    }

    pagesContainer.innerHTML = '';
    state.pages = [];

    const totalPages = state.pdfDoc.numPages;

    for (let pageNum = 1; pageNum <= totalPages; pageNum += 1) {
        const page = await state.pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale: 0.45 });

        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;

        const context = canvas.getContext('2d', { alpha: false });
        await page.render({ canvasContext: context, viewport }).promise;

        const previewDiv = document.createElement('div');
        previewDiv.className = 'page-preview';

        const previewCanvas = document.createElement('canvas');
        previewCanvas.width = viewport.width;
        previewCanvas.height = viewport.height;
        previewCanvas.getContext('2d').drawImage(canvas, 0, 0);

        const pageNumber = document.createElement('div');
        pageNumber.className = 'page-number';
        pageNumber.textContent = `第 ${pageNum} 页`;

        previewDiv.appendChild(previewCanvas);
        previewDiv.appendChild(pageNumber);
        pagesContainer.appendChild(previewDiv);

        state.pages.push({ pageNum, width: viewport.width, height: viewport.height });
    }
}

async function convertToPptx(options = defaultOptions) {
    if (!state.pdfDoc) {
        showError('请先上传并解析 PDF 文件。');
        return;
    }

    const convertOptions = {
        ...defaultOptions,
        ...options,
        mode: normalizeMode(options.mode || defaultOptions.mode)
    };

    clearMessage();
    cleanupDownloadUrl();
    downloadArea.classList.add('hidden');
    progressContainer.classList.remove('hidden');
    setProgress({ extract: 0, generate: 0, phase: '准备开始...' });
    setBusy(true);

    try {
        if (convertOptions.mode === 'local_high_precision') {
            await convertViaLocalService(convertOptions);
            return;
        }

        const extractor = new PDFExtractor(state.pdfDoc);
        const pageDataList = [];
        const totalPages = state.pdfDoc.numPages;

        for (let pageNum = 1; pageNum <= totalPages; pageNum += 1) {
            const pageData = await extractor.extractPage(pageNum, convertOptions);
            pageDataList.push(pageData);

            const extractProgress = Math.round((pageNum / totalPages) * 100);
            setProgress({
                extract: extractProgress,
                generate: 0,
                phase: `提取内容中（${pageNum}/${totalPages}）`
            });
        }

        const pptxBlob = await generatePPTX(pageDataList, convertOptions, (generateProgress) => {
            setProgress({
                extract: 100,
                generate: generateProgress,
                phase: `生成 PPTX 中（${generateProgress}%）`
            });
        });

        const url = URL.createObjectURL(pptxBlob);
        state.downloadUrl = url;

        downloadBtn.href = url;
        downloadBtn.download = `${state.fileName || 'converted'}.pptx`;
        downloadArea.classList.remove('hidden');

        state.stats = {
            skippedTextRuns: extractor.stats.skippedTextRuns,
            mergedTextRuns: extractor.stats.mergedTextRuns,
            rasterPages: extractor.stats.rasterPages
        };

        if (convertOptions.debug) {
            logDebugInfo(convertOptions, pageDataList, extractor.stats);
        }

        showSuccess(`转换完成，共导出 ${pageDataList.length} 页。`);
    } catch (error) {
        showError('转换失败，请重试。', error);
    } finally {
        state.activeJobId = null;
        setBusy(false);
        progressContainer.classList.add('hidden');
    }
}

async function convertViaLocalService(options) {
    if (!state.sourceFile) {
        throw new Error('未找到源 PDF 文件，请重新选择文件后重试。');
    }

    setProgress({ extract: 0, generate: 0, phase: '提交本地高精度任务...' });
    const jobId = await createLocalJob(options);
    state.activeJobId = jobId;

    const jobStatus = await pollLocalJob(jobId);
    const pptxBlob = await fetchLocalJobDownload(jobId);
    const report = await fetchLocalJobReport(jobId);

    const url = URL.createObjectURL(pptxBlob);
    state.downloadUrl = url;

    downloadBtn.href = url;
    downloadBtn.download = `${state.fileName || 'converted'}.pptx`;
    downloadArea.classList.remove('hidden');

    if (report) {
        state.stats = {
            skippedTextRuns: Number(report.text_count || 0),
            mergedTextRuns: Number(report.vector_icons_ok || 0),
            rasterPages: Number(report.vector_icons_fallback || 0)
        };
    }

    if (options.debug) {
        console.group('[Local High Precision Debug]');
        console.table({
            jobId,
            status: jobStatus.status,
            progress: jobStatus.progress,
            vector_icons_ok: report?.vector_icons_ok ?? 0,
            vector_icons_fallback: report?.vector_icons_fallback ?? 0,
            text_count: report?.text_count ?? 0,
            image_count: report?.image_count ?? 0
        });
        if (Array.isArray(report?.warnings) && report.warnings.length > 0) {
            console.warn('warnings:', report.warnings);
        }
        console.groupEnd();
    }

    showSuccess(formatLocalReportMessage(report));
}

async function createLocalJob(options) {
    const endpoint = `${LOCAL_API_BASE}/api/v1/jobs`;
    const formData = new FormData();
    formData.append('file', state.sourceFile, state.sourceFile.name);
    formData.append(
        'options',
        JSON.stringify({
            mode: 'local_high_precision',
            vector_tolerance_pt: 0.6,
            cluster_gap_pt: 6.0,
            background_filter_ratio: 0.35,
            min_icon_size_pt: 8.0,
            max_icon_size_pt: 220.0,
            debug: Boolean(options.debug)
        })
    );

    let response;
    try {
        response = await fetch(endpoint, {
            method: 'POST',
            body: formData
        });
    } catch (error) {
        throw new Error('无法连接本地服务，请先启动 backend 服务（127.0.0.1:8000）。');
    }

    if (!response.ok) {
        throw new Error(await parseErrorResponse(response, '创建本地任务失败'));
    }

    const payload = await response.json();
    if (!payload.jobId) {
        throw new Error('本地服务返回了无效任务 ID。');
    }
    return payload.jobId;
}

async function pollLocalJob(jobId) {
    const startTime = Date.now();

    while (true) {
        if (Date.now() - startTime > LOCAL_TIMEOUT_MS) {
            throw new Error('本地高精度任务超时，请稍后重试。');
        }

        const status = await fetchLocalJobStatus(jobId);
        const progress = clamp(Number(status.progress) || 0, 0, 100);
        setProgress({
            extract: progress,
            generate: progress,
            phase: `本地服务：${status.stage || status.status || '处理中'}`
        });

        if (status.status === 'done') {
            return status;
        }
        if (status.status === 'failed') {
            const detail = status.error ? `：${status.error}` : '';
            throw new Error(`本地服务任务失败${detail}`);
        }

        await sleep(LOCAL_POLL_INTERVAL_MS);
    }
}

async function fetchLocalJobStatus(jobId) {
    const endpoint = `${LOCAL_API_BASE}/api/v1/jobs/${encodeURIComponent(jobId)}`;
    let response;

    try {
        response = await fetch(endpoint);
    } catch (error) {
        throw new Error('查询本地服务状态失败，请确认服务仍在运行。');
    }

    if (!response.ok) {
        throw new Error(await parseErrorResponse(response, '查询任务状态失败'));
    }
    return response.json();
}

async function fetchLocalJobDownload(jobId) {
    const endpoint = `${LOCAL_API_BASE}/api/v1/jobs/${encodeURIComponent(jobId)}/download`;
    let response;

    try {
        response = await fetch(endpoint);
    } catch (error) {
        throw new Error('下载本地服务生成结果失败。');
    }

    if (!response.ok) {
        throw new Error(await parseErrorResponse(response, '下载结果失败'));
    }
    return response.blob();
}

async function fetchLocalJobReport(jobId) {
    const endpoint = `${LOCAL_API_BASE}/api/v1/jobs/${encodeURIComponent(jobId)}/report`;
    let response;
    try {
        response = await fetch(endpoint);
    } catch (error) {
        return null;
    }
    if (!response.ok) {
        return null;
    }
    return response.json();
}

function formatLocalReportMessage(report) {
    if (!report) {
        return '本地高精度转换完成。';
    }

    const vectorOk = Number(report.vector_icons_ok || 0);
    const fallback = Number(report.vector_icons_fallback || 0);
    const textCount = Number(report.text_count || 0);
    const imageCount = Number(report.image_count || 0);
    return `本地高精度转换完成：文本 ${textCount}，图片 ${imageCount}，图标矢量 ${vectorOk}，回退图片 ${fallback}。`;
}

async function parseErrorResponse(response, fallbackMessage) {
    try {
        const payload = await response.json();
        const detail = typeof payload.detail === 'string' ? payload.detail : JSON.stringify(payload.detail || payload);
        return `${fallbackMessage}：${detail}`;
    } catch (error) {
        return `${fallbackMessage}（HTTP ${response.status}）`;
    }
}

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

class PDFExtractor {
    constructor(pdfDoc) {
        this.pdfDoc = pdfDoc;
        this.stats = {
            skippedTextRuns: 0,
            mergedTextRuns: 0,
            rasterPages: 0
        };
    }

    /**
     * @param {number} pageNum
     * @param {ConvertOptions} options
     * @returns {Promise<PageData>}
     */
    async extractPage(pageNum, options) {
        const page = await this.pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale: 1.0 });

        const textContent = await page.getTextContent();
        const rawTextRuns = this.extractText(textContent, viewport);
        const sortedTextRuns = sortTextRuns(rawTextRuns);
        const mergedTextRuns = mergeAdjacentTextRuns(sortedTextRuns);

        this.stats.mergedTextRuns += Math.max(0, sortedTextRuns.length - mergedTextRuns.length);

        const images = await this.extractImages(page);

        let pageRaster = null;
        if (options.mode !== 'editable') {
            pageRaster = await this.renderPageRaster(page, options.imageScale);
            if (pageRaster) {
                this.stats.rasterPages += 1;
            }
        }

        const bgColor = this.estimateBackgroundColor();

        return {
            pageNum,
            width: viewport.width,
            height: viewport.height,
            texts: mergedTextRuns,
            images,
            pageRaster,
            bgColor
        };
    }

    extractText(textContent, viewport) {
        const texts = [];

        for (const item of textContent.items) {
            if (!item || typeof item.str !== 'string') {
                this.stats.skippedTextRuns += 1;
                continue;
            }

            const text = item.str.replace(/\s+/g, ' ').trim();
            if (!text) {
                this.stats.skippedTextRuns += 1;
                continue;
            }

            if (!Array.isArray(item.transform) || item.transform.length < 6) {
                this.stats.skippedTextRuns += 1;
                continue;
            }

            const tx = pdfjsLib.Util.transform(viewport.transform, item.transform);
            const x = Number(tx[4]);
            const y = viewport.height - Number(tx[5]);

            if (!Number.isFinite(x) || !Number.isFinite(y)) {
                this.stats.skippedTextRuns += 1;
                continue;
            }

            const fontSize = clamp(Number(item.height) || 12, FONT_SIZE_MIN, FONT_SIZE_MAX);
            const width = Math.max(0, Number(item.width) || 0);

            texts.push({
                text,
                x,
                y,
                w: width,
                h: fontSize,
                fontSize,
                fontFace: normalizeFontName(item.fontName)
            });
        }

        return texts;
    }

    async extractImages(page) {
        // Placeholder for future deep extraction logic.
        // We keep this interface to preserve extension compatibility.
        void page;
        return [];
    }

    async renderPageRaster(page, imageScale) {
        const safeScale = clamp(Number(imageScale) || defaultOptions.imageScale, 1, 3);
        const viewport = page.getViewport({ scale: safeScale });

        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;

        const context = canvas.getContext('2d', { alpha: false });
        await page.render({ canvasContext: context, viewport }).promise;

        const dataUrl = canvas.toDataURL('image/png');

        canvas.width = 1;
        canvas.height = 1;

        return {
            dataUrl,
            width: viewport.width,
            height: viewport.height
        };
    }

    estimateBackgroundColor() {
        return 'FFFFFF';
    }
}

function sortTextRuns(textRuns) {
    return [...textRuns].sort((left, right) => {
        const lineDiff = Math.abs(left.y - right.y);
        if (lineDiff <= LINE_TOLERANCE) {
            return left.x - right.x;
        }
        return right.y - left.y;
    });
}

function mergeAdjacentTextRuns(textRuns) {
    if (textRuns.length === 0) {
        return [];
    }

    const merged = [];
    let current = { ...textRuns[0] };

    for (let index = 1; index < textRuns.length; index += 1) {
        const next = textRuns[index];
        const sameLine = Math.abs(current.y - next.y) <= LINE_TOLERANCE;
        const closeFontSize = Math.abs(current.fontSize - next.fontSize) <= FONT_TOLERANCE;
        const currentRight = current.x + Math.max(current.w, 0);
        const gap = next.x - currentRight;
        const shouldMerge = sameLine && closeFontSize && gap >= -1 && gap <= MERGE_GAP;

        if (!shouldMerge) {
            merged.push(current);
            current = { ...next };
            continue;
        }

        const needsSpace = shouldInsertSpace(current.text, next.text, gap);
        current.text = `${current.text}${needsSpace ? ' ' : ''}${next.text}`;
        current.w = Math.max(next.x + next.w - current.x, current.w);
        current.h = Math.max(current.h, next.h);
        current.fontSize = Math.max(current.fontSize, next.fontSize);
    }

    merged.push(current);
    return merged;
}

function shouldInsertSpace(currentText, nextText, gap) {
    if (gap <= 0.2) {
        return false;
    }

    if (!currentText || !nextText) {
        return false;
    }

    const lastChar = currentText.slice(-1);
    const firstChar = nextText.charAt(0);

    if (/\s/.test(lastChar) || /\s/.test(firstChar)) {
        return false;
    }

    if (/[,.;:!?)]/.test(firstChar)) {
        return false;
    }

    return true;
}

/**
 * @param {PageData[]} pageDataList
 * @param {ConvertOptions} options
 * @param {(progress:number)=>void} onProgress
 * @returns {Promise<Blob>}
 */
async function generatePPTX(pageDataList, options, onProgress = () => {}) {
    const pptx = new PptxGenJS();
    const layout = options.slideLayout in SLIDE_LAYOUT_SIZES ? options.slideLayout : defaultOptions.slideLayout;

    pptx.layout = layout;

    const slideSize = SLIDE_LAYOUT_SIZES[layout];
    const totalPages = pageDataList.length;

    for (let index = 0; index < totalPages; index += 1) {
        const pageData = pageDataList[index];
        const slide = pptx.addSlide();

        if (pageData.bgColor) {
            slide.background = { color: pageData.bgColor };
        }

        if (shouldUseRasterFallback(pageData, options)) {
            slide.addImage({
                data: pageData.pageRaster.dataUrl,
                x: 0,
                y: 0,
                w: slideSize.width,
                h: slideSize.height
            });
        }

        for (const textRun of pageData.texts) {
            const safeText = String(textRun.text || '').trim();
            if (!safeText) {
                continue;
            }

            const rect = pdfRectToSlideRect(textRun, pageData, slideSize);
            const fontSize = clamp(textRun.fontSize * 0.75, FONT_SIZE_MIN, FONT_SIZE_MAX);

            slide.addText(safeText, {
                x: rect.x,
                y: rect.y,
                w: rect.w,
                h: rect.h,
                fontSize,
                fontFace: textRun.fontFace || 'Arial',
                color: '363636',
                valign: 'top',
                breakLine: false,
                fit: 'shrink'
            });
        }

        const progress = Math.round(((index + 1) / totalPages) * 100);
        onProgress(progress);
    }

    return pptx.write({ outputType: 'blob' });
}

function shouldUseRasterFallback(pageData, options) {
    if (!pageData.pageRaster || !pageData.pageRaster.dataUrl) {
        return false;
    }

    if (options.mode === 'fidelity') {
        return true;
    }

    if (options.mode === 'editable') {
        return false;
    }

    return pageData.texts.length < BALANCED_TEXT_THRESHOLD;
}

function pdfRectToSlideRect(textRun, pageData, slideSize) {
    const pageWidth = Math.max(1, Number(pageData.width) || 1);
    const pageHeight = Math.max(1, Number(pageData.height) || 1);

    const rawX = (Number(textRun.x) / pageWidth) * slideSize.width;
    const rawY = slideSize.height - ((Number(textRun.y) / pageHeight) * slideSize.height);
    const rawW = (Math.max(0, Number(textRun.w) || 0) / pageWidth) * slideSize.width;
    const rawH = Math.max(0.08, (Math.max(1, Number(textRun.h) || 1) / pageHeight) * slideSize.height * 1.2);

    const x = clamp(rawX, 0.05, slideSize.width - 0.15);
    const y = clamp(rawY - rawH, 0.05, slideSize.height - 0.1);
    const maxWidth = Math.max(0.1, slideSize.width - x - 0.05);
    const w = clamp(rawW || 1, 0.1, maxWidth);
    const h = clamp(rawH, 0.08, slideSize.height - y - 0.05);

    return { x, y, w, h };
}

function clamp(value, min, max) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) {
        return min;
    }
    return Math.min(max, Math.max(min, numeric));
}

function normalizeFontName(fontName) {
    const name = String(fontName || '').trim();
    if (!name) {
        return 'Arial';
    }

    if (name.includes('+')) {
        return name.split('+').pop() || 'Arial';
    }

    return name;
}

function logDebugInfo(options, pageDataList, extractorStats) {
    const summary = {
        options,
        totalPages: pageDataList.length,
        totalTextRuns: pageDataList.reduce((sum, page) => sum + page.texts.length, 0),
        pagesWithRaster: pageDataList.filter((page) => Boolean(page.pageRaster)).length,
        skippedTextRuns: extractorStats.skippedTextRuns,
        mergedTextRuns: extractorStats.mergedTextRuns,
        rasterPages: extractorStats.rasterPages
    };

    console.group('[PDF->PPTX Debug Summary]');
    console.table(summary);
    console.groupEnd();
}
