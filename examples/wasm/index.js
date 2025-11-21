// xlsxzero WASM Example - JavaScript code

import init, {
    convert_excel_to_markdown,
    convert_excel_to_markdown_custom,
    get_version
} from './pkg/xlsxzero_wasm.js';

let wasmInitialized = false;
let currentFileData = null;

// Initialize WASM module
async function initWasm() {
    if (wasmInitialized) return;
    
    try {
        await init();
        wasmInitialized = true;
        console.log('WASM module initialized. Version:', get_version());
    } catch (error) {
        console.error('Failed to initialize WASM module:', error);
        showError('WASM モジュールの初期化に失敗しました: ' + error.message);
    }
}

// Initialize on page load
initWasm();

// DOM elements
const fileInput = document.getElementById('fileInput');
const uploadSection = document.getElementById('uploadSection');
const fileInfo = document.getElementById('fileInfo');
const convertBtn = document.getElementById('convertBtn');
const clearBtn = document.getElementById('clearBtn');
const resultSection = document.getElementById('resultSection');
const result = document.getElementById('result');
const copyBtn = document.getElementById('copyBtn');
const downloadBtn = document.getElementById('downloadBtn');
const statusMessage = document.getElementById('statusMessage');
const sheetIndexInput = document.getElementById('sheetIndex');
const mergeStrategySelect = document.getElementById('mergeStrategy');
const dateFormatInput = document.getElementById('dateFormat');

// File input handler
fileInput.addEventListener('change', handleFileSelect);

// Drag and drop handlers
uploadSection.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadSection.classList.add('dragover');
});

uploadSection.addEventListener('dragleave', () => {
    uploadSection.classList.remove('dragover');
});

uploadSection.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadSection.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

// File selection handler
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

// Handle file (both from input and drag-drop)
async function handleFile(file) {
    // Check file type
    if (!file.name.match(/\.(xlsx|xlsm|xlsb)$/i)) {
        showError('Excel ファイル (.xlsx, .xlsm, .xlsb) を選択してください。');
        return;
    }

    // Check file size (limit to 50MB)
    if (file.size > 50 * 1024 * 1024) {
        showError('ファイルサイズが大きすぎます（最大 50MB）。');
        return;
    }

    // Read file as ArrayBuffer
    try {
        showStatus('ファイルを読み込んでいます...', 'info');
        const arrayBuffer = await file.arrayBuffer();
        currentFileData = new Uint8Array(arrayBuffer);
        
        fileInfo.textContent = `選択中: ${file.name} (${formatFileSize(file.size)})`;
        convertBtn.disabled = false;
        hideStatus();
    } catch (error) {
        console.error('Error reading file:', error);
        showError('ファイルの読み込みに失敗しました: ' + error.message);
    }
}

// Convert button handler
convertBtn.addEventListener('click', async () => {
    if (!currentFileData || !wasmInitialized) {
        showError('ファイルが選択されていないか、WASM が初期化されていません。');
        return;
    }

    try {
        showStatus('変換中...', 'loading');
        convertBtn.disabled = true;
        resultSection.classList.add('hidden');

        // Get options
        const sheetIndexStr = sheetIndexInput.value.trim();
        const sheetIndex = sheetIndexStr ? parseInt(sheetIndexStr, 10) : null;
        
        const mergeStrategy = mergeStrategySelect.value;
        const dateFormatStr = dateFormatInput.value.trim();
        const dateFormat = dateFormatStr || null;

        // Convert
        let markdown;
        if (sheetIndex !== null || mergeStrategy || dateFormat) {
            // Use custom conversion
            markdown = convert_excel_to_markdown_custom(
                currentFileData,
                sheetIndex,
                mergeStrategy || null,
                dateFormat
            );
        } else {
            // Use default conversion
            markdown = convert_excel_to_markdown(currentFileData);
        }

        // Display result
        result.value = markdown;
        resultSection.classList.remove('hidden');
        showStatus('変換が完了しました！', 'success');
        
    } catch (error) {
        console.error('Conversion error:', error);
        showError('変換エラー: ' + error.message);
    } finally {
        convertBtn.disabled = false;
    }
});

// Clear button handler
clearBtn.addEventListener('click', () => {
    fileInput.value = '';
    currentFileData = null;
    fileInfo.textContent = 'またはファイルをここにドラッグ＆ドロップ';
    convertBtn.disabled = true;
    resultSection.classList.add('hidden');
    result.value = '';
    sheetIndexInput.value = '';
    dateFormatInput.value = '';
    mergeStrategySelect.value = 'data_duplication';
    hideStatus();
});

// Copy button handler
copyBtn.addEventListener('click', async () => {
    try {
        await navigator.clipboard.writeText(result.value);
        showStatus('クリップボードにコピーしました！', 'success');
        setTimeout(hideStatus, 2000);
    } catch (error) {
        console.error('Copy failed:', error);
        showError('コピーに失敗しました: ' + error.message);
    }
});

// Download button handler
downloadBtn.addEventListener('click', () => {
    const blob = new Blob([result.value], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'converted.md';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showStatus('ダウンロードを開始しました！', 'success');
    setTimeout(hideStatus, 2000);
});

// Utility functions
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

function showStatus(message, type = 'info') {
    statusMessage.textContent = message;
    statusMessage.className = type === 'error' ? 'error' : 
                             type === 'success' ? 'success' : 
                             type === 'loading' ? 'loading' : '';
    statusMessage.classList.remove('hidden');
}

function showError(message) {
    showStatus(message, 'error');
}

function hideStatus() {
    statusMessage.classList.add('hidden');
}

