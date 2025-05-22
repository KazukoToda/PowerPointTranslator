document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const dropArea = document.getElementById('drop-area');
    const fileInput = document.getElementById('file-input');
    const selectFileBtn = document.getElementById('select-file-btn');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const uploadBtn = document.getElementById('upload-btn');
    const cancelBtn = document.getElementById('cancel-btn');
    const progressSection = document.getElementById('progress-section');
    const progressBar = document.getElementById('progress-bar');
    const progressStatus = document.getElementById('progress-status');
    const resultSection = document.getElementById('result-section');
    const downloadBtn = document.getElementById('download-btn');
    const newTranslationBtn = document.getElementById('new-translation-btn');
    const errorSection = document.getElementById('error-section');
    const errorMessage = document.getElementById('error-message');
    const tryAgainBtn = document.getElementById('try-again-btn');

    // Variables
    let selectedFile = null;
    let translatedFilePath = null;

    // Event Listeners
    selectFileBtn.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', handleFileSelect);

    dropArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropArea.classList.add('active');
    });

    dropArea.addEventListener('dragleave', () => {
        dropArea.classList.remove('active');
    });

    dropArea.addEventListener('drop', handleFileDrop);

    uploadBtn.addEventListener('click', uploadFile);
    cancelBtn.addEventListener('click', resetUpload);
    newTranslationBtn.addEventListener('click', resetAll);
    tryAgainBtn.addEventListener('click', resetAll);
    downloadBtn.addEventListener('click', downloadTranslatedFile);

    // Functions
    function handleFileSelect(e) {
        const file = e.target.files[0];
        if (file && validateFile(file)) {
            selectedFile = file;
            showFileInfo();
        }
    }

    function handleFileDrop(e) {
        e.preventDefault();
        dropArea.classList.remove('active');
        
        const file = e.dataTransfer.files[0];
        if (file && validateFile(file)) {
            selectedFile = file;
            fileInput.files = e.dataTransfer.files;
            showFileInfo();
        }
    }

    function validateFile(file) {
        // Check file type
        const validTypes = ['application/vnd.openxmlformats-officedocument.presentationml.presentation'];
        if (!validTypes.includes(file.type)) {
            showError('PowerPoint (.pptx) ファイルのみサポートしています。');
            return false;
        }

        // Check file size (50MB max)
        const maxSize = 50 * 1024 * 1024; // 50MB
        if (file.size > maxSize) {
            showError('ファイルサイズは50MB以下にしてください。');
            return false;
        }

        return true;
    }

    function showFileInfo() {
        dropArea.classList.add('hidden');
        fileInfo.classList.remove('hidden');
        fileName.textContent = selectedFile.name;
    }

    function resetUpload() {
        fileInfo.classList.add('hidden');
        dropArea.classList.remove('hidden');
        fileInput.value = '';
        selectedFile = null;
    }

    function resetAll() {
        resetUpload();
        progressSection.classList.add('hidden');
        resultSection.classList.add('hidden');
        errorSection.classList.add('hidden');
        progressBar.style.width = '0%';
        translatedFilePath = null;
    }

    function showError(message) {
        errorMessage.textContent = message;
        errorSection.classList.remove('hidden');
        dropArea.classList.add('hidden');
        fileInfo.classList.add('hidden');
        progressSection.classList.add('hidden');
        resultSection.classList.add('hidden');
    }

    async function uploadFile() {
        if (!selectedFile) return;

        // Show progress section
        fileInfo.classList.add('hidden');
        progressSection.classList.remove('hidden');
        
        try {
            // Simulate progress
            simulateProgress();
            
            // Create form data
            const formData = new FormData();
            formData.append('file', selectedFile);
            
            // Upload file
            const response = await fetch('/translate', {
                method: 'POST',
                body: formData,
            });
            
            const data = await response.json();
            
            if (!response.ok) {
                throw new Error(data.error || 'ファイルの翻訳に失敗しました。');
            }
            
            // Store translated file path
            translatedFilePath = data.filePath;
            
            // Show success
            progressSection.classList.add('hidden');
            resultSection.classList.remove('hidden');
        } catch (error) {
            console.error('Error:', error);
            showError(error.message || 'ファイルの翻訳に失敗しました。');
        }
    }

    function simulateProgress() {
        let progress = 0;
        const interval = setInterval(() => {
            progress += Math.random() * 10;
            if (progress > 90) {
                progress = 90; // Cap at 90% until actual completion
                clearInterval(interval);
            }
            progressBar.style.width = `${progress}%`;
        }, 500);
        
        // Store interval ID to clear it later
        window.progressInterval = interval;
    }

    function downloadTranslatedFile() {
        if (!translatedFilePath) return;
        
        const downloadUrl = `/download/${translatedFilePath}`;
        window.location.href = downloadUrl;
    }
});