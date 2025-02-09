document.addEventListener('DOMContentLoaded', function() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const browseButton = document.getElementById('browseButton');
    const uploadForm = document.getElementById('uploadForm');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const convertButton = document.getElementById('convertButton');
    const progressArea = document.getElementById('progressArea');
    const progressBar = progressArea.querySelector('.progress-bar');
    const errorMessage = document.getElementById('errorMessage');

    // Handle drag and drop events
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove('dragover');
        });
    });

    // Handle file drop
    dropZone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        handleFiles(files);
    });

    // Handle file browse
    browseButton.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
    });

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            if (file.type === 'application/pdf') {
                showFileInfo(file);
                convertButton.disabled = false;
                errorMessage.classList.add('d-none');
            } else {
                showError('Please select a PDF file');
                resetForm();
            }
        }
    }

    function showFileInfo(file) {
        fileName.textContent = file.name;
        fileInfo.classList.remove('d-none');
    }

    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.remove('d-none');
    }

    function resetForm() {
        fileInput.value = '';
        fileInfo.classList.add('d-none');
        convertButton.disabled = true;
        progressArea.classList.add('d-none');
        progressBar.style.width = '0%';
    }

    // Handle form submission
    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const formData = new FormData();
        const file = fileInput.files[0];
        if (!file) {
            showError('Please select a file first');
            return;
        }
        formData.append('file', file);

        try {
            // Show progress
            progressArea.classList.remove('d-none');
            progressBar.style.width = '50%';
            convertButton.disabled = true;
            errorMessage.classList.add('d-none');

            const response = await fetch('/upload', {
                method: 'POST',
                body: formData,
                headers: {
                    'Accept': 'application/json',
                },
            });

            if (!response.ok) {
                const contentType = response.headers.get('content-type');
                if (contentType && contentType.includes('application/json')) {
                    const data = await response.json();
                    throw new Error(data.error || 'Error processing file');
                } else {
                    throw new Error('Error processing file');
                }
            }

            // Complete progress
            progressBar.style.width = '100%';

            // Get the file
            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);
            const downloadButton = document.getElementById('downloadButton');
            downloadButton.href = downloadUrl;
            downloadButton.download = file.name.replace('.pdf', '.xlsx');
            downloadButton.style.display = 'inline-block';
            
            // Complete progress
            progressBar.style.width = '100%';
            convertButton.disabled = false;

        } catch (error) {
            console.error('Error:', error);
            showError(error.message || 'An error occurred while processing the file');
            progressArea.classList.add('d-none');
            progressBar.style.width = '0%';
            convertButton.disabled = false;
        }
    });
});