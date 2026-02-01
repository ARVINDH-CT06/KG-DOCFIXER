let selectedFile = null;
let formattedBlob = null;

// File input change handler
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        handleFileSelect(file);
    }
});

// Drag and drop handlers
const uploadBox = document.getElementById('uploadBox');

uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadBox.style.borderColor = '#2563eb';
    uploadBox.style.background = '#eff6ff';
});

uploadBox.addEventListener('dragleave', (e) => {
    e.preventDefault();
    uploadBox.style.borderColor = '#cbd5e1';
    uploadBox.style.background = 'transparent';
});

uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadBox.style.borderColor = '#cbd5e1';
    uploadBox.style.background = 'transparent';
    
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.docx')) {
        handleFileSelect(file);
    } else {
        alert('Please upload a .docx file only!');
    }
});

function handleFileSelect(file) {
    selectedFile = file;
    
    // Show file info
    document.getElementById('uploadBox').style.display = 'none';
    document.getElementById('fileInfo').style.display = 'block';
    
    // Display file details
    document.getElementById('fileName').textContent = file.name;
    document.getElementById('fileSize').textContent = formatFileSize(file.size);
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

async function processDocument() {
    if (!selectedFile) {
        alert('Please select a file first!');
        return;
    }
    
    // Show processing screen
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('processing').style.display = 'block';
    
    // Send file to backend for processing
    const formData = new FormData();
    formData.append('file', selectedFile);
    
    try {
        const response = await fetch('/format', {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            formattedBlob = await response.blob();
            showResult();
        } else {
            throw new Error('Processing failed');
        }
    } catch (error) {
        console.error('Error:', error);
        alert('Error processing document. Please make sure the backend server is running!');
        resetApp();
    }
}

function showResult() {
    document.getElementById('processing').style.display = 'none';
    document.getElementById('result').style.display = 'block';
    
    // Setup download button
    document.getElementById('downloadBtn').onclick = function() {
        const url = window.URL.createObjectURL(formattedBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'formatted_' + selectedFile.name;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    };
}

function resetApp() {
    selectedFile = null;
    formattedBlob = null;
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadBox').style.display = 'block';
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('processing').style.display = 'none';
    document.getElementById('result').style.display = 'none';
}