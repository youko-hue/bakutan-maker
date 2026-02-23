// File upload and processing functionality

// Function to handle file selection validation
function validateFile(file) {
    const allowedExtensions = /(.jpg|.jpeg|.png|.gif|.pdf|.xlsx)$/i;
    return allowedExtensions.test(file.name);
}

// Function to upload file to /process endpoint
async function uploadFile(file) {
    if (!validateFile(file)) {
        alert('Invalid file type.');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch('/process', {
        method: 'POST',
        body: formData,
        onUploadProgress: (progressEvent) => {
            const percentCompleted = Math.round((progressEvent.loaded * 100) / progressEvent.total);
            updateProgress(percentCompleted);
        }
    });

    if (response.ok) {
        const downloadLink = await response.json();
        showDownloadLink(downloadLink);
    } else {
        alert('File upload failed.');
    }
}

// Function to update progress bar
function updateProgress(percent) {
    const progressBar = document.getElementById('progress-bar');
    progressBar.style.width = percent + '%';
    progressBar.innerText = percent + '%';
}

// Function to show the download link
function showDownloadLink(link) {
    const downloadContainer = document.getElementById('download-link');
    downloadContainer.innerHTML = `<a href='${link}' target='_blank'>Download Processed File</a>`;
}

// Event listener for file upload
document.getElementById('file-input').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        uploadFile(file);
    }
});