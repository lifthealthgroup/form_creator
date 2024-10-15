const fileInput = document.getElementById('file-input');
const filePreview = document.getElementById('file-preview');
let selectedFiles = [];

// Function to update the file preview
function updateFilePreview() {
    filePreview.innerHTML = '';  // Clear the preview
    filePreview.style.display = selectedFiles.length > 0 ? 'block' : 'none'; // Show or hide the preview

    selectedFiles.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <span>${file.name}</span>
            <button class="remove-btn" data-index="${index}">X</button>
        `;
        filePreview.appendChild(fileItem);
    });

    // Add event listeners to all "remove" buttons
    document.querySelectorAll('.remove-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const index = this.getAttribute('data-index');
            removeFile(index);
        });
    });
}

// Function to remove a file from the selected list
function removeFile(index) {
    selectedFiles.splice(index, 1);  // Remove the file
    updateFilePreview();  // Update the preview
}

// Open the file dialog when the Choose Files button is clicked
fileInput.addEventListener('change', function() {
    // Add the selected files to the list
    for (const file of this.files) {
        // Check if the file is already in the array to prevent duplicates
        if (!selectedFiles.some(f => f.name === file.name)) {
            selectedFiles.push(file);
        }
    }
    updateFilePreview();  // Update the preview
});
