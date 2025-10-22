document.addEventListener('DOMContentLoaded', () => {
    const processButton = document.getElementById('processButton');
    processButton.disabled = true;
    processButton.textContent = 'Loading...';

    // Section: Dark Mode Toggle
    // Handles dark mode preference and toggle
    const darkModeToggle = document.getElementById('darkModeToggle');
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    const prefersDark = mediaQuery.matches;
    document.body.classList.toggle('dark', prefersDark);
    if (darkModeToggle) darkModeToggle.checked = prefersDark;

    mediaQuery.addEventListener('change', (e) => {
        document.body.classList.toggle('dark', e.matches);
        if (darkModeToggle) darkModeToggle.checked = e.matches;
    });

    if (darkModeToggle) {
        darkModeToggle.addEventListener('change', () => {
            document.body.classList.toggle('dark');
        });
    }

    // Section: Scroll to Top Button
    // Displays button when scrolled down
    const scrollTopBtn = document.getElementById('scrollTopBtn');
    window.addEventListener('scroll', () => {
        scrollTopBtn.classList.toggle("show", window.scrollY > 300);
    });
    scrollTopBtn.addEventListener('click', () => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
});

// Section: PyScript Readiness
// Hides loader and enables button when PyScript is ready
window.addEventListener('py:all-done', () => {
    const loaderOverlay = document.querySelector('.loader-overlay');
    loaderOverlay.style.display = 'none';
    const processButton = document.getElementById('processButton');
    processButton.disabled = false;
    processButton.textContent = 'Process and Download';
});

// Section: Form Submission Handling
// Processes uploaded files and handles Python interaction
document.getElementById('extractForm').addEventListener('submit', async function(event) {
    event.preventDefault();

    const files = document.getElementById('files').files;
    if (files.length === 0) {
        displayMessage('Please upload at least one file.', 'error');
        return;
    }

    if (typeof window.process_files !== 'function') {
        displayMessage('PyScript is not ready yet. Please wait a moment and try again.', 'error');
        return;
    }

    let fileInfos = [];
    for (let file of files) {
        const arrayBuf = await file.arrayBuffer();
        const uint8 = new Uint8Array(arrayBuf);
        const dataArray = Array.from(uint8);
        fileInfos.push({name: file.name, data: dataArray});
    }

    try {
        const result = window.process_files(fileInfos);
        displayMessage(result.message, result.type || 'success');
        if (result.buffer) {
            const js_buffer = new Uint8Array(result.buffer);
            const blob = new Blob([js_buffer], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            const today = new Date();
            const dd = String(today.getDate()).padStart(2, '0');
            const mm = String(today.getMonth() + 1).padStart(2, '0');
            const yyyy = today.getFullYear();
            a.download = `Sales Report ${dd}-${mm}-${yyyy}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);
        }
    } catch (error) {
        console.error(error);
        displayMessage('An error occurred during processing.', 'error');
    }
});

// Section: Message Display
// Displays messages with styling for errors/warnings
function displayMessage(msg, type) {
    const messageDiv = document.getElementById('message');
    messageDiv.innerHTML = msg;
    messageDiv.className = type;  // e.g., 'error', 'warning', 'success'
}