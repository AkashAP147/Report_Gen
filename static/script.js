document.addEventListener('DOMContentLoaded', () => {
    // Step 1 Elements
    const step1 = document.getElementById('step-1');
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');
    const parseProgress = document.getElementById('parse-progress');
    
    // Step 2 Elements
    const step2 = document.getElementById('step-2');
    const fileName = document.getElementById('file-name');
    const prefLeft = document.getElementById('pref-left');
    const prefRight = document.getElementById('pref-right');
    const generateBtn = document.getElementById('generate-btn');
    const generateProgress = document.getElementById('generate-progress');
    
    const errorMessage = document.getElementById('error-message');
    
    let currentSessionData = null;

    // Trigger file selection on click
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // Handle Drag & Drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        uploadArea.addEventListener(eventName, () => uploadArea.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, () => uploadArea.classList.remove('dragover'), false);
    });

    uploadArea.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        handleFiles(dt.files);
    });

    // Handle file input change
    fileInput.addEventListener('change', function() {
        handleFiles(this.files);
    });

    async function handleFiles(files) {
        if (files.length === 0) return;
        
        const validTypes = ['.xls', '.xlsx', '.csv'];
        let allValid = true;
        
        Array.from(files).forEach(file => {
            const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
            if (!validTypes.includes(fileExt)) {
                allValid = false;
            }
        });
        
        if (!allValid) {
            showError("Please upload valid Excel or CSV files (.xls, .xlsx, .csv).");
            return;
        }

        // Show parsing state
        errorMessage.style.display = 'none';
        uploadArea.style.display = 'none';
        parseProgress.style.display = 'flex';

        const formData = new FormData();
        Array.from(files).forEach(file => {
            formData.append('file', file);
        });

        try {
            const response = await fetch('/parse', {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                const data = await response.json();
                currentSessionData = data;
                
                // Populate dropdowns
                populateDropdowns(data.subjects);
                if (files.length > 1) {
                    fileName.textContent = `${files.length} files selected`;
                } else {
                    fileName.textContent = files[0].name;
                }
                
                // Transition to Step 2
                step1.classList.remove('active');
                step2.classList.add('active');
            } else {
                const errorData = await response.json();
                showError(errorData.error || 'Failed to parse file.');
                resetStep1();
            }
        } catch (error) {
            showError('A network error occurred while parsing the file.');
            resetStep1();
        }
    }

    function populateDropdowns(subjects) {
        prefLeft.innerHTML = '<option value="">-- Let algorithm decide --</option>';
        prefRight.innerHTML = '<option value="">-- Let algorithm decide --</option>';
        
        subjects.forEach(subj => {
            const opt1 = document.createElement('option');
            opt1.value = subj;
            opt1.textContent = subj;
            prefLeft.appendChild(opt1);
            
            const opt2 = document.createElement('option');
            opt2.value = subj;
            opt2.textContent = subj;
            prefRight.appendChild(opt2);
        });
        
        // Auto-select first two if available
        if (subjects.length > 0) prefLeft.value = subjects[0];
        if (subjects.length > 1) prefRight.value = subjects[1];
    }

    // Generate Button Click
    generateBtn.addEventListener('click', async () => {
        if (!currentSessionData) return;

        // UI updates for generating state
        generateBtn.style.display = 'none';
        generateProgress.style.display = 'flex';
        errorMessage.style.display = 'none';

        const payload = {
            session_id: currentSessionData.session_id,
            filenames: currentSessionData.filenames,
            pref_left: prefLeft.value || null,
            pref_right: prefRight.value || null
        };

        try {
            const response = await fetch('/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const data = await response.json();
                
                // Trigger multiple downloads gracefully
                downloadFiles(data.files);
                
                // Show completion
                setTimeout(() => {
                    generateProgress.querySelector('.progress-text').textContent = 'Downloads complete! You can refresh to generate more.';
                    generateProgress.querySelector('.spinner').style.display = 'none';
                }, 1000);
            } else {
                const errorData = await response.json();
                showError(errorData.error || 'Failed to generate reports.');
                resetBtn();
            }
        } catch (error) {
            showError('A network error occurred during generation.');
            resetBtn();
        }
    });
    
    async function downloadFiles(urls) {
        // Trigger downloads with a slight delay between each to avoid browser blocking
        for (let i = 0; i < urls.length; i++) {
            const url = urls[i];
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = '';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            // Wait 500ms between downloads
            await new Promise(resolve => setTimeout(resolve, 500));
        }
    }

    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.style.display = 'block';
    }

    function resetStep1() {
        uploadArea.style.display = 'block';
        parseProgress.style.display = 'none';
        fileInput.value = '';
    }

    function resetBtn() {
        generateBtn.style.display = 'block';
        generateProgress.style.display = 'none';
    }
});
