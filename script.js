let templateImg = new Image();
let excelData = [];
let positions = {};
let textStyles = {};
let selectedFields = new Set();
let hoverField = null;

const canvas = document.getElementById('preview');
const ctx = canvas.getContext('2d');
let isDragging = false;
let selectedField = null;
let offsetX, offsetY;
let MOVE_STEP = 5; // Dynamic step size

// Enhanced Google Fonts list
const googleFonts = [
    'Arial', 'Times New Roman', 'Courier', 'Inter', 'Roboto', 'Open Sans',
    'Lato', 'Montserrat', 'Merriweather', 'Oswald', 'Raleway', 'Source Sans Pro',
    'Playfair Display', 'Ubuntu', 'Nunito', 'Pacifico', 'Lobster', 'Bree Serif',
    'Poppins', 'Quicksand', 'Dancing Script', 'Crimson Text', 'Fira Sans',
    'Work Sans', 'PT Sans', 'Libre Baskerville', 'Cabin', 'Dosis'
];

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    setupEventListeners();
    loadGoogleFonts();
    checkFirstVisit();
    initializeTheme();
});

function initializeApp() {
    // Hide canvas initially and show placeholder
    canvas.style.display = 'none';
    document.getElementById('canvasPlaceholder').style.display = 'flex';
    
    // Initialize move step
    updateMoveStep();
}

function setupEventListeners() {
    // File upload listeners
    document.getElementById('excelInput').addEventListener('change', handleExcelUpload);
    document.getElementById('templateInput').addEventListener('change', handleTemplateUpload);
    
    // Style control listeners
    document.getElementById('fieldSelect').addEventListener('change', handleFieldSelect);
    document.getElementById('fontSearch').addEventListener('input', handleFontSearch);
    document.getElementById('fontSelect').addEventListener('change', handleFontChange);
    document.getElementById('fontSize').addEventListener('change', handleFontSizeChange);
    document.getElementById('fontSizeRange').addEventListener('input', handleFontSizeRangeChange);
    document.getElementById('textColor').addEventListener('change', handleColorChange);
    document.getElementById('moveStep').addEventListener('change', updateMoveStep);
    
    // Alignment button listeners
    document.querySelectorAll('.align-btn').forEach(btn => {
        btn.addEventListener('click', handleAlignmentChange);
    });
    
    // Canvas interaction listeners
    canvas.addEventListener('mousedown', handleMouseDown);
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
    canvas.addEventListener('mousemove', handleCanvasHover);
    canvas.addEventListener('mouseleave', () => {
        canvas.style.cursor = 'default';
        hoverField = null;
    });
}

// File upload handlers
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    // Update UI to show loading
    const label = e.target.previousElementSibling;
    const originalContent = label.innerHTML;
    label.innerHTML = '<i class="fas fa-spinner fa-spin"></i><span>Loading...</span>';
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(sheet);
            
            if (excelData.length === 0) {
                throw new Error('Excel file is empty');
            }

            // Initialize data structures
            const headings = Object.keys(excelData[0]);
            positions = {};
            textStyles = {};
            selectedFields.clear();
            
            headings.forEach((field, index) => {
                positions[field] = { x: 100, y: 100 + (index * 60) };
                textStyles[field] = { 
                    font: 'Inter', 
                    size: 24, 
                    align: 'left', 
                    color: '#000000' 
                };
                selectedFields.add(field);
            });

            populateFieldOptions(headings);
            updatePreview();
            
            // Success feedback
            label.classList.add('success');
            label.innerHTML = '<i class="fas fa-check"></i><span>Excel Loaded</span><small>(' + excelData.length + ' records)</small>';
            
        } catch (error) {
            console.error('Error reading Excel file:', error);
            label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Error loading file</span><small>Please try again</small>';
        }
    };
    reader.readAsArrayBuffer(file);
}

function handleTemplateUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const label = e.target.previousElementSibling;
    const originalContent = label.innerHTML;
    label.innerHTML = '<i class="fas fa-spinner fa-spin"></i><span>Loading...</span>';
    
    const reader = new FileReader();
    reader.onload = function(e) {
        templateImg.src = e.target.result;
        templateImg.onload = function() {
            // Calculate optimal canvas size
            const maxWidth = 800;
            const maxHeight = 600;
            let { width, height } = templateImg;
            
            if (width > maxWidth || height > maxHeight) {
                const ratio = Math.min(maxWidth / width, maxHeight / height);
                width *= ratio;
                height *= ratio;
            }
            
            canvas.width = width;
            canvas.height = height;
            
            // Show canvas and hide placeholder
            canvas.style.display = 'block';
            document.getElementById('canvasPlaceholder').style.display = 'none';
            
            updatePreview();
            
            // Success feedback
            label.classList.add('success');
            label.innerHTML = '<i class="fas fa-check"></i><span>Template Loaded</span><small>(' + templateImg.width + 'x' + templateImg.height + ')</small>';
        };
    };
    reader.readAsDataURL(file);
}

// Field management
function populateFieldOptions(headings) {
    const fieldSelect = document.getElementById('fieldSelect');
    const fieldCheckboxes = document.getElementById('fieldCheckboxes');
    
    fieldSelect.innerHTML = '<option value="">Select a field to edit</option>';
    fieldCheckboxes.innerHTML = '';

    headings.forEach(field => {
        // Populate dropdown
        const option = document.createElement('option');
        option.value = field;
        option.textContent = field;
        fieldSelect.appendChild(option);

        // Populate checkboxes
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = true;
        checkbox.value = field;
        checkbox.addEventListener('change', function() {
            if (this.checked) {
                selectedFields.add(field);
            } else {
                selectedFields.delete(field);
            }
            updatePreview();
        });
        
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(field));
        fieldCheckboxes.appendChild(label);
    });

    if (headings.length > 0) {
        selectedField = headings[0];
        fieldSelect.value = selectedField;
        updateStyleControls();
    }
}

// Style control handlers
function handleFieldSelect() {
    selectedField = document.getElementById('fieldSelect').value;
    updateStyleControls();
    updatePreview();
}

function handleFontSearch(e) {
    loadGoogleFonts(e.target.value);
}

function handleFontChange() {
    if (!selectedField) return;
    
    const selectedFont = document.getElementById('fontSelect').value;
    textStyles[selectedField].font = selectedFont;

    // Load Google Font dynamically
    if (!googleFonts.includes(selectedFont)) {
        WebFont.load({
            google: { families: [selectedFont] },
            active: () => updatePreview(),
            inactive: () => {
                console.warn(`Font ${selectedFont} failed to load`);
                textStyles[selectedField].font = 'Arial';
                updatePreview();
            }
        });
    } else {
        updatePreview();
    }
}

function handleFontSizeChange() {
    if (!selectedField) return;
    
    const size = parseInt(document.getElementById('fontSize').value);
    textStyles[selectedField].size = size;
    document.getElementById('fontSizeRange').value = size;
    updatePreview();
}

function handleFontSizeRangeChange() {
    if (!selectedField) return;
    
    const size = parseInt(document.getElementById('fontSizeRange').value);
    textStyles[selectedField].size = size;
    document.getElementById('fontSize').value = size;
    updatePreview();
}

function handleColorChange() {
    if (!selectedField) return;
    
    const color = document.getElementById('textColor').value;
    textStyles[selectedField].color = color;
    document.querySelector('.color-value').textContent = color;
    updatePreview();
}

function handleAlignmentChange(e) {
    if (!selectedField) return;
    
    // Update button states
    document.querySelectorAll('.align-btn').forEach(btn => btn.classList.remove('active'));
    e.target.closest('.align-btn').classList.add('active');
    
    const alignment = e.target.closest('.align-btn').dataset.align;
    textStyles[selectedField].align = alignment;
    document.getElementById('textAlign').value = alignment;
    updatePreview();
}

function updateMoveStep() {
    MOVE_STEP = parseInt(document.getElementById('moveStep').value);
}

// Canvas interaction handlers
function handleMouseDown(e) {
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (canvas.width / rect.width);
    const y = (e.clientY - rect.top) * (canvas.height / rect.height);
    
    // Check if clicking on any text field
    for (let field of selectedFields) {
        const pos = positions[field];
        const styles = textStyles[field];
        const sampleData = excelData.length > 0 ? excelData[0] : {};
        const text = sampleData[field] || field;
        
        // Calculate text bounds
        ctx.font = `${styles.size}px "${styles.font}"`;
        const metrics = ctx.measureText(text);
        const textWidth = metrics.width;
        const textHeight = styles.size;
        
        let textX = pos.x;
        if (styles.align === 'center') textX -= textWidth / 2;
        else if (styles.align === 'right') textX -= textWidth;
        
        const padding = 10;
        if (x >= textX - padding && 
            x <= textX + textWidth + padding && 
            y >= pos.y - textHeight - padding && 
            y <= pos.y + padding) {
            
            isDragging = true;
            selectedField = field;
            offsetX = x - pos.x;
            offsetY = y - pos.y;
            
            // Update UI
            document.getElementById('fieldSelect').value = field;
            updateStyleControls();
            canvas.style.cursor = 'grabbing';
            break;
        }
    }
}

function handleMouseMove(e) {
    if (!isDragging || !selectedField) return;
    
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (canvas.width / rect.width);
    const y = (e.clientY - rect.top) * (canvas.height / rect.height);

    // Update position with bounds checking
    positions[selectedField] = {
        x: Math.max(0, Math.min(x - offsetX, canvas.width)),
        y: Math.max(20, Math.min(y - offsetY, canvas.height - 20))
    };
    
    updatePreview();
}

function handleMouseUp() {
    if (isDragging) {
        isDragging = false;
        canvas.style.cursor = 'grab';
    }
}

function handleCanvasHover(e) {
    if (isDragging) return;
    
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (canvas.width / rect.width);
    const y = (e.clientY - rect.top) * (canvas.height / rect.height);
    
    let foundField = null;
    
    for (let field of selectedFields) {
        const pos = positions[field];
        const styles = textStyles[field];
        const sampleData = excelData.length > 0 ? excelData[0] : {};
        const text = sampleData[field] || field;
        
        ctx.font = `${styles.size}px "${styles.font}"`;
        const metrics = ctx.measureText(text);
        const textWidth = metrics.width;
        const textHeight = styles.size;
        
        let textX = pos.x;
        if (styles.align === 'center') textX -= textWidth / 2;
        else if (styles.align === 'right') textX -= textWidth;
        
        const padding = 10;
        if (x >= textX - padding && 
            x <= textX + textWidth + padding && 
            y >= pos.y - textHeight - padding && 
            y <= pos.y + padding) {
            foundField = field;
            break;
        }
    }
    
    canvas.style.cursor = foundField ? 'grab' : 'default';
    hoverField = foundField;
}

// Movement functions
function moveText(direction) {
    if (!selectedField) {
        alert('Please select a field first!');
        return;
    }

    const pos = positions[selectedField];
    switch (direction) {
        case 'up':
            pos.y = Math.max(20, pos.y - MOVE_STEP);
            break;
        case 'down':
            pos.y = Math.min(canvas.height - 20, pos.y + MOVE_STEP);
            break;
        case 'left':
            pos.x = Math.max(0, pos.x - MOVE_STEP);
            break;
        case 'right':
            pos.x = Math.min(canvas.width, pos.x + MOVE_STEP);
            break;
    }
    
    updatePreview();
}

function resetPosition() {
    if (!selectedField) return;
    
    const headings = Object.keys(positions);
    const idx = headings.indexOf(selectedField);
    positions[selectedField] = { x: 100, y: 100 + (idx * 60) };
    updatePreview();
}

function resetAllPositions() {
    const headings = Object.keys(positions);
    headings.forEach((field, index) => {
        positions[field] = { x: 100, y: 100 + (index * 60) };
    });
    updatePreview();
}

// Font management
function loadGoogleFonts(searchTerm = '') {
    const fontSelect = document.getElementById('fontSelect');
    fontSelect.innerHTML = '';

    const filteredFonts = googleFonts.filter(font =>
        font.toLowerCase().includes(searchTerm.toLowerCase())
    );

    filteredFonts.forEach(font => {
        const option = document.createElement('option');
        option.value = font;
        option.textContent = font;
        fontSelect.appendChild(option);
    });

    if (selectedField && textStyles[selectedField]) {
        fontSelect.value = textStyles[selectedField].font || 'Inter';
    }
}

// UI update functions
function updateStyleControls() {
    if (!selectedField || !textStyles[selectedField]) return;
    
    const styles = textStyles[selectedField];
    
    document.getElementById('fontSelect').value = styles.font || 'Inter';
    document.getElementById('fontSize').value = styles.size;
    document.getElementById('fontSizeRange').value = styles.size;
    document.getElementById('textAlign').value = styles.align;
    document.getElementById('textColor').value = styles.color;
    document.querySelector('.color-value').textContent = styles.color;
    
    // Update alignment buttons
    document.querySelectorAll('.align-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.align === styles.align);
    });
}

function updatePreview() {
    if (!templateImg.src || !canvas.width || !canvas.height) return;
    
    // Clear and draw template
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(templateImg, 0, 0, canvas.width, canvas.height);
    
    const sampleData = excelData.length > 0 ? excelData[0] : {};
    
    // Draw all selected fields
    for (let field of selectedFields) {
        const styles = textStyles[field];
        const pos = positions[field];
        const text = sampleData[field] || field;
        
        // Set text properties
        ctx.font = `${styles.size}px "${styles.font}"`;
        ctx.textAlign = styles.align;
        ctx.fillStyle = styles.color;
        
        // Draw text
        ctx.fillText(text, pos.x, pos.y);
        
        // Draw selection indicator
        if (field === selectedField) {
            const metrics = ctx.measureText(text);
            const textWidth = metrics.width;
            const textHeight = styles.size;
            
            let textX = pos.x;
            if (styles.align === 'center') textX -= textWidth / 2;
            else if (styles.align === 'right') textX -= textWidth;
            
            // Draw selection box
            ctx.strokeStyle = '#6366f1';
            ctx.lineWidth = 2;
            ctx.setLineDash([5, 5]);
            ctx.strokeRect(textX - 5, pos.y - textHeight - 5, textWidth + 10, textHeight + 10);
            ctx.setLineDash([]);
            
            // Draw position indicator
            ctx.fillStyle = '#6366f1';
            ctx.beginPath();
            ctx.arc(pos.x, pos.y, 4, 0, 2 * Math.PI);
            ctx.fill();
        }
    }
}

// Certificate generation
function generateCertificates() {
    if (!excelData.length) {
        alert('Please upload an Excel file first!');
        return;
    }
    if (!templateImg.src) {
        alert('Please upload a template image first!');
        return;
    }
    if (selectedFields.size === 0) {
        alert('Please select at least one field to display!');
        return;
    }

    // Show loading state
    const btn = event.target;
    const originalContent = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
    btn.disabled = true;

    const zip = new JSZip();
    let completedCertificates = 0;

    excelData.forEach((data, index) => {
        // Clear and draw template
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(templateImg, 0, 0, canvas.width, canvas.height);

        // Draw text for selected fields
        for (let field of selectedFields) {
            const styles = textStyles[field];
            const pos = positions[field];
            const text = data[field] || 'N/A';
            
            ctx.font = `${styles.size}px "${styles.font}"`;
            ctx.textAlign = styles.align;
            ctx.fillStyle = styles.color;
            ctx.fillText(text, pos.x, pos.y);
        }

        // Generate filename
        const fileName = Array.from(selectedFields)
            .slice(0, 2) // Use first 2 fields for filename
            .map(f => (data[f] || 'NA').toString().replace(/[^a-zA-Z0-9]/g, '_'))
            .join('_') + '_certificate.png';

        // Add to ZIP
        const dataURL = canvas.toDataURL('image/png');
        const base64Data = dataURL.replace(/^data:image\/png;base64,/, '');
        zip.file(fileName, base64Data, { base64: true });

        completedCertificates++;
        
        // Update progress
        btn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> Generating... ${completedCertificates}/${excelData.length}`;
        
        if (completedCertificates === excelData.length) {
            zip.generateAsync({ type: 'blob' }).then(function(blob) {
                saveAs(blob, 'certificates.zip');
                
                // Reset button
                btn.innerHTML = originalContent;
                btn.disabled = false;
                
                // Show success message
                alert(`Successfully generated ${excelData.length} certificates!`);
                
                // Restore preview
                updatePreview();
            }).catch(function(error) {
                console.error('Error generating ZIP:', error);
                alert('An error occurred while generating the ZIP file.');
                btn.innerHTML = originalContent;
                btn.disabled = false;
            });
        }
    });
}

// Patch Notes Functionality
function checkFirstVisit() {
    const hasSeenPatchNotes = localStorage.getItem('certgen_patch_notes_v2');
    const patchNotesBtn = document.querySelector('.patch-notes-btn');
    
    if (!hasSeenPatchNotes) {
        // Show patch notes automatically for first-time users
        setTimeout(() => {
            showPatchNotes();
        }, 1000);
    } else {
        // Remove the notification dot if already seen
        patchNotesBtn.classList.add('seen');
    }
}

function showPatchNotes() {
    const overlay = document.getElementById('patchNotesOverlay');
    overlay.classList.add('show');
    document.body.style.overflow = 'hidden';
    
    // Mark as seen
    const patchNotesBtn = document.querySelector('.patch-notes-btn');
    patchNotesBtn.classList.add('seen');
}

function closePatchNotes() {
    const overlay = document.getElementById('patchNotesOverlay');
    const dontShowAgain = document.getElementById('dontShowAgain');
    
    overlay.classList.remove('show');
    document.body.style.overflow = '';
    
    // Save preference if checkbox is checked
    if (dontShowAgain.checked) {
        localStorage.setItem('certgen_patch_notes_v2', 'seen');
    }
    
    // Always mark the current session as seen (removes the dot)
    localStorage.setItem('certgen_patch_notes_v2_session', 'seen');
}

// Close patch notes when clicking outside the modal
document.getElementById('patchNotesOverlay').addEventListener('click', function(e) {
    if (e.target === this) {
        closePatchNotes();
    }
});

// Close patch notes with Escape key
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        const overlay = document.getElementById('patchNotesOverlay');
        if (overlay.classList.contains('show')) {
            closePatchNotes();
        }
        
        // Close theme dropdown
        const themeDropdown = document.getElementById('themeDropdown');
        if (themeDropdown.classList.contains('show')) {
            themeDropdown.classList.remove('show');
        }
    }
});

// Theme Management
function initializeTheme() {
    const savedTheme = localStorage.getItem('certgen_theme') || 'system';
    applyTheme(savedTheme);
    updateThemeIcon(savedTheme);
    
    // Listen for system theme changes
    if (window.matchMedia) {
        window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
            const currentTheme = localStorage.getItem('certgen_theme') || 'system';
            if (currentTheme === 'system') {
                applyTheme('system');
            }
        });
    }
    
    // Close theme dropdown when clicking outside
    document.addEventListener('click', function(e) {
        const themeDropdown = document.getElementById('themeDropdown');
        const themeBtn = document.querySelector('.theme-btn');
        
        if (!themeBtn.contains(e.target) && !themeDropdown.contains(e.target)) {
            themeDropdown.classList.remove('show');
        }
    });
}

function toggleTheme() {
    const themeDropdown = document.getElementById('themeDropdown');
    themeDropdown.classList.toggle('show');
}

function setTheme(theme) {
    localStorage.setItem('certgen_theme', theme);
    applyTheme(theme);
    updateThemeIcon(theme);
    
    // Close dropdown
    document.getElementById('themeDropdown').classList.remove('show');
    
    // Update active state in dropdown
    document.querySelectorAll('.theme-option').forEach(option => {
        option.classList.remove('active');
    });
    document.querySelector(`[onclick="setTheme('${theme}')"]`).classList.add('active');
}

function applyTheme(theme) {
    const html = document.documentElement;
    
    if (theme === 'dark') {
        html.setAttribute('data-theme', 'dark');
    } else if (theme === 'light') {
        html.removeAttribute('data-theme');
    } else if (theme === 'system') {
        // Use system preference
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            html.setAttribute('data-theme', 'dark');
        } else {
            html.removeAttribute('data-theme');
        }
    }
}

function updateThemeIcon(theme) {
    const themeIcon = document.getElementById('themeIcon');
    const isDark = theme === 'dark' || (theme === 'system' && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches);
    
    if (theme === 'system') {
        themeIcon.className = 'fas fa-desktop';
    } else if (isDark) {
        themeIcon.className = 'fas fa-moon';
    } else {
        themeIcon.className = 'fas fa-sun';
    }
    
    // Update active state in dropdown
    document.querySelectorAll('.theme-option').forEach(option => {
        option.classList.remove('active');
    });
    document.querySelector(`[onclick="setTheme('${theme}')"]`)?.classList.add('active');
}