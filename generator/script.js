let templateImg = new Image();
let excelData = [];
let positions = {};
let textStyles = {};
let selectedFields = new Set();

const canvas = document.getElementById('preview');
const ctx = canvas.getContext('2d');
let isDragging = false;
let selectedField = null;
let offsetX, offsetY;
const MOVE_STEP = 5;
const MOVE_INTERVAL_TIME = 50; // Time between movements in milliseconds
let moveInterval = null; // To store the interval ID for continuous movement

const googleFonts = [
    'Arial', 'Times New Roman', 'Courier', 'Poppins', 'Roboto', 'Open Sans',
    'Lato', 'Montserrat', 'Merriweather', 'Oswald', 'Raleway', 'Source Sans Pro',
    'Playfair Display', 'Ubuntu', 'Nunito', 'Pacifico', 'Lobster', 'Bree Serif'
];

document.getElementById('excelInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) {
        console.error('No Excel file selected');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(sheet);
            console.log('Excel data loaded:', excelData);
            
            if (excelData.length === 0) {
                console.warn('Excel file is empty');
                return;
            }
 
            const headings = Object.keys(excelData[0]);
            positions = {};
            textStyles = {};
            selectedFields.clear();
    
            headings.forEach((field, index) => {
                positions[field] = { x: 50, y: 50 + (index * 50) };
                textStyles[field] = { font: 'Arial', size: 16, align: 'left', color: '#000000' };
                selectedFields.add(field);
            });

            populateFieldOptions(headings);
            updatePreview();
        } catch (error) {
            console.error('Error reading Excel file:', error);
        }
    };
    reader.onerror = function() {
        console.error('Error reading file');
    };
    reader.readAsArrayBuffer(file);
});

function populateFieldOptions(headings) {
    const fieldSelect = document.getElementById('fieldSelect');
    const fieldCheckboxes = document.getElementById('fieldCheckboxes');
    fieldSelect.innerHTML = '';
    fieldCheckboxes.innerHTML = '<p>Select fields to display:</p>';

    headings.forEach(field => {
        const option = document.createElement('option');
        option.value = field;
        option.textContent = field;
        fieldSelect.appendChild(option);
 
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
        label.appendChild(document.createTextNode(` ${field}`));
        fieldCheckboxes.appendChild(label);
        fieldCheckboxes.appendChild(document.createElement('br'));
    });

    selectedField = headings[0];
    updateStyleControls();
}

document.getElementById('templateInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) {
        console.error('No template image selected');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        templateImg.src = e.target.result;
        templateImg.onload = function() {
            canvas.width = templateImg.width;
            canvas.height = templateImg.height;
            console.log('Template loaded, canvas size:', canvas.width, 'x', canvas.height);
            updatePreview();
        };
        templateImg.onerror = function() {
            console.error('Error loading template image');
        };
    };
    reader.readAsDataURL(file);
});

canvas.addEventListener('mousedown', function(e) {
    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    for (let field in positions) {
        if (!selectedFields.has(field)) continue;
        const pos = positions[field];
        const grabDistance = Math.max(textStyles[field].size, 20);

        if (
            x >= pos.x - grabDistance &&
            x <= pos.x + grabDistance &&
            y >= pos.y - grabDistance &&
            y <= pos.y + grabDistance
        ) {
            isDragging = true;
            selectedField = field;
            offsetX = x - pos.x;
            offsetY = y - pos.y;
            document.getElementById('fieldSelect').value = field;
            updateStyleControls();
            console.log(`Dragging ${selectedField} at (${x}, ${y})`);
            break;
        }
    }
});

document.addEventListener('mousemove', function(e) {
    if (isDragging && selectedField) {
        const rect = canvas.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;

        positions[selectedField] = { x: x - offsetX, y: y - offsetY };
        updatePreview();
    }
});

document.addEventListener('mouseup', function() {
    if (isDragging) {
        console.log(`Stopped dragging ${selectedField}`);
    }
    isDragging = false;
});

function moveText(direction) {
    if (!selectedField) {
        alert('Please select a field first!');
        return;
    }

    const pos = positions[selectedField];
    
    // Function to move the field in the specified direction
    const move = () => {
        switch (direction) {
            case 'up':
                pos.y -= MOVE_STEP;
                break;
            case 'down':
                pos.y += MOVE_STEP;
                break;
            case 'left':
                pos.x -= MOVE_STEP;
                break;
            case 'right':
                pos.x += MOVE_STEP;
                break;
        }

        // Clamp position within canvas bounds
        pos.x = Math.max(0, Math.min(pos.x, canvas.width));
        pos.y = Math.max(0, Math.min(pos.y, canvas.height));

        console.log(`Moved ${selectedField} ${direction} to (${pos.x}, ${pos.y})`);
        updatePreview();
    };

    // Start moving immediately on mousedown
    move();

    // Start continuous movement
    moveInterval = setInterval(move, MOVE_INTERVAL_TIME);
}

// Stop movement when mouse is released or leaves the button
function stopMoving() {
    if (moveInterval) {
        clearInterval(moveInterval);
        moveInterval = null;
        console.log(`Stopped moving ${selectedField}`);
    }
}

// Add event listeners to directional buttons
document.querySelectorAll('.dpad-btn').forEach(button => {
    button.addEventListener('mousedown', (e) => {
        const direction = e.target.classList.contains('up') ? 'up' :
                          e.target.classList.contains('down') ? 'down' :
                          e.target.classList.contains('left') ? 'left' :
                          e.target.classList.contains('right') ? 'right' : null;
        if (direction) {
            moveText(direction);
        }
    });

    button.addEventListener('mouseup', stopMoving);
    button.addEventListener('mouseleave', stopMoving); // Stop if mouse leaves button while held
});

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
 
    if (selectedField) {
        const styles = textStyles[selectedField];
        fontSelect.value = styles.font || 'Arial';
    }
}

document.getElementById('fontSearch').addEventListener('input', function(e) {
    loadGoogleFonts(e.target.value);
});

function updateStyleControls() {
    if (!selectedField) selectedField = document.getElementById('fieldSelect').value;
    const styles = textStyles[selectedField];
    document.getElementById('fontSelect').value = styles.font || 'Arial';
    document.getElementById('fontSize').value = styles.size;
    document.getElementById('textAlign').value = styles.align;
    document.getElementById('textColor').value = styles.color;
}

document.getElementById('fieldSelect').addEventListener('change', function() {
    selectedField = this.value;
    updateStyleControls();
    updatePreview();
});

document.getElementById('fontSelect').addEventListener('change', function() {
    const field = selectedField || document.getElementById('fieldSelect').value;
    const selectedFont = this.value;
    textStyles[field].font = selectedFont;
 
    WebFont.load({
        google: {
            families: [selectedFont]
        },
        active: function() {
            updatePreview();
        },
        inactive: function() {
            console.warn(`Font ${selectedFont} failed to load. Using default (Arial).`);
            textStyles[field].font = 'Arial';
            updatePreview();
        }
    });
});

document.getElementById('fontSize').addEventListener('change', function() {
    const field = selectedField || document.getElementById('fieldSelect').value;
    textStyles[field].size = parseInt(this.value);
    updatePreview();
});

document.getElementById('textAlign').addEventListener('change', function() {
    const field = selectedField || document.getElementById('fieldSelect').value;
    textStyles[field].align = this.value;
    updatePreview();
});

document.getElementById('textColor').addEventListener('change', function() {
    const field = selectedField || document.getElementById('fieldSelect').value;
    textStyles[field].color = this.value;
    updatePreview();
});

function updatePreview() {
    if (!templateImg.src || !canvas.width || !canvas.height) {
        console.warn('No template loaded yet or canvas not initialized');
        return;
    }

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(templateImg, 0, 0);

    const sampleData = excelData.length > 0 ? excelData[0] : {};
    console.log('Preview data:', sampleData);

    for (let field of selectedFields) {
        const styles = textStyles[field];
        ctx.font = `${styles.size}px "${styles.font}"`;
        ctx.textAlign = styles.align;
        ctx.fillStyle = styles.color;
        const text = sampleData[field] || 'N/A';
        console.log(`Rendering ${field}: "${text}" at (${positions[field].x}, ${positions[field].y})`);
        ctx.fillText(text, positions[field].x, positions[field].y);
 
        ctx.fillStyle = 'red';
        ctx.beginPath();
        ctx.arc(positions[field].x, positions[field].y, 5, 0, 2 * Math.PI);
        ctx.fill();
    }
}

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

    const zip = new JSZip();
    let completedCertificates = 0;

    excelData.forEach((data, index) => {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(templateImg, 0, 0);
 
        for (let field of selectedFields) {
            const styles = textStyles[field];
            ctx.font = `${styles.size}px "${styles.font}"`;
            ctx.textAlign = styles.align;
            ctx.fillStyle = styles.color;
            ctx.fillText(data[field] || 'N/A', positions[field].x, positions[field].y);
        }
 
        const fileName = Array.from(selectedFields)
            .map(f => data[f] || 'NA')
            .join('_') + '_certificate.png';
        const dataURL = canvas.toDataURL('image/png');
        const base64Data = dataURL.replace(/^data:image\/png;base64,/, '');
        zip.file(fileName, base64Data, { base64: true });

        completedCertificates++;
         
        if (completedCertificates === excelData.length) {
            zip.generateAsync({ type: 'blob' }).then(function(blob) {
                saveAs(blob, 'certificates.zip');
                alert('Certificates generated and downloaded as ZIP successfully!');
            }).catch(function(error) {
                console.error('Error generating ZIP:', error);
                alert('An error occurred while generating the ZIP file.');
            });
        }
    });
}

document.addEventListener('DOMContentLoaded', function() {
    selectedField = null;
    updateStyleControls();
    loadGoogleFonts();
});