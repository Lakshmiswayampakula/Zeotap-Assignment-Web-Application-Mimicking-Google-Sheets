// Data Quality Functions for Google Sheets Clone

// TRIM: Removes leading and trailing whitespace from a cell
function trimCell() {
    if (!activeCell) {
        alert("Please select a cell first");
        return;
    }
    
    const cellId = activeCell.id;
    const cellContent = activeSheetObject[cellId].content;
    
    if (cellContent) {
        const trimmedContent = cellContent.trim();
        activeCell.innerText = trimmedContent;
        activeSheetObject[cellId].content = trimmedContent;
        formula.value = trimmedContent;
    }
}

// UPPER: Converts the text in a cell to uppercase
function upperCell() {
    if (!activeCell) {
        alert("Please select a cell first");
        return;
    }
    
    const cellId = activeCell.id;
    const cellContent = activeSheetObject[cellId].content;
    
    if (cellContent) {
        const upperContent = cellContent.toUpperCase();
        activeCell.innerText = upperContent;
        activeSheetObject[cellId].content = upperContent;
        formula.value = upperContent;
    }
}

// LOWER: Converts the text in a cell to lowercase
function lowerCell() {
    if (!activeCell) {
        alert("Please select a cell first");
        return;
    }
    
    const cellId = activeCell.id;
    const cellContent = activeSheetObject[cellId].content;
    
    if (cellContent) {
        const lowerContent = cellContent.toLowerCase();
        activeCell.innerText = lowerContent;
        activeSheetObject[cellId].content = lowerContent;
        formula.value = lowerContent;
    }
}

// REMOVE_DUPLICATES: Removes duplicate rows from the selected row
function removeDuplicates() {
    if (!activeCell) {
        alert("Please select a cell first");
        return;
    }
    
    // Get the row number from the active cell
    const rowNum = activeCell.id.substring(1);
    
    // Find all cells in the same row
    const rowCells = {};
    for (let j = 65; j <= 90; j++) {
        const key = String.fromCharCode(j) + rowNum;
        if (activeSheetObject[key]) {
            rowCells[key] = activeSheetObject[key].content;
        }
    }
    
    // Find duplicate rows
    const duplicateRows = [];
    for (let i = 1; i <= 100; i++) {
        if (i === parseInt(rowNum)) continue; // Skip the selected row
        
        let isDuplicate = true;
        for (let j = 65; j <= 90; j++) {
            const currentKey = String.fromCharCode(j) + i;
            const selectedKey = String.fromCharCode(j) + rowNum;
            
            // If any cell content doesn't match, it's not a duplicate
            if (activeSheetObject[currentKey]?.content !== activeSheetObject[selectedKey]?.content) {
                isDuplicate = false;
                break;
            }
        }
        
        if (isDuplicate && Object.values(rowCells).some(content => content)) {
            duplicateRows.push(i);
        }
    }
    
    // Remove duplicate rows
    for (const row of duplicateRows) {
        for (let j = 65; j <= 90; j++) {
            const key = String.fromCharCode(j) + row;
            if (activeSheetObject[key]) {
                activeSheetObject[key].content = "";
                const cell = document.getElementById(key);
                if (cell) {
                    cell.innerText = "";
                }
            }
        }
    }
    
    if (duplicateRows.length > 0) {
        alert(`Removed ${duplicateRows.length} duplicate rows.`);
    } else {
        alert("No duplicate rows found.");
    }
}

// FIND_AND_REPLACE: Allows users to find and replace specific text within the selected cell
function findAndReplace() {
    if (!activeCell) {
        alert("Please select a cell first");
        return;
    }
    
    // Create a modal dialog for find and replace
    const modal = document.createElement('div');
    modal.className = 'find-replace-modal';
    modal.innerHTML = `
        <div class="find-replace-content">
            <h3>Find and Replace</h3>
            <div class="input-group">
                <label for="find-text">Find:</label>
                <input type="text" id="find-text">
            </div>
            <div class="input-group">
                <label for="replace-text">Replace with:</label>
                <input type="text" id="replace-text">
            </div>
            <div class="button-group">
                <button id="replace-btn">Replace in this cell</button>
                <button id="replace-all-btn">Replace in all cells</button>
                <button id="cancel-btn">Cancel</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Add event listeners
    document.getElementById('replace-btn').addEventListener('click', () => {
        const findText = document.getElementById('find-text').value;
        const replaceText = document.getElementById('replace-text').value;
        
        if (!findText) {
            alert('Please enter text to find.');
            return;
        }
        
        const cellId = activeCell.id;
        const cellContent = activeSheetObject[cellId].content;
        
        if (cellContent && cellContent.includes(findText)) {
            const newContent = cellContent.replaceAll(findText, replaceText);
            activeCell.innerText = newContent;
            activeSheetObject[cellId].content = newContent;
            formula.value = newContent;
            
            alert(`Replaced in cell ${cellId}.`);
        } else {
            alert(`"${findText}" not found in cell ${cellId}.`);
        }
        
        modal.remove();
    });
    
    document.getElementById('replace-all-btn').addEventListener('click', () => {
        const findText = document.getElementById('find-text').value;
        const replaceText = document.getElementById('replace-text').value;
        
        if (!findText) {
            alert('Please enter text to find.');
            return;
        }
        
        let replacements = 0;
        
        // Replace in all cells of the active sheet
        for (const key in activeSheetObject) {
            const cell = activeSheetObject[key];
            if (cell.content && cell.content.includes(findText)) {
                const newContent = cell.content.replaceAll(findText, replaceText);
                cell.content = newContent;
                
                // Update the cell in the UI
                const cellElement = document.getElementById(key);
                if (cellElement) {
                    cellElement.innerText = newContent;
                }
                
                replacements++;
            }
        }
        
        alert(`Replaced ${replacements} occurrences.`);
        modal.remove();
    });
    
    document.getElementById('cancel-btn').addEventListener('click', () => {
        modal.remove();
    });
    
    // Focus the find input
    setTimeout(() => {
        document.getElementById('find-text').focus();
    }, 100);
    
    // Add CSS for the modal
    const style = document.createElement('style');
    style.innerHTML = `
        .find-replace-modal {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.5);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
        }
        
        .find-replace-content {
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            width: 350px;
        }
        
        .input-group {
            margin-bottom: 15px;
        }
        
        .input-group label {
            display: block;
            margin-bottom: 5px;
        }
        
        .input-group input {
            width: 100%;
            padding: 5px;
        }
        
        .button-group {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            flex-wrap: wrap;
        }
        
        .button-group button {
            padding: 5px 10px;
            cursor: pointer;
        }
    `;
    document.head.appendChild(style);
}