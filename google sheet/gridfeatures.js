// gridFeatures.js - Complete implementation for spreadsheet functionality

// Variables for drag and selection functionality
let isDragging = false;
let selectionStart = null;
let selectionEnd = null;
let selectedCells = [];
let draggedData = null;
let dragSourceCell = null;
let resizingColumn = null;
let resizingRow = null;
let initialSize = 0;
let initialMousePos = 0;
let activeCell = null; // Track the currently active cell

// Initialize grid features when the document is ready
document.addEventListener('DOMContentLoaded', function() {
    initializeGridFeatures();
});

function initializeGridFeatures() {
    // Set up cell focus tracking
    document.querySelectorAll('.cell-focus').forEach(cell => {
        cell.addEventListener('focus', function() {
            activeCell = this;
        });
    });
    
    // Add event listeners for cell selection
    document.querySelectorAll('.cell-focus').forEach(cell => {
        cell.addEventListener('mousedown', startSelection);
    });
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', endSelection);
    
    // Initialize drag and drop functionality
    initializeDragAndDrop();
    
    // Add column resize handlers
    addColumnResizeHandlers();
    
    // Add row resize handlers
    addRowResizeHandlers();
    
    // Add context menu for row/column operations
    addContextMenu();
    
    // Add toolbar buttons for row/column operations
    addRowColumnButtons();
}

// Selection functionality
function startSelection(event) {
    // Only start selection on left mouse button
    if (event.button !== 0) return;
    
    selectionStart = event.target.id;
    isDragging = true;
    clearSelection();
    
    // If shift key is pressed, extend selection from active cell
    if (event.shiftKey && activeCell) {
        selectionStart = activeCell.id;
        updateSelection(event.target.id);
    } else {
        // Mark the starting cell as selected
        highlightCell(event.target.id);
    }
    
    // Prevent text selection during drag
    document.body.style.userSelect = 'none';
    
    // If this is the start of a drag operation, prevent default to allow custom drag
    if (event.target.classList.contains('selected')) {
        event.preventDefault();
    }
}

function handleMouseMove(event) {
    if (!isDragging) return;
    
    // Find the cell element under the cursor
    const cellElement = findCellElementFromPoint(event.clientX, event.clientY);
    
    if (cellElement && cellElement.classList.contains('cell-focus')) {
        selectionEnd = cellElement.id;
        updateSelection(selectionEnd);
    }
}

function endSelection() {
    if (!isDragging) return;
    
    isDragging = false;
    document.body.style.userSelect = '';
    
    // If only one cell is selected, make it the active cell
    if (selectedCells.length === 1) {
        const cell = document.getElementById(selectedCells[0]);
        if (cell) {
            cell.focus();
            activeCell = cell;
        }
    }
}

function findCellElementFromPoint(x, y) {
    // Get all elements at the point
    const elements = document.elementsFromPoint(x, y);
    
    // Find the first element with cell-focus class
    for (const element of elements) {
        if (element.classList.contains('cell-focus')) {
            return element;
        }
    }
    
    return null;
}

function updateSelection(endCellId) {
    clearSelection();
    
    const startColChar = selectionStart.charAt(0);
    const startRow = parseInt(selectionStart.substring(1));
    const endColChar = endCellId.charAt(0);
    const endRow = parseInt(endCellId.substring(1));
    
    // Determine selection rectangle bounds
    const startColIndex = startColChar.charCodeAt(0);
    const endColIndex = endColChar.charCodeAt(0);
    
    const minCol = Math.min(startColIndex, endColIndex);
    const maxCol = Math.max(startColIndex, endColIndex);
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    
    // Highlight all cells in the selection rectangle
    for (let colIndex = minCol; colIndex <= maxCol; colIndex++) {
        const colChar = String.fromCharCode(colIndex);
        for (let row = minRow; row <= maxRow; row++) {
            const cellId = colChar + row;
            highlightCell(cellId);
        }
    }
}

function highlightCell(cellId) {
    const cell = document.getElementById(cellId);
    if (cell) {
        cell.classList.add('selected');
        selectedCells.push(cellId);
    }
}

function clearSelection() {
    // Remove selected class from all previously selected cells
    selectedCells.forEach(cellId => {
        const cell = document.getElementById(cellId);
        if (cell) {
            cell.classList.remove('selected');
        }
    });
    
    selectedCells = [];
}

// Drag and drop functionality
function initializeDragAndDrop() {
    document.querySelectorAll('.cell-focus').forEach(cell => {
        cell.setAttribute('draggable', 'true');
        cell.addEventListener('dragstart', handleDragStart);
        cell.addEventListener('dragover', handleDragOver);
        cell.addEventListener('drop', handleDrop);
        cell.addEventListener('dragend', handleDragEnd);
    });
}

function handleDragStart(event) {
    // If no cells are selected, select this one
    if (selectedCells.length === 0) {
        clearSelection();
        highlightCell(event.target.id);
    }
    
    // Only allow drag if cells are selected
    if (!event.target.classList.contains('selected')) {
        clearSelection();
        highlightCell(event.target.id);
    }
    
    dragSourceCell = event.target.id;
    
    // Create a data object containing all selected cells' data
    draggedData = {};
    selectedCells.forEach(cellId => {
        // Ensure activeSheetObject exists and has the cell data
        if (!activeSheetObject) {
            // Initialize if not defined
            window.activeSheetObject = {};
        }
        
        if (!activeSheetObject[cellId]) {
            // Initialize cell if not defined
            activeSheetObject[cellId] = {
                content: document.getElementById(cellId).innerText || '',
                fontFamily_data: 'Arial',
                fontSize_data: 16,
                isBold: false,
                isItalic: false,
                isUnderlined: false,
                textAlign: 'left',
                color: '#000000',
                backgroundColor: '#ffffff'
            };
        }
        
        draggedData[cellId] = {
            content: activeSheetObject[cellId].content || document.getElementById(cellId).innerText || '',
            formatting: {
                fontFamily: activeSheetObject[cellId].fontFamily_data || 'Arial',
                fontSize: activeSheetObject[cellId].fontSize_data || 16,
                isBold: activeSheetObject[cellId].isBold || false,
                isItalic: activeSheetObject[cellId].isItalic || false,
                isUnderlined: activeSheetObject[cellId].isUnderlined || false,
                textAlign: activeSheetObject[cellId].textAlign || 'left',
                color: activeSheetObject[cellId].color || '#000000',
                backgroundColor: activeSheetObject[cellId].backgroundColor || '#ffffff'
            }
        };
    });
    
    // Set drag effect and image
    event.dataTransfer.effectAllowed = 'move';
    
    // Create a custom drag image showing number of cells
    const dragImage = document.createElement('div');
    dragImage.textContent = `${selectedCells.length} cell${selectedCells.length > 1 ? 's' : ''}`;
    dragImage.style.backgroundColor = '#d0d0d0';
    dragImage.style.padding = '4px 8px';
    dragImage.style.borderRadius = '4px';
    dragImage.style.position = 'absolute';
    dragImage.style.top = '-1000px';
    document.body.appendChild(dragImage);
    
    event.dataTransfer.setDragImage(dragImage, 0, 0);
    
    // Store cells being dragged as JSON
    event.dataTransfer.setData('text/plain', JSON.stringify({
        cells: selectedCells,
        sourceCell: dragSourceCell
    }));
    
    // Clean up the drag image after a short delay
    setTimeout(() => {
        document.body.removeChild(dragImage);
    }, 100);
}

function handleDragOver(event) {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    
    // Highlight the target cell
    event.target.classList.add('drag-over');
}

function handleDrop(event) {
    event.preventDefault();
    
    // Remove any drag-over highlight
    document.querySelectorAll('.drag-over').forEach(cell => {
        cell.classList.remove('drag-over');
    });
    
    if (!draggedData) return;
    
    const targetCellId = event.target.id;
    
    // Calculate offset between source and target
    const sourceCol = dragSourceCell.charCodeAt(0);
    const sourceRow = parseInt(dragSourceCell.substring(1));
    const targetCol = targetCellId.charCodeAt(0);
    const targetRow = parseInt(targetCellId.substring(1));
    
    const colOffset = targetCol - sourceCol;
    const rowOffset = targetRow - sourceRow;
    
    // Create a copy of the data to avoid modifying during iteration
    const draggedDataCopy = { ...draggedData };
    
    // Clear selection to prepare for new one
    clearSelection();
    
    // Ensure activeSheetObject is initialized
    if (!window.activeSheetObject) {
        window.activeSheetObject = {};
    }
    
    // Define initialCellState if not already defined
    if (!window.initialCellState) {
        window.initialCellState = {
            content: '',
            fontFamily_data: 'Arial',
            fontSize_data: 16,
            isBold: false,
            isItalic: false,
            isUnderlined: false,
            textAlign: 'left',
            color: '#000000',
            backgroundColor: '#ffffff'
        };
    }
    
    // Move all selected cells to their new positions
    Object.keys(draggedDataCopy).forEach(sourceCellId => {
        const sourceColChar = sourceCellId.charAt(0);
        const sourceRowNum = parseInt(sourceCellId.substring(1));
        
        const newColCharCode = sourceColChar.charCodeAt(0) + colOffset;
        // Check if new column is within range (A-Z)
        if (newColCharCode < 65 || newColCharCode > 90) return;
        
        const newRowNum = sourceRowNum + rowOffset;
        // Check if new row is within range (1-100)
        if (newRowNum < 1 || newRowNum > 100) return;
        
        const newCellId = String.fromCharCode(newColCharCode) + newRowNum;
        
        // Initialize target cell if needed
        if (!activeSheetObject[newCellId]) {
            activeSheetObject[newCellId] = { ...initialCellState };
        }
        
        // Copy data to new cell
        activeSheetObject[newCellId].content = draggedDataCopy[sourceCellId].content;
        
        // Copy formatting
        const formatting = draggedDataCopy[sourceCellId].formatting;
        activeSheetObject[newCellId].fontFamily_data = formatting.fontFamily;
        activeSheetObject[newCellId].fontSize_data = formatting.fontSize;
        activeSheetObject[newCellId].isBold = formatting.isBold;
        activeSheetObject[newCellId].isItalic = formatting.isItalic;
        activeSheetObject[newCellId].isUnderlined = formatting.isUnderlined;
        activeSheetObject[newCellId].textAlign = formatting.textAlign;
        activeSheetObject[newCellId].color = formatting.color;
        activeSheetObject[newCellId].backgroundColor = formatting.backgroundColor;
        
        // Update UI for the new cell
        const targetCell = document.getElementById(newCellId);
        if (targetCell) {
            targetCell.innerText = activeSheetObject[newCellId].content;
            targetCell.style.fontFamily = formatting.fontFamily;
            targetCell.style.fontSize = formatting.fontSize + 'px';
            targetCell.style.fontWeight = formatting.isBold ? 'bold' : 'normal';
            targetCell.style.fontStyle = formatting.isItalic ? 'italic' : 'normal';
            targetCell.style.textDecoration = formatting.isUnderlined ? 'underline' : 'none';
            targetCell.style.textAlign = formatting.textAlign;
            targetCell.style.color = formatting.color;
            targetCell.style.backgroundColor = formatting.backgroundColor;
            
            // Add to new selection
            highlightCell(newCellId);
        }
    });
    
    // After drop, check if this was a copy operation (with Ctrl key)
    // If not, clear source cells
    if (!event.ctrlKey) {
        Object.keys(draggedDataCopy).forEach(sourceCellId => {
            if (activeSheetObject[sourceCellId]) {
                activeSheetObject[sourceCellId].content = '';
            }
            
            const sourceCell = document.getElementById(sourceCellId);
            if (sourceCell) {
                sourceCell.innerText = '';
            }
        });
    }
}

function handleDragEnd(event) {
    // Remove any drag-over highlights
    document.querySelectorAll('.drag-over').forEach(cell => {
        cell.classList.remove('drag-over');
    });
    
    draggedData = null;
    dragSourceCell = null;
}

// Column resize functionality
function addColumnResizeHandlers() {
    // Add resize handles to column headers
    const headerCells = document.querySelectorAll('.grid-header-col');
    
    headerCells.forEach(headerCell => {
        if (headerCell.id && headerCell.id.length === 1) { // Only actual column headers (A-Z)
            // Remove any existing resize handle to avoid duplicates
            const existingHandle = headerCell.querySelector('.column-resize-handle');
            if (existingHandle) {
                existingHandle.remove();
            }
            
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'column-resize-handle';
            resizeHandle.style.position = 'absolute';
            resizeHandle.style.right = '0';
            resizeHandle.style.top = '0';
            resizeHandle.style.width = '5px';
            resizeHandle.style.height = '100%';
            resizeHandle.style.cursor = 'col-resize';
            resizeHandle.style.zIndex = '10';
            
            headerCell.style.position = 'relative';
            headerCell.appendChild(resizeHandle);
            
            resizeHandle.addEventListener('mousedown', startColumnResize);
        }
    });
}

function startColumnResize(event) {
    event.preventDefault();
    event.stopPropagation();
    
    resizingColumn = event.target.parentElement.id;
    initialMousePos = event.clientX;
    
    // Get initial width
    const columnCells = document.querySelectorAll(`[id^="${resizingColumn}"]`);
    if (columnCells.length > 0) {
        initialSize = columnCells[0].offsetWidth;
    }
    
    document.addEventListener('mousemove', resizeColumn);
    document.addEventListener('mouseup', endColumnResize);
}

function resizeColumn(event) {
    if (!resizingColumn) return;
    
    const delta = event.clientX - initialMousePos;
    const newWidth = Math.max(40, initialSize + delta); // Minimum 40px width
    
    // Update column header width
    const columnHeader = document.getElementById(resizingColumn);
    if (columnHeader) {
        columnHeader.style.width = `${newWidth}px`;
    }
    
    // Update all cells in the column
    const columnCells = document.querySelectorAll(`[id^="${resizingColumn}"]`);
    columnCells.forEach(cell => {
        cell.style.width = `${newWidth}px`;
    });
}

function endColumnResize() {
    resizingColumn = null;
    document.removeEventListener('mousemove', resizeColumn);
    document.removeEventListener('mouseup', endColumnResize);
}

// Row resize functionality
function addRowResizeHandlers() {
    // Add resize handles to row headers
    const rowHeaders = document.querySelectorAll('.grid > .row > .grid-cell:first-child');
    
    rowHeaders.forEach(rowHeader => {
        if (rowHeader.id && !isNaN(parseInt(rowHeader.id))) { // Only row headers (1-100)
            // Remove existing handle to avoid duplicates
            const existingHandle = rowHeader.querySelector('.row-resize-handle');
            if (existingHandle) {
                existingHandle.remove();
            }
            
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'row-resize-handle';
            resizeHandle.style.position = 'absolute';
            resizeHandle.style.left = '0';
            resizeHandle.style.bottom = '0';
            resizeHandle.style.width = '100%';
            resizeHandle.style.height = '5px';
            resizeHandle.style.cursor = 'row-resize';
            resizeHandle.style.zIndex = '10';
            
            rowHeader.style.position = 'relative';
            rowHeader.appendChild(resizeHandle);
            
            resizeHandle.addEventListener('mousedown', startRowResize);
        }
    });
}

function startRowResize(event) {
    event.preventDefault();
    event.stopPropagation();
    
    resizingRow = event.target.parentElement.id;
    initialMousePos = event.clientY;
    
    // Get initial height
    const rowCells = document.querySelectorAll(`.row:nth-child(${parseInt(resizingRow) + 1}) .grid-cell`);
    if (rowCells.length > 0) {
        initialSize = rowCells[0].offsetHeight;
    }
    
    document.addEventListener('mousemove', resizeRow);
    document.addEventListener('mouseup', endRowResize);
}

function resizeRow(event) {
    if (!resizingRow) return;
    
    const delta = event.clientY - initialMousePos;
    const newHeight = Math.max(20, initialSize + delta); // Minimum 20px height
    
    // Update all cells in the row
    const rowIndex = parseInt(resizingRow);
    const rowCells = document.querySelectorAll(`.row:nth-child(${rowIndex + 1}) .grid-cell`);
    rowCells.forEach(cell => {
        cell.style.height = `${newHeight}px`;
    });
}

function endRowResize() {
    resizingRow = null;
    document.removeEventListener('mousemove', resizeRow);
    document.removeEventListener('mouseup', endRowResize);
}

// Context menu for row/column operations
function addContextMenu() {
    // Create context menu element
    const contextMenu = document.createElement('div');
    contextMenu.id = 'grid-context-menu';
    contextMenu.style.position = 'absolute';
    contextMenu.style.backgroundColor = '#ffffff';
    contextMenu.style.border = '1px solid #ccc';
    contextMenu.style.boxShadow = '2px 2px 5px rgba(0,0,0,0.2)';
    contextMenu.style.padding = '5px 0';
    contextMenu.style.zIndex = '1000';
    contextMenu.style.display = 'none';
    
    // Add menu items
    const menuItems = [
        { text: 'Insert row above', action: insertRowAbove },
        { text: 'Insert row below', action: addRow },
        { text: 'Delete row', action: deleteRow },
        { text: 'Insert column left', action: insertColumnLeft },
        { text: 'Insert column right', action: addColumn },
        { text: 'Delete column', action: deleteColumn },
        { text: 'Clear contents', action: clearContents },
        { text: 'Cut', action: cutSelection },
        { text: 'Copy', action: copySelection },
        { text: 'Paste', action: pasteSelection }
    ];
    
    menuItems.forEach(item => {
        const menuItem = document.createElement('div');
        menuItem.className = 'context-menu-item';
        menuItem.style.padding = '5px 10px';
        menuItem.style.cursor = 'pointer';
        menuItem.innerText = item.text;
        menuItem.addEventListener('mouseenter', () => {
            menuItem.style.backgroundColor = '#f0f0f0';
        });
        menuItem.addEventListener('mouseleave', () => {
            menuItem.style.backgroundColor = '';
        });
        menuItem.addEventListener('click', () => {
            item.action();
            contextMenu.style.display = 'none';
        });
        
        contextMenu.appendChild(menuItem);
    });
    
    document.body.appendChild(contextMenu);
    
    // Show context menu on right click
    document.addEventListener('contextmenu', function(event) {
        const target = event.target;
        
        // Only show context menu for grid cells
        if (target.classList.contains('cell-focus') || 
            target.classList.contains('grid-header-col') || 
            (target.classList.contains('grid-cell') && target.parentElement.classList.contains('row'))) {
            
            event.preventDefault();
            
            // Position the menu
            contextMenu.style.left = `${event.pageX}px`;
            contextMenu.style.top = `${event.pageY}px`;
            contextMenu.style.display = 'block';
            
            // If no cell is selected or clicked cell is not selected, clear and select the clicked cell
            if (target.classList.contains('cell-focus') && !target.classList.contains('selected')) {
                clearSelection();
                highlightCell(target.id);
                activeCell = target;
            }
        }
    });
    
    // Hide context menu when clicking elsewhere
    document.addEventListener('click', function() {
        contextMenu.style.display = 'none';
    });
}

// Add/delete row and column functionality
function addRowColumnButtons() {
    // Get or create toolbar
    let toolbar = document.querySelector('.cell-styles');
    if (!toolbar) {
        toolbar = document.createElement('div');
        toolbar.className = 'cell-styles';
        document.querySelector('.menu-bar').appendChild(toolbar);
    }
    
    // Check if buttons already exist
    if (toolbar.querySelector('.row-col-operations')) {
        return;
    }
    
    const rowColContainer = document.createElement('div');
    rowColContainer.className = 'row-col-operations';
    rowColContainer.style.display = 'flex';
    rowColContainer.style.gap = '5px';
    rowColContainer.style.marginLeft = '15px';
    
    // Add row button
    const addRowBtn = document.createElement('button');
    addRowBtn.innerHTML = '+ Row';
    addRowBtn.title = 'Add row below';
    addRowBtn.className = 'row-col-btn';
    addRowBtn.style.padding = '2px 8px';
    addRowBtn.style.cursor = 'pointer';
    addRowBtn.onclick = addRow;
    
    // Delete row button
    const deleteRowBtn = document.createElement('button');
    deleteRowBtn.innerHTML = '- Row';
    deleteRowBtn.title = 'Delete selected row';
    deleteRowBtn.className = 'row-col-btn';
    deleteRowBtn.style.padding = '2px 8px';
    deleteRowBtn.style.cursor = 'pointer';
    deleteRowBtn.onclick = deleteRow;
    
    // Add column button
    const addColBtn = document.createElement('button');
    addColBtn.innerHTML = '+ Col';
    addColBtn.title = 'Add column to the right';
    addColBtn.className = 'row-col-btn';
    addColBtn.style.padding = '2px 8px';
    addColBtn.style.cursor = 'pointer';
    addColBtn.onclick = addColumn;
    
    // Delete column button
    const deleteColBtn = document.createElement('button');
    deleteColBtn.innerHTML = '- Col';
    deleteColBtn.title = 'Delete selected column';
    deleteColBtn.className = 'row-col-btn';
    deleteColBtn.style.padding = '2px 8px';
    deleteColBtn.style.cursor = 'pointer';
    deleteColBtn.onclick = deleteColumn;
    
    // Add buttons to container
    rowColContainer.appendChild(addRowBtn);
    rowColContainer.appendChild(deleteRowBtn);
    rowColContainer.appendChild(addColBtn);
    rowColContainer.appendChild(deleteColBtn);
    
    // Add container to toolbar
    toolbar.appendChild(rowColContainer);
}

// Insert a row above the active cell's row
function insertRowAbove() {
    if (!activeCell) {
        alert('Please select a cell first');
        return;
    }
    
    const rowToInsertBefore = parseInt(activeCell.id.substring(1));
    
    // Max row check
    if (document.querySelectorAll('.row').length >= 100) {
        alert('Maximum number of rows reached (100)');
        return;
    }
    
    // Shift all rows below this one in the data
    shiftRowsDown(rowToInsertBefore - 1);
    
    // Update UI: Add a new row to the DOM
    insertRowInDOM(rowToInsertBefore - 1);
    
    // Reattach events
    reattachCellEvents();
    
    // Reinitialize features
    initializeGridFeatures();
}

function addRow() {
    // Get the active row number
    let rowToInsertAfter = 1;
    if (activeCell) {
        rowToInsertAfter = parseInt(activeCell.id.substring(1));
    }
    
    // Max row check
    if (document.querySelectorAll('.row').length >= 100) {
        alert('Maximum number of rows reached (100)');
        return;
    }
    
    // Shift all rows below this one in the data
    shiftRowsDown(rowToInsertAfter);
    
    // Update UI: Add a new row to the DOM
    insertRowInDOM(rowToInsertAfter);
    
    // Reattach events
    reattachCellEvents();
    
    // Reinitialize features
    initializeGridFeatures();
}

function insertRowInDOM(rowAfter) {
    const rowIndex = rowAfter + 1;
    const existingRows = document.querySelectorAll('.row');
    
    // Create new row element
    const newRow = document.createElement('div');
    newRow.className = 'row';
    
    // Create row header
    const rowHeader = document.createElement('div');
    rowHeader.className = 'grid-cell';
    rowHeader.id = rowIndex.toString();
    rowHeader.innerText = rowIndex;
    newRow.appendChild(rowHeader);
    
    // Create cells for this row
    for (let j = 65; j <= 90; j++) {
        const cell = document.createElement('div');
        cell.className = 'grid-cell cell-focus';
        cell.id = String.fromCharCode(j) + rowIndex;
        cell.contentEditable = true;
        newRow.appendChild(cell);
    }
    
    // Insert the new row after the target row
    const grid = document.querySelector('.grid');
    if (rowAfter < existingRows.length) {
        grid.insertBefore(newRow, existingRows[rowAfter].nextSibling);
    } else {
        grid.appendChild(newRow);
    }
    
    // Update row numbers for all rows below
    for (let i = rowIndex + 1; i <= existingRows.length + 1; i++) {
        const rowHeader = document.querySelector(`.row:nth-child(${i + 1}) .grid-cell:first-child`);
        if (rowHeader) {
            rowHeader.innerText = i;
            rowHeader.id = i.toString();
        }
        
        // Update cell IDs in this row
        for (let j = 65; j <= 90; j++) {
            const colChar = String.fromCharCode(j);
            const oldCellId = colChar + (i - 1);
            const newCellId = colChar + i;
            const cell = document.getElementById(oldCellId);
            if (cell) {
                cell.id = newCellId;
            }
        }
    }
}

function shiftRowsDown(rowAfter) {
    // Ensure activeSheetObject is initialized
    if (!window.activeSheetObject) {
        window.activeSheetObject = {};
    }
    
    // Define initialCellState if not already defined
    if (!window.initialCellState) {
        window.initialCellState = {
            content: '',
            fontFamily_data: 'Arial',
            fontSize_data: 16,
            isBold: false,
            isItalic: false,
            isUnderlined: false,
            textAlign: 'left',
            color: '#000000',
            backgroundColor: '#ffffff'
        };
    }
    
    // Start from the last row and move up to avoid overwriting
    for (let i = 100; i > rowAfter; i--) {
        for (let j = 65; j <= 90; j++) {
            const colChar = String.fromCharCode(j);
            const oldCellId = colChar + i;
            const newCellId = colChar + (i + 1);
            
            // Copy data from current row to the row below
            if (activeSheetObject[oldCellId]) {
                activeSheetObject[newCellId] = { ...activeSheetObject[oldCellId] };
            } else {
                activeSheetObject[newCellId] = { ...initialCellState };
            }
        }
    }
    
    // Initialize the new row with empty cells
    for (let j = 65; j <= 90; j++) {
        const colChar = String.fromCharCode(j);
        const newCellId = colChar + (rowAfter + 1);
        activeSheetObject[newCellId] = { ...initialCellState };
    }
}

