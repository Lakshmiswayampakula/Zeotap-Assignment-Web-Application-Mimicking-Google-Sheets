
function setFont(target){
    if(activeCell){
        let fontInput = target.value;
        console.log(fontInput);
        activeSheetObject[activeCell.id].fontFamily_data = fontInput;
        activeCell.style.fontFamily = fontInput;
        activeCell.focus();
    }
}
function setSize(target){
    if(activeCell){
        let sizeInput = target.value;
        activeSheetObject[activeCell.id].fontSize_data = sizeInput;
        activeCell.style.fontSize = sizeInput+'px';
        activeCell.focus();
    }
}

// bug fix
boldBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    toggleBold();
})
function toggleBold(){
    if(activeCell){
        if(!activeSheetObject[activeCell.id].isBold) {
            activeCell.style.fontWeight = '600';
        }
        else{
            activeCell.style.fontWeight = '400';
        }
        activeSheetObject[activeCell.id].isBold = !activeSheetObject[activeCell.id].isBold;
        activeCell.focus();
    }
}

// bug fix
italicBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    toggleItalic();
})
function toggleItalic(){
    if(activeCell){
        if(!activeSheetObject[activeCell.id].isItalic) {
            activeCell.style.fontStyle = 'italic';
        }
        else{
            activeCell.style.fontStyle = 'normal';
        }
        activeSheetObject[activeCell.id].isItalic = !activeSheetObject[activeCell.id].isItalic;
        activeCell.focus();
    }
}

// bug fix
underlineBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    toggleUnderline();
})
function toggleUnderline(){
    if(activeCell){
        if(!activeSheetObject[activeCell.id].isUnderlined) {
            activeCell.style.textDecoration = 'underline';
        }
        else{
            activeCell.style.textDecoration = 'none';
        }
        activeSheetObject[activeCell.id].isUnderlined = !activeSheetObject[activeCell.id].isUnderlined;
        activeCell.focus();
    }
}


// bug ix
document.querySelectorAll('.color-prop').forEach(e => {
    e.addEventListener('click', (event) => event.stopPropagation());
})
function textColor(target){
    if(activeCell){
        let colorInput = target.value;
        activeSheetObject[activeCell.id].color = colorInput;
        activeCell.style.color = colorInput;
        activeCell.focus();
    }
}
function backgroundColor(target){
    if(activeCell){
        let colorInput = target.value;
        activeSheetObject[activeCell.id].backgroundColor = colorInput;
        activeCell.style.backgroundColor = colorInput;
        activeCell.focus();
    }
}

// bug fix
document.querySelectorAll('.alignment').forEach(e => {
    e.addEventListener('click', (event) => {
        event.stopPropagation();
        let align = e.className.split(' ')[0];
        alignment(align);
    });
})
function alignment(align){
    if(activeCell){
        activeCell.style.textAlign = align;
        activeSheetObject[activeCell.id].textAlign = align;
        activeCell.focus();
    }
}



document.querySelector('.copy').addEventListener('click', (event) => {
    event.stopPropagation();
    if (activeCell) {
        navigator.clipboard.writeText(activeCell.innerText);
        activeCell.focus();
    }
})

document.querySelector('.cut').addEventListener('click', (event) => {
    event.stopPropagation();
    if (activeCell) {
        navigator.clipboard.writeText(activeCell.innerText);
        activeCell.innerText = '';
        activeCell.focus();
    }
})

document.querySelector('.paste').addEventListener('click', (event) => {
    event.stopPropagation();
    if (activeCell) {
        navigator.clipboard.readText().then((text) => {
            formula.value = text;
            activeCell.innerText = text;
        })
        activeCell.focus();
    }
})

downloadBtn.addEventListener("click",(e)=>{
    let jsonData = JSON.stringify(sheetsArray);
    let file = new Blob([jsonData],{type: "application/json"});
    let a = document.createElement("a");
    a.href = URL.createObjectURL(file);
    a.download = "SheetData.json";
    a.click();
});

openBtn.addEventListener("click",(e)=>{
    let input = document.createElement("input");
    input.setAttribute("type","file");
    input.click();

    input.addEventListener("change",(e)=>{
        let fr = new FileReader();
        let files = input.files;
        let fileObj = files[0];

        fr.readAsText(fileObj);
        fr.addEventListener("load",(e)=>{
            let readSheetData = JSON.parse(fr.result);
            readSheetData.forEach(e => {
                document.querySelector('.new-sheet').click();
                sheetsArray[activeSheetIndex] = e;
                activeSheetObject = e;
                changeSheet();
            })
        });
    });
});

formula.addEventListener('input', (event) => {
    activeCell.innerText = event.target.value;
    activeSheetObject[activeCell.id].content = event.target.value;
})



// try for bug fix
document.querySelector('body').addEventListener('click', () => {
    resetFunctionality();
})

// bug fix
formula.addEventListener('click', (event) => event.stopPropagation());

document.querySelectorAll('.select,.color-prop>*').forEach(e => {
    e.addEventListener('click', event => {
        event.stopPropagation();
    });
})

// Add this code to functionalities.js after all your existing functions

// Formula evaluation when cell content changes
function evaluateCellFormulas() {
    // Update all cells that contain formulas
    for (let cellId in activeSheetObject) {
        const cellData = activeSheetObject[cellId];
        if (cellData.content && typeof cellData.content === 'string' && cellData.content.startsWith('=')) {
            const result = evaluateFormula(cellData.content);
            // Store both formula and result
            cellData.formula = cellData.content;
            cellData.content = result;
            
            // Update the display
            const cell = document.getElementById(cellId);
            if (cell) {
                cell.innerText = result;
            }
        }
    }
}

// Add math operation buttons to the UI
function addMathOperationButtons() {
    const cellStyles = document.querySelector('.cell-styles');
    
    // Create math operations container
    const mathOperations = document.createElement('div');
    mathOperations.className = 'math-operations';
    mathOperations.style.display = 'flex';
    mathOperations.style.gap = '5px';
    mathOperations.style.marginLeft = '10px';
    
    // Add button
    const addBtn = document.createElement('button');
    addBtn.innerHTML = '+';
    addBtn.title = 'Add numbers in cell';
    addBtn.className = 'math-btn';
    addBtn.style.padding = '2px 8px';
    addBtn.style.cursor = 'pointer';
    addBtn.onclick = function() { performMathOperation('add'); };
    
    // Multiply button
    const multiplyBtn = document.createElement('button');
    multiplyBtn.innerHTML = '×';
    multiplyBtn.title = 'Multiply numbers in cell';
    multiplyBtn.className = 'math-btn';
    multiplyBtn.style.padding = '2px 8px';
    multiplyBtn.style.cursor = 'pointer';
    multiplyBtn.onclick = function() { performMathOperation('multiply'); };
    
    // Subtract button
    const subtractBtn = document.createElement('button');
    subtractBtn.innerHTML = '-';
    subtractBtn.title = 'Subtract numbers in cell';
    subtractBtn.className = 'math-btn';
    subtractBtn.style.padding = '2px 8px';
    subtractBtn.style.cursor = 'pointer';
    subtractBtn.onclick = function() { performMathOperation('subtract'); };
    
    // Divide button
    const divideBtn = document.createElement('button');
    divideBtn.innerHTML = '÷';
    divideBtn.title = 'Divide numbers in cell';
    divideBtn.className = 'math-btn';
    divideBtn.style.padding = '2px 8px';
    divideBtn.style.cursor = 'pointer';
    divideBtn.onclick = function() { performMathOperation('divide'); };
    
    // Formula Helper button
    const formulaBtn = document.createElement('button');
    formulaBtn.innerHTML = 'f(x)';
    formulaBtn.title = 'Insert formula';
    formulaBtn.className = 'math-btn';
    formulaBtn.style.padding = '2px 8px';
    formulaBtn.style.cursor = 'pointer';
    formulaBtn.onclick = function() { showFormulaHelper(); };
    
    // Add buttons to the container
    mathOperations.appendChild(addBtn);
    mathOperations.appendChild(multiplyBtn);
    mathOperations.appendChild(subtractBtn);
    mathOperations.appendChild(divideBtn);
    mathOperations.appendChild(formulaBtn);
    
    // Add the container to cell styles
    cellStyles.appendChild(mathOperations);
}


// Perform math operation on active cell with validation
// Perform math operation on active cell with multiple comma-separated values
function performMathOperation(operation) {
    if (!activeCell) {
        alert('Please select a cell first');
        return;
    }
    
    const cellId = activeCell.id;
    const cellContent = activeSheetObject[cellId].content;
    
    // Check if cell is empty
    if (!cellContent || cellContent.trim() === "") {
        activeSheetObject[cellId].content = "NaN";
        activeCell.innerText = "NaN";
        formula.value = "NaN";
        return;
    }
    
    try {
        // Parse comma-separated values from the cell content
        const values = cellContent.split(',').map(item => item.trim());
        
        // Check if all values are valid numbers
        const numberValues = [];
        let hasInvalidValue = false;
        
        for (const value of values) {
            if (value === "" || isNaN(parseFloat(value))) {
                hasInvalidValue = true;
                break;
            }
            numberValues.push(parseFloat(value));
        }
        
        // If any value is not a valid number, show NaN
        if (hasInvalidValue || numberValues.length === 0) {
            activeSheetObject[cellId].content = "NaN";
            activeCell.innerText = "NaN";
            formula.value = "NaN";
            return;
        }
        
        // Perform the operation on all numbers
        let result;
        switch(operation) {
            case 'add':
                result = numberValues.reduce((sum, value) => sum + value, 0);
                break;
            case 'subtract':
                // Start with first value and subtract all others
                result = numberValues.slice(1).reduce((diff, value) => diff - value, numberValues[0]);
                break;
            case 'multiply':
                result = numberValues.reduce((product, value) => product * value, 1);
                break;
            case 'divide':
                // Check for division by zero in any value after the first
                if (numberValues.slice(1).some(val => val === 0)) {
                    alert("Cannot divide by zero");
                    activeSheetObject[cellId].content = "NaN";
                    activeCell.innerText = "NaN";
                    formula.value = "NaN";
                    return;
                }
                // Start with first value and divide by all others
                result = numberValues.slice(1).reduce((quotient, value) => quotient / value, numberValues[0]);
                break;
            default:
                result = "NaN";
        }
        
        // Format the result to prevent excessively long decimals
        result = typeof result === 'number' ? 
                 (Number.isInteger(result) ? result : parseFloat(result.toFixed(4))) : 
                 result;
        
        // Update cell content with valid result
        activeSheetObject[cellId].content = result;
        activeCell.innerText = result;
        formula.value = result;
        
    } catch (error) {
        console.error("Math operation error:", error);
        // Show NaN in cell for any error
        activeSheetObject[cellId].content = "NaN";
        activeCell.innerText = "NaN";
        formula.value = "NaN";
    }
}
// Show formula helper dropdown
function showFormulaHelper() {
    if (!activeCell) {
        alert('Please select a cell first');
        return;
    }
    
    const formulaTypes = [
        { name: 'SUM', syntax: '=SUM(A1:B4)' },
        { name: 'AVERAGE', syntax: '=AVERAGE(A1:B4)' },
        { name: 'MAX', syntax: '=MAX(A1:B4)' },
        { name: 'MIN', syntax: '=MIN(A1:B4)' },
        { name: 'COUNT', syntax: '=COUNT(A1:B4)' }
    ];
    
    // Create dropdown for formulas
    const dropdown = document.createElement('div');
    dropdown.className = 'formula-dropdown';
    dropdown.style.position = 'absolute';
    dropdown.style.backgroundColor = 'white';
    dropdown.style.border = '1px solid #ccc';
    dropdown.style.borderRadius = '4px';
    dropdown.style.padding = '5px 0';
    dropdown.style.zIndex = '100';
    dropdown.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
    
    // Calculate position (under the formula button)
    const formulaBtn = document.querySelector('.math-operations').lastChild;
    const rect = formulaBtn.getBoundingClientRect();
    dropdown.style.top = (rect.bottom + 5) + 'px';
    dropdown.style.left = rect.left + 'px';
    
    // Add formula options
    formulaTypes.forEach(formula => {
        const option = document.createElement('div');
        option.className = 'formula-option';
        option.innerHTML = `<b>${formula.name}</b> - ${formula.syntax}`;
        option.style.padding = '8px 15px';
        option.style.cursor = 'pointer';
        option.style.hover = 'background: #f0f0f0';
        
        option.addEventListener('mouseover', () => {
            option.style.backgroundColor = '#f0f0f0';
        });
        
        option.addEventListener('mouseout', () => {
            option.style.backgroundColor = 'transparent';
        });
        
        option.addEventListener('click', () => {
            // Insert formula template
            activeSheetObject[activeCell.id].content = formula.syntax;
            formula.value = formula.syntax;
            activeCell.innerText = formula.syntax;
            
            // Remove dropdown
            document.body.removeChild(dropdown);
        });
        
        dropdown.appendChild(option);
    });
    
    // Add to body
    document.body.appendChild(dropdown);
    
    // Close dropdown when clicking outside
    document.addEventListener('click', function closeDropdown(e) {
        if (!dropdown.contains(e.target) && e.target !== formulaBtn) {
            if (document.body.contains(dropdown)) {
                document.body.removeChild(dropdown);
            }
            document.removeEventListener('click', closeDropdown);
        }
    });
}

// Update existing cellInput function to handle formulas
const originalCellInput = window.cellInput || function(){};

window.cellInput = function(event) {
    // Call the original function
    originalCellInput.call(this, event);
    
    // Check if content is a formula
    const cellId = event.target.id;
    const cellContent = activeSheetObject[cellId].content;
    
    if (typeof cellContent === 'string' && cellContent.startsWith('=')) {
        // Evaluate the formula
        const result = evaluateFormula(cellContent);
        
        // Store formula and display result
        activeSheetObject[cellId].formula = cellContent;
        activeSheetObject[cellId].content = result;
        
        // Update display
        event.target.innerText = result;
        formula.value = cellContent; // Show formula in formula bar
    }
};

// Update existing changeSheet function to handle formulas
const originalChangeSheet = window.changeSheet || function(){};

window.changeSheet = function() {
    // Call the original function
    originalChangeSheet.call(this);
    
    // Re-evaluate formulas after sheet change
    evaluateCellFormulas();
};

// Modify formula bar event listener to handle formulas
document.querySelector('.formula-bar').addEventListener('input', function(e) {
    if (!activeCell) return;
    
    activeSheetObject[activeCell.id].content = e.target.value;
    
    // If it's a formula, evaluate it
    if (e.target.value.startsWith('=')) {
        const result = evaluateFormula(e.target.value);
        
        // Store formula and display result
        activeSheetObject[activeCell.id].formula = e.target.value;
        activeSheetObject[activeCell.id].content = result;
        
        // Update display
        activeCell.innerText = result;
    } else {
        // Not a formula, just update the cell
        activeCell.innerText = e.target.value;
    }
});

// Initialize math buttons when page loads
document.addEventListener('DOMContentLoaded', function() {
    addMathOperationButtons();
});
