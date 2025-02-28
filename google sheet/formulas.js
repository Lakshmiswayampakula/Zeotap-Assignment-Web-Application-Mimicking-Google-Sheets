// Add this code to a new file called 'formulas.js'

// Parse and evaluate formulas in cells
function evaluateFormula(formula) {
    if (!formula || !formula.startsWith('=')) return formula;
    
    try {
        // Remove the equals sign and trim whitespace
        const expression = formula.substring(1).trim();
        
        // Handle different formula types
        if (expression.includes('SUM')) {
            return calculateSum(expression);
        } else if (expression.includes('AVERAGE')) {
            return calculateAverage(expression);
        } else if (expression.includes('MAX')) {
            return calculateMax(expression);
        } else if (expression.includes('MIN')) {
            return calculateMin(expression);
        } else if (expression.includes('COUNT')) {
            return calculateCount(expression);
        } else {
            // For basic arithmetic expressions
            return evaluateArithmeticExpression(expression);
        }
    } catch (error) {
        console.error("Formula error:", error);
        return "ERROR";
    }
}

// Extract cell range from formula (e.g., "A1:B4")
function extractCellRange(expression) {
    const rangeRegex = /([A-Z]+[0-9]+):([A-Z]+[0-9]+)/;
    const match = expression.match(rangeRegex);
    
    if (!match) {
        throw new Error("Invalid cell range format");
    }
    
    const startCell = match[1];
    const endCell = match[2];
    
    return getCellsInRange(startCell, endCell);
}

// Get all cells between start and end cells
function getCellsInRange(startCell, endCell) {
    const startCol = startCell.charAt(0).charCodeAt(0);
    const startRow = parseInt(startCell.substring(1));
    
    const endCol = endCell.charAt(0).charCodeAt(0);
    const endRow = parseInt(endCell.substring(1));
    
    const cells = [];
    
    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
            const cellId = String.fromCharCode(col) + row;
            cells.push(cellId);
        }
    }
    
    return cells;
}

// Get numerical values from a range of cells
function getValuesFromRange(cellsRange) {
    const values = [];
    
    for (const cellId of cellsRange) {
        if (!activeSheetObject[cellId] || !activeSheetObject[cellId].content) continue;
        
        const cellContent = activeSheetObject[cellId].content;
        const numValue = parseFloat(cellContent);
        
        if (!isNaN(numValue)) {
            values.push(numValue);
        }
    }
    
    return values;
}

// Calculate SUM function
function calculateSum(expression) {
    const cellsRange = extractCellRange(expression);
    const values = getValuesFromRange(cellsRange);
    
    if (values.length === 0) return 0;
    
    return values.reduce((sum, value) => sum + value, 0);
}

// Calculate AVERAGE function
function calculateAverage(expression) {
    const cellsRange = extractCellRange(expression);
    const values = getValuesFromRange(cellsRange);
    
    if (values.length === 0) return 0;
    
    const sum = values.reduce((total, value) => total + value, 0);
    return sum / values.length;
}

// Calculate MAX function
function calculateMax(expression) {
    const cellsRange = extractCellRange(expression);
    const values = getValuesFromRange(cellsRange);
    
    if (values.length === 0) return 0;
    
    return Math.max(...values);
}

// Calculate MIN function
function calculateMin(expression) {
    const cellsRange = extractCellRange(expression);
    const values = getValuesFromRange(cellsRange);
    
    if (values.length === 0) return 0;
    
    return Math.min(...values);
}

// Calculate COUNT function
function calculateCount(expression) {
    const cellsRange = extractCellRange(expression);
    const values = getValuesFromRange(cellsRange);
    
    return values.length;
}

// Evaluate arithmetic expressions (supports basic operations)
function evaluateArithmeticExpression(expression) {
    // Check for cell references like A1, B2, etc.
    const cellRegex = /[A-Z]+[0-9]+/g;
    const cellReferences = expression.match(cellRegex) || [];
    
    let processedExpression = expression;
    
    // Replace cell references with their values
    for (const cellRef of cellReferences) {
        if (activeSheetObject[cellRef] && activeSheetObject[cellRef].content) {
            const cellValue = activeSheetObject[cellRef].content;
            const numValue = parseFloat(cellValue);
            
            if (isNaN(numValue)) {
                return "NaN"; // If any cell contains non-numeric data
            }
            
            processedExpression = processedExpression.replace(cellRef, numValue);
        } else {
            processedExpression = processedExpression.replace(cellRef, 0);
        }
    }
    
    // Safely evaluate the arithmetic expression
    try {
        // Validate the expression contains only safe characters
        if (/^[0-9+\-*/() .]+$/.test(processedExpression)) {
            // Use Function constructor to evaluate the expression
            return new Function(`return ${processedExpression}`)();
        } else {
            return "ERROR: Invalid characters";
        }
    } catch (error) {
        console.error("Arithmetic evaluation error:", error);
        return "ERROR";
    }
}


// Update the processInlineMath function to handle comma-separated values
function processInlineMath(cellContent, operation) {
    if (!cellContent) return "0";
    
    try {
        // Split by commas and extract numeric values
        const values = cellContent.split(',').map(item => item.trim());
        const numValues = [];
        
        // Validate all values are numbers
        for (const value of values) {
            if (value === "" || isNaN(parseFloat(value))) {
                return "NaN";
            }
            numValues.push(parseFloat(value));
        }
        
        if (numValues.length === 0) {
            return "NaN";
        }
        
        switch (operation) {
            case 'add':
                return numValues.reduce((sum, val) => sum + val, 0);
            case 'multiply':
                return numValues.reduce((product, val) => product * val, 1);
            case 'subtract':
                // Subtract all values from the first one
                return numValues.slice(1).reduce((result, val) => result - val, numValues[0]);
            case 'divide':
                // Check for division by zero in any value after the first
                if (numValues.slice(1).some(val => val === 0)) {
                    throw new Error("Division by zero");
                }
                // Divide first value by all others
                return numValues.slice(1).reduce((result, val) => result / val, numValues[0]);
            default:
                return "ERROR: Invalid operation";
        }
    } catch (error) {
        console.error("Math processing error:", error);
        return "NaN";
    }
}