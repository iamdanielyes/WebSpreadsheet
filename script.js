const spreadSheetContainer = document.querySelector("#spreadsheet-container")
//one row and one column for the headers
const ROWS = 101
const COLS = 101
let spreadsheet = []
let cell = null

class Cell {
    constructor(isHeader, disabled, data, formula, row, column, rowName, columnName, active = false, fontStyle, fontWeight) {
        this.isHeader = isHeader
        this.disabled = disabled
        this.data = data
        this.formula = formula
        this.row = row
        this.rowName = rowName
        this.column = column
        this.columnName = columnName
        this.active = active
        this.fontStyle = fontStyle
        this.fontWeight = fontWeight
    }
}

initSpreadsheet()
createRefreshButton()

function initSpreadsheet() {
    for (let i = 0; i < COLS; i++) {
        let spreadsheetRow = []
        for ( j = 0; j < COLS; j++) {
            let cellData = ""
            let isHeader = false
            let disabled = false
            if (j === 0) {
                cellData = i
                isHeader = true
                disabled = true
            }

            if (i === 0) {
                isHeader = true
                disabled = true
                if (j === 0){
                    cellData = 0
                } else {
                    cellData = numberToLetters(j-1)
                }   
            } else {
                //test
                cellData = numberToLetters(j-1) + i
            }

            if (!cellData) {
                cellData = ""
            }
            rowName = i
            if ( j === 0) {
                columnName = 0
            } else {
                columnName = numberToLetters(j-1)
            }
            cell = new Cell(isHeader, disabled, cellData, "", i, j, rowName, columnName, false, "", "")
            spreadsheetRow.push(cell)
        }
        spreadsheet.push(spreadsheetRow)
    }

    drawSheet()
}

function drawSheet() {
    spreadSheetContainer.innerHTML = ""
    for (let i = 0; i < spreadsheet.length; i++) {
        let rowContainerEl = document.createElement("div")
        rowContainerEl.className = "cell-row"

        for (let j = 0; j < spreadsheet[i].length; j++) {
            let cell = spreadsheet[i][j]
            rowContainerEl.append(createCellEl(cell, i, j))
        }
        spreadSheetContainer.append(rowContainerEl)
    }
}

function createCellEl(cell, i, j) {
    let cellEl = document.createElement("input")
    cellEl.className = "cell"
    if (cell.isHeader) {
        cellEl.classList.add("header")
        cellEl.id = "hdr_cell_" + cell.columnName + cell.row
    } else {
        cellEl.id = "cell_" + cell.columnName + cell.row
        cellEl.style.fontWeight = cell.fontStyle
    }
    if (cell.formula) {
        cellEl.value = evaluateFormula(cell.formula)
        spreadsheet[i][j].data = cellEl.value
    } else{
        cellEl.value = cell.data
    }
    cellEl.disabled = cell.disabled

    cellEl.onclick = () => handleCellClick(cell)
    cellEl.onchange = (e) => handleOnChange(e.target.value, cell)

    return cellEl
}

function handleCellClick(cell) {
    clearHeaderActiveStates()

    if (cell.formula) {
        document.getElementById("cell_" + cell.columnName + cell.row).value = cell.formula
    }
    let columnHeader = spreadsheet[0][cell.column]
    let rowHeader = spreadsheet[cell.row][0]
    let columnHeaderEl = getElFromRowCol(columnHeader.row, columnHeader.columnName)
    let rowHeaderEl = getElFromRowCol(rowHeader.row, rowHeader.column)
    columnHeaderEl.classList.add("active")
    rowHeaderEl.classList.add("active")
    document.querySelector("#cell-status").innerHTML = cell.columnName + "" + cell.rowName
}

function handleOnChange(data, cell) {
       //check for the "=" sign
    if (data.substring(0,1) === "=") {
        cell.formula = data
        resultFormula = evaluateFormula(cell.formula)
        document.getElementById("cell_" + cell.columnName + cell.row).value = resultFormula
        spreadsheet[cell.row][cell.column].data = resultFormula
    } else {
        cell.data = data
    }
   
}

function clearHeaderActiveStates() {
    const headers = document.querySelectorAll(".header")
      
    headers.forEach((header) => {
        header.classList.remove("active")
    })
}

function getElFromRowCol(row, col) {
    if (row === 0 || col === 0) {
        return document.querySelector("#hdr_cell_" + col + row)
    } else {
        return document.querySelector("#cell_" + col + row)
    }
}

function numberToLetters(num) {
    let letters = ''
    while (num >= 0) {
        letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[num % 26] + letters
        num = Math.floor(num / 26) - 1
    }
    return letters
}

//button for refreshing the spreadsheet
function createRefreshButton() {
    //create button element
    let button = document.createElement("button")
    button.id = "button"
    button.textContent = "Refresh"

    //append button to document body
    document.body.appendChild(button)

    button.addEventListener("click", function ()  {
        drawSheet()
    })
}

function evaluateFormula(formula) {
    //remove the "=" sign
    //if (formula.substring(0,1) === "=") {
        formula = formula.substring(1)
    //} else {
    //    return "ERROR: not a valid formula"
    //}

    //check if formula is an arithmetic expression
    if (/^[A-Z]+[0-9]+[+\-*/][A-Z]+[0-9]+$/.test(formula)) {
    // Split the formula by the operator
    let parts = formula.split(/([+\-*/])/)

    // Get the operands and the operator
    let operand1 = parts[0]
    let operator = parts[1]
    let operand2 = parts[2]

    // Get the values of the operands from the data object
    let value1 = spreadsheet[getRowIndex(operand1)][getColumnNumber(operand1)].data
    let value2 = spreadsheet[getRowIndex(operand2)][getColumnNumber(operand2)].data

    value1 = Number(value1)
    value2 = Number(value2)

    // Check if the values are valid numbers
    if (isNaN(value1) || isNaN(value2)) {
        return "ERROR: at least one parameter is not a number. "
    }

    // Perform the arithmetic operation and return the result
    switch (operator) {
        case "+":
            return value1 + value2
        case "-":
            return value1 - value2
        case "*":
            return value1 * value2
        case "/":
            return value1 / value2
        default:
            return "ERROR: not arithmetic"
    }
}

// Check if the formula is a sum function
if (/^sum\([A-Z]+[0-9]+:[A-Z]+[0-9]+\)$/.test(formula)) {
    // Remove the sum prefix and suffix
    formula = formula.substring(4, formula.length - 1)

    // Split the formula by the colon
    let parts = formula.split(":")

    // Get the start and end cell addresses
    let start = parts[0]
    let end = parts[1]

    // Get the column and row indexes of the start and end cells
    let startColumn = getColumnNumber(start)
    let startRow = getRowIndex(start)
    let endColumn = getColumnNumber(end)
    let endRow = getRowIndex(end)

    // Initialize a variable to store the sum
    let sum = 0

    //TO DO
    // Loop through all the cells in the range
    // Get the cell values
    // Check if the value is a valid number

    // Return the sum
    return sum
}

// If none of the above cases match, return an error message
return "ERROR"
}

function getColumnNumber(cell) {
    let column = cell.match(/[A-Z]+/)[0]
    let base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', i, j, result = 0

    for (i = 0, j = column.length - 1; i < column.length; i += 1, j -= 1) {
        result += Math.pow(base.length, j) * (base.indexOf(column[i]) + 1)
    }
     return result
}

//function to get the row index from a cell address (e.g. A1 -> 0)
function getRowIndex(cell) {
    // Get the second part of the cell address (e.g. A1 -> 1)
    let row = cell.match(/[0-9]+/)[0]

    // Convert the row part to a number
    let index = Number(row)
    return index
}