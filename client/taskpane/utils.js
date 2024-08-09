/* global document, Office, Word */

/**
 *
 */
export const TABLE_HEADERS = {
  COLUMN_HEADERS: ["PHASE", "FS DESIGN FEE", "LAUNDRY DESIGN FEE", "TOTAL"],
  ROW_HEADERS: [
    "Schematic Design",
    "Design Development",
    "Contract Documents",
    "Bidding and Contract Award",
    "Services During Construction",
  ],
};

let test = false;

//ignore first 'n' non-numeric columns - this will be used as lower bound in for loop to create input cells
const N = 1;

const COLS = TABLE_HEADERS.COLUMN_HEADERS.length;
const ROWS = TABLE_HEADERS.ROW_HEADERS.length;

let dialog;
let list;
const numberRegex = /^\d+(\.\d+)?$/; // Regular expression to match numeric strings

/**
 *
 */
export class NumberCell {
  /**
   *
   * @param {*} parentElement
   */
  constructor(parentElement) {
    this.cell = document.createElement("td");
    this.numberInput = document.createElement("input");
    this.numberInput.type = "number";
    this.numberInput.classList.add("num-cell");
    this.cell.classList.add("num-cell");
    this.cell.appendChild(this.numberInput);
    parentElement.appendChild(this.cell);

    // Add event listener for change event to recalculate totals
    this.numberInput.addEventListener("input", () => {
      calculateRowTotal(parentElement);
      calculateColumnTotal(this.cell.cellIndex);
      calculateGrandTotal();
    });
  }
  /**
   *
   * @returns
   */
  getValue() {
    return parseFloat(this.numberInput.value) || 0; // Return 0 if value is empty
  }

  /**
   *
   * @param {*} value
   */
  setValue(value) {
    this.numberInput.value = value;
  }
}

/**
 *
 */
export class TotalCell {
  /**
   *
   * @param {*} parentElement
   */
  constructor(parentElement) {
    this.cell = document.createElement("td");
    this.value = 0;
    this.cell.classList.add("total-cell");
    parentElement.appendChild(this.cell);
  }
  /**
   *
   * @param {*} value
   */
  setValue(value) {
    this.value = value;
    this.cell.textContent = value.toFixed(2);
  }

  /**
   *
   * @returns
   */
  getValue() {
    return this.value;
  }
}

/**
 *
 * @param {*} obj
 * @param {*} searchValue
 * @param {*} replaceValue
 * @returns
 */
export function replaceAllInObject(obj, searchValue, replaceValue) {
  if (typeof obj === "string") {
    return obj.replace(new RegExp(searchValue, "g"), replaceValue);
  }
  if (Array.isArray(obj)) {
    return obj.map((item) => replaceAllInObject(item, searchValue, replaceValue));
  }
  if (typeof obj === "object" && obj !== null) {
    const newObj = {};
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        newObj[key] = replaceAllInObject(obj[key], searchValue, replaceValue);
      }
    }
    return newObj;
  }
  return obj;
}

/**
 *
 * @param {*} id
 * @returns
 */
export function createTableValuesFromHtmlTable(id) {
  const addDollarSigns = (tableData) => {
    const numberRegex = /^\d+(\.\d+)?$/; // Regular expression to match numeric strings
    tableData.forEach((row) => {
      row.forEach((cell) => {
        if (numberRegex.test(cell.value)) {
          cell.value = `$${cell.value}`;
        }
      });
    });
  };

  const htmlTable = document.getElementById(id);
  let rows = Array.from(htmlTable.querySelectorAll("tr"));

  const tableData = Array.from(rows).map((row, rowIndex) => {
    return Array.from(row.cells).map((cell) => {
      const input = cell.querySelector("input");
      let value = input ? input.value : cell.textContent.trim();

      return {
        value: value,
        bold: rowIndex === 0 || ["optional services", "total design fee", "grand total"].includes(value.toLowerCase()),
        italic: ["optional services", "total design fee", "grand total", "hourly as needed"].includes(
          value.toLowerCase()
        ),
      };
    });
  });

  // Assuming the "totals" row is the last row
  const totalsRow = tableData[tableData.length - 1];

  const columnsToRemove = [];
  totalsRow.forEach((cell, columnIndex) => {
    if (columnIndex === 0) return; // Always include the first column
    const value = parseFloat(cell.value.replace(/[^0-9.-]+/g, ""));
    if (isNaN(value) || value === 0) {
      columnsToRemove.push(columnIndex);
    }
  });

  // Check if we need to remove the rightmost total column
  let removeRightmostTotalColumn = false;
  for (let i = 1; i < totalsRow.length - 1; i++) {
    const value = parseFloat(totalsRow[i].value.replace(/[^0-9.-]+/g, ""));
    if (isNaN(value) || value === 0) {
      removeRightmostTotalColumn = true;
      break;
    }
  }

  if (removeRightmostTotalColumn) {
    columnsToRemove.push(totalsRow.length - 1);
  }

  const cleanedTableData = tableData.map((row) => row.filter((_, index) => !columnsToRemove.includes(index)));

  addDollarSigns(cleanedTableData);

  const tableValues = cleanedTableData.map((row) =>
    row.map((cell) => ({
      value: cell.value,
      bold: cell.bold,
      italic: cell.italic,
    }))
  );

  const totalColumns = tableValues[0].length;
  if (totalColumns === 2) {
    // Ensure table data structure is correct
    tableValues.push([
      { value: "", bold: false, italic: false },
      { value: "", bold: false, italic: false },
    ]);
    tableValues.push([
      { value: "Optional Services", bold: true, italic: false },
      { value: "", bold: false, italic: false },
    ]);
    tableValues.push([
      { value: "Bidding and Contract Award", bold: true, italic: true },
      { value: "Hourly as Needed", bold: false, italic: true },
    ]);
  } else if (totalColumns === 4) {
    // Ensure table data structure is correct
    tableValues.push([
      { value: "", bold: false, italic: false },
      { value: "", bold: false, italic: false },
      { value: "", bold: false, italic: false },
      { value: "", bold: false, italic: false },
    ]);
    tableValues.push([
      { value: "Optional Services", bold: true, italic: false },
      { value: "", bold: false, italic: false },
      { value: "", bold: false, italic: false },
      { value: "", bold: false, italic: false },
    ]);
    tableValues.push([
      { value: "Bidding and Contract Award", bold: true, italic: true },
      { value: "Hourly as Needed", bold: false, italic: true },
      { value: "Hourly as Needed", bold: false, italic: true },
      { value: "", bold: false, italic: false },
    ]);
  }

  // Determine the column titles
  const columnTitles = tableValues[0].map((cell) => cell.value);

  return { tableValues, columnTitles };
}

/**
 *
 * @param {*} TABLE_HEADERS
 * @param {*} document
 * @returns
 */
export function generateCostTable(TABLE_HEADERS, document) {
  const includeExistingCheckbox = document.getElementById("include_existing");
  const cost_table = document.getElementById("cost-table");

  includeExistingCheckbox.addEventListener("change", () => {
    const tableBody = cost_table.querySelector("tbody");
    const existingEquipRow = tableBody.querySelector("#existing_equipment");

    if (includeExistingCheckbox.checked && !existingEquipRow) {
      // Create and append the existing row
      const existingEquipRow = document.createElement("tr");
      existingEquipRow.id = "existing_equipment";

      const existingHeader = document.createElement("td");
      existingHeader.textContent = "Existing Equipment";
      existingEquipRow.appendChild(existingHeader);

      for (let j = 1; j < TABLE_HEADERS.COLUMN_HEADERS.length - 1; j++) {
        // Adjusted loop to match column length
        const existingCell = document.createElement("td");
        existingCell.classList.add("num-cell");
        existingCell.id = sanitizeId("existing_" + TABLE_HEADERS.COLUMN_HEADERS[j]);
        existingCell.innerHTML = '<input type="number" class="num-cell">';
        existingCell.querySelector("input").addEventListener("input", () => {
          calculateRowTotal(existingEquipRow);
          calculateColumnTotal(j);
          calculateGrandTotal();
        });
        existingEquipRow.appendChild(existingCell);
      }

      const existingTotal = document.createElement("td");
      existingTotal.classList.add("total-cell", "row-total");
      existingTotal.textContent = "0.00";
      existingEquipRow.appendChild(existingTotal);

      tableBody.insertBefore(existingEquipRow, tableBody.firstChild);

      // Recalculate totals after adding the row
      calculateRowTotal(existingEquipRow);
      for (let j = 1; j < TABLE_HEADERS.COLUMN_HEADERS.length - 1; j++) {
        calculateColumnTotal(j);
      }
      calculateGrandTotal();
    } else if (!includeExistingCheckbox.checked && existingEquipRow) {
      // Remove the existing row
      tableBody.removeChild(existingEquipRow);

      // Recalculate totals after removing the row
      for (let i = 0; i < tableBody.rows.length - 1; i++) {
        calculateRowTotal(tableBody.rows[i]);
      }
      for (let j = 1; j < TABLE_HEADERS.COLUMN_HEADERS.length - 1; j++) {
        calculateColumnTotal(j);
      }
      calculateGrandTotal();
    }
  });

  /**
   *
   * @param {*} str
   * @returns
   */
  function sanitizeId(str) {
    return str.toLowerCase().replace(/\W/g, "_"); // Convert to lowercase and replace non-alphanumeric with underscores
  }

  // Create header row
  const tableHead = document.createElement("thead");
  const tableBody = document.createElement("tbody");

  const row_headers = TABLE_HEADERS.ROW_HEADERS;
  const col_headers = TABLE_HEADERS.COLUMN_HEADERS;

  const headerRow = document.createElement("tr");

  for (let i = 0; i < col_headers.length; i++) {
    const cell = document.createElement("th");
    cell.textContent = col_headers[i];
    cell.id = sanitizeId(cell.textContent);

    if (i === 1 || i === 2) {
      cell.classList.add("num-column");
    }

    headerRow.appendChild(cell);
  }

  console.log(headerRow);
  tableHead.appendChild(headerRow);
  cost_table.appendChild(tableHead);

  for (let i = 0; i < ROWS; i++) {
    const tableRow = document.createElement("tr");

    const headerCell = document.createElement("td");
    const header = row_headers[i];
    headerCell.textContent = header;
    headerCell.id = sanitizeId(header);
    tableRow.appendChild(headerCell);

    for (let j = N; j < COLS - 1; j++) {
      const numberCell = new NumberCell(tableRow);
      numberCell.cell.id = sanitizeId(header + col_headers[j]);
      tableRow.appendChild(numberCell.cell);
    }

    const totalCell = new TotalCell(tableRow);
    totalCell.cell.classList.add("row-total");
    totalCell.cell.textContent = "0.00";
    console.log(tableRow);
    tableBody.appendChild(tableRow);
  }

  // Add totals row and column
  const totalsRow = document.createElement("tr");

  const totalRowHeader = document.createElement("th");
  totalRowHeader.textContent = "Total Design Fee";
  totalRowHeader.id = "vertical_sums_label";
  totalsRow.appendChild(totalRowHeader);

  for (let j = 1; j < TABLE_HEADERS.COLUMN_HEADERS.length; j++) {
    const totalCell = new TotalCell(totalsRow);
    totalCell.cell.classList.add("num-cell", "column-total");
    totalCell.cell.id = sanitizeId(TABLE_HEADERS.COLUMN_HEADERS[j] + "total");
    totalCell.cell.textContent = "0.00";
  }

  tableBody.appendChild(totalsRow);
  cost_table.appendChild(tableBody);

  return cost_table;
}

/**
 *
 * @param {*} row
 */
export function calculateRowTotal(row) {
  const cells = row.querySelectorAll(".num-cell");
  let total = 0;

  cells.forEach((cell) => {
    const input = cell.querySelector("input");
    if (input) {
      total += parseFloat(input.value) || 0;
    }
  });

  const totalCell = row.querySelector(".row-total");
  if (totalCell) {
    totalCell.textContent = total.toFixed(2);
  }
}

/**
 *
 * @param {\} colIndex
 */
export function calculateColumnTotal(colIndex) {
  const a_table = document.getElementById("cost-table");
  const rows = a_table.querySelectorAll("tbody tr");
  let total = 0;

  rows.forEach((row) => {
    const cell = row.cells[colIndex];
    if (cell) {
      const input = cell.querySelector("input");
      if (input) {
        total += parseFloat(input.value) || 0;
      } else {
        total += !cell.classList.contains("num-cell") ? parseFloat(cell.textContent) || 0 : 0;
      }
    }
  });

  const totalRow = a_table.querySelector("tbody tr:last-child");
  const totalCell = totalRow.cells[colIndex];
  if (totalCell) {
    totalCell.textContent = total.toFixed(2);
  }
}

/**
 *
 */
export function calculateGrandTotal() {
  const b_table = document.getElementById("cost-table");
  const rows = b_table.querySelectorAll("tbody tr:not(:last-child)"); // Exclude the last row (totals row)
  let grandTotal = 0;

  rows.forEach((row) => {
    const totalCell = row.querySelector(".row-total");
    if (totalCell) {
      grandTotal += parseFloat(totalCell.textContent) || 0;
    }
  });

  const bottomRightCell = b_table.querySelector("tbody tr:last-child .column-total:last-child");
  if (bottomRightCell) {
    bottomRightCell.textContent = grandTotal.toFixed(2);
  }
}

/**
 *
 * @returns
 */
export function validateForm() {
  //console.log("validating");
  const sig = document.getElementById("sig").value.trim();
  const ref = document.getElementById("ref").value.trim();
  const loc = document.getElementById("loc").value.trim();
  if (!sig || !ref || !loc) {
    return false;
  }
  return true;
}

/**
 *
 * @param {*} msg
 */
export function showError(msg) {
  Office.context.ui.displayDialogAsync(
    "https://dev--proposaltemplate.netlify.app/incompleteForm.html",
    { height: 30, width: 20 },
    function (result) {
      dialog = result.value;
      // console.log('got here')
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      // console.log('got here')
    }
  );
}

/**
 *
 * @param {*} arg
 */
export function processMessage(arg) {
  // console.log(arg)
  if (arg.message == "close") {
    // console.log('messaged')
    dialog.close();
  }
}

// Helper function to fetch the image and convert it to a base64 string
/**
 *
 * @param {*} imageUrl
 * @returns
 */
export async function getImageBase64(imageUrl) {
  const response = await fetch(imageUrl);
  const blob = await response.blob();

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
