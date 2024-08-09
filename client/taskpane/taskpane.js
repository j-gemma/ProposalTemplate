/* global document, Office, Word */
let data = require("../../assets/data.js");
// const data = require("../../assets/lorem.js");

const utils = require("./utils.js");
const checkIds = ["sig", "ref", "loc"];
let test = false;

let table;
let generated = false;

/**
 * Initialize office.js, generate html form and cost table for user input
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("generate").onclick = validateFormAndRun;
    table = utils.generateCostTable(utils.TABLE_HEADERS, document);
  }
});

/**
 *
 */
export function validateFormAndRun() {
  //Uncomment to test
  // const test = false; // Set this to true to use test data

  let ready = false;

  //Uncomment for form validation
  ready = utils.validateForm(checkIds);
  //console.log(ready);

  if (ready) {
    let userData;

    if (test) {
      userData = {
        sig: ["FULL NAME, TITLE", "RECIPIENT POSITION", "FIRM NAME", "ADDRESS LINE 1", "CITY, ST ZIP"],
        ref: "A REFERENCE",
        loc: "LOCATION",
        existing: true,
        insurance: false,
        cost_table: {
          headers: TABLE_HEADERS,
          values: [[1, 2, 3]],
        },
      };
      //console.log("Using test data:", userData);
    } else {
      // Function to extract data from the form
      function getUserDataFromForm() {
        const sig = document.getElementById("sig").value.split("\n");
        const ref = document.getElementById("ref").value;
        const loc = document.getElementById("loc").value;
        const existing = document.getElementById("include_existing").checked;
        const insurance = document.getElementById("include_insurance").checked;

        // Extract table data
        const table = document.getElementById("cost-table");
        const headers = Array.from(table.querySelectorAll("thead th")).map((th) => th.innerText);
        const values = Array.from(table.querySelectorAll("tbody tr")).map((tr) => {
          return Array.from(tr.querySelectorAll("td")).map((td) => td.innerText);
        });

        return {
          sig,
          ref,
          loc,
          existing,
          insurance,
          cost_table: {
            headers,
            values,
          },
        };
      }

      userData = getUserDataFromForm();
      //console.log("User data extracted from form:", userData);
      run(userData);
    }
  } else {
    utils.showError("incomplete form");
  }

  // let ready = false;
  //let ready = true;
  // if (ready) {
  //   const checkbox1 = document.getElementById('include_existing');
  //   const checked1 = checkbox1.checked

  //   const checkbox2 = document.getElementById('include_insurance');
  //   const checked2 = checkbox2.checked

  //   let userData = {
  //     sig: [],
  //     ref: document.getElementById("ref").value,
  //     loc: document.getElementById("loc").value,
  //     existing: checked1,
  //     insurance: checked2,
  //     cost_table: {headers: TABLE_HEADERS,
  //       values : []
  //     }
  //   }

  //   run(userData);
  // }
}

/**
 *
 * @param {*} userData
 * @returns
 */
export async function run(userData) {
  if (test == true) {
    userData = {
      sig: ["FULL NAME, TITLE", "RECIPIENT POSITION", "FIRM NAME", "ADDRESS LINE 1", "CITY, ST ZIP"],
      ref: "A REFERENCE",
      loc: "LOCATION",
      existing: true,
      insurance: false,
      cost_table: { headers: utils.TABLE_HEADERS, values: [[1, 2, 3]] },
    };
  }

  /**
   *
   */
  return Word.run(async (context) => {
    // URL of the image in the assets folder
    const imageUrl = "../../assets/logoHiRes.jpg"; // Replace with the actual URL
    const img = await utils.getImageBase64(imageUrl);

    const sigImgUrl = "../../assets/sigImg.jpg";
    const sigImg = await utils.getImageBase64(sigImgUrl);

    let { tableValues, columnTitles } = utils.createTableValuesFromHtmlTable("cost-table");
    console.log(tableValues);

    // Clone the data object to avoid mutating the original
    let updatedData = JSON.parse(JSON.stringify(data));

    // Check if tableValues has 2 columns
    if (tableValues[0].length === 2) {
      const secondColumnTitle = columnTitles[1].toLowerCase();
      if (secondColumnTitle.includes("fs design fee")) {
        updatedData = utils.replaceAllInObject(updatedData, "foodservice and laundry", "foodservice");
      } else if (secondColumnTitle.includes("laundry design fee")) {
        updatedData = utils.replaceAllInObject(updatedData, "foodservice and laundry", "laundry");
        updatedData = utils.replaceAllInObject(updatedData, "foodservice", "laundry");
      }
    }

    data = updatedData;

    // Insert the image into the document at the current selection
    const doc = context.document;
    const body = doc.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("alignment, firstLineIndent, font, lineUnitBefore, lineUnitAfter,");
    await context.sync();

    let para = paragraphs.getFirstOrNullObject();

    para.alignment = "Left";

    const logo = para.insertInlinePictureFromBase64(img, Word.InsertLocation.start);
    logo.width = 50;
    logo.height = 50;

    body.font.name = "Arial Narrow";

    let today;
    let fsHeader;
    let fsIsPleased;
    let sectionHeaders;
    let documentStructure;

    if (data.today) {
      today = data.today;
    }

    if (data.sectionHeaders) {
      sectionHeaders = data.sectionHeaders;
    }

    if (data.fsHeader) {
      fsHeader = data.fsHeader;
    }

    if (data.fsIsPleased) {
      fsIsPleased = data.fsIsPleased;
    }

    if (data.sectionHeaders) {
      sectionHeaders = data.sectionHeaders;
    }

    if (data.documentStructure) {
      documentStructure = data.documentStructure;
    }

    const numMeetings = document.getElementById("num_meetings").value;
    const numVisits = document.getElementById("num_visits").value;

    for (const section of data.documentStructure.sections) {
      if (section.title === "Fees and Expenses") {
        if (numMeetings > 0 && numVisits > 0) {
          section.endText =
            `We have included up to ${numMeetings} meetings in the design phases and up to ${numVisits} site visits / punch lists during construction, ` +
            "as needed. Conference calls are included. In addition to the above fees, we will invoice separately at cost for typical " +
            "reimbursable expenses including travel costs. Other administrative expenses are included in the above fees. " +
            "Invoices are due and payable upon presentation monthly.";
        } else if (numMeetings == 0 && numVisits > 0) {
          section.endText =
            `We have included up to ${numVisits} site visits / punch lists during construction, ` +
            "as needed. Conference calls are included. In addition to the above fees, we will invoice separately at cost for typical " +
            "reimbursable expenses including travel costs. Other administrative expenses are included in the above fees. " +
            "Invoices are due and payable upon presentation monthly.";
        } else if (numMeetings > 0 && numVisits == 0) {
          section.endText =
            `We have included up to ${numMeetings} meetings in the design phases` +
            "Conference calls are included. In addition to the above fees, we will invoice separately at cost for typical " +
            "reimbursable expenses including travel costs. Other administrative expenses are included in the above fees. " +
            "Invoices are due and payable upon presentation monthly.";
        }
      }
    }

    para = paragraphs.getLastOrNullObject();
    for (const line of fsHeader) {
      para.insertParagraph(line, "After");
      para = paragraphs.getLastOrNullObject();
      para.alignment = "Right";
      para.font.name = "AvantGarde Bk Bt";
      para.lineSpacing = 12;
      para.font.size = 10;
      para.spaceAfter = 0;
      para.spaceBefore = 0;
      para.font.bold = false;
    }

    para.insertParagraph(today, "After");
    para = paragraphs.getLastOrNullObject();

    para.spaceAfter = 11;
    para.alignment = "left";
    para.font.name = "Arial Narrow";
    para.font.size = 11;
    para.font.bold = false;

    para = paragraphs.getLastOrNullObject();
    for (const line of userData.sig) {
      para.insertParagraph(line, "After");
      para = paragraphs.getLastOrNullObject();
      para.lineSpacing = 12;
      para.spaceAfter = 0;
      para.spaceBefore = 0;
      para.alignment = "Left";
    }

    const normalStyle = context.document.getStyles().getByNameOrNullObject("Normal");
    normalStyle.load();
    await context.sync();

    normalStyle.font.size = 11;
    normalStyle.font.bold = false;
    normalStyle.font.name = "Arial Narrow";
    normalStyle.font.color = "black";
    normalStyle.paragraphFormat.keepWithNext = false;

    const underlineStyle = context.document.getStyles().getByNameOrNullObject("No Spacing");
    underlineStyle.load();
    await context.sync();

    underlineStyle.font.size = 11;
    underlineStyle.font.bold = false;
    underlineStyle.font.underline = "Single";
    underlineStyle.font.name = "Arial Narrow";
    underlineStyle.font.color = "black";
    underlineStyle.paragraphFormat.keepWithNext = true;
    underlineStyle.paragraphFormat.spaceAfter = 0;

    para.insertParagraph("Reference: " + userData.ref, "After");
    para = paragraphs.getLastOrNullObject();
    para.spaceAfter = 11;
    const firstName = userData.sig[0].split(" ")[0];

    para.insertParagraph("Dear: " + firstName, "after");
    para = paragraphs.getLastOrNullObject();

    para.insertParagraph(fsIsPleased[0] + " ", "After");
    para = paragraphs.getLastOrNullObject();
    para.insertText(userData.ref + " in " + userData.loc + ". " + fsIsPleased[1], "End");

    para = paragraphs.getLastOrNullObject();
    para.insertParagraph("", "After");

    for (const section of documentStructure.sections) {
      if (
        (section.title === "Insurance" && !userData.insurance) ||
        (section.title === "Existing Equipment" && !userData.existing)
      ) {
        continue;
      }

      let title = section?.title;
      let smallCapsTitle = "<p style='font-variant: small-caps'>" + "<b>" + title + "<b>" + "</p>";

      para = paragraphs.getLastOrNullObject();
      // Insert the HTML with small caps
      para.insertHtml(smallCapsTitle ?? "Please provide title for section in data.js file", "End");

      para = paragraphs.getLastOrNullObject();

      para.leftIndent = 0;

      if (section.beginText) {
        if (section.title == "Planning Criteria") {
          para = paragraphs.getLastOrNullObject();
          para.insertParagraph("The " + userData.ref + " " + section.beginText, "After");
          para = paragraphs.getLastOrNullObject();
          para.style = normalStyle.nameLocal;
        } else {
          para.insertParagraph(section.beginText, "After");
          para = paragraphs.getLastOrNullObject();
          para.style = normalStyle.nameLocal;
        }
      }

      if (section.title == "Acceptance") {
        para.insertParagraph("Sincerely,", "After");
        para = paragraphs.getLastOrNullObject();
        para.spaceAfter = 0;
        para.insertParagraph("FoodStrategy, INC.", "After");
        para = paragraphs.getLastOrNullObject();
        para.insertParagraph("", "After");
        para = paragraphs.getLastOrNullObject();
        para.alignment = "Left";
        para.insertInlinePictureFromBase64(sigImg, Word.InsertLocation.end);
        para.insertParagraph("", "After");
      }

      if (section.title == "Fees and Expenses") {
        // Create a table with the same number of rows and columns as the HTML table
        const wordTable = para.insertTable(tableValues.length, tableValues[0].length, "After");

        tableValues.forEach((row, rowIndex) => {
          row.forEach((cell, cellIndex) => {
            const tableCell = wordTable.getCell(rowIndex, cellIndex);
            tableCell.body.insertText(cell.value, Word.InsertLocation.end);

            if (cell.bold) {
              //console.log('bold');
              tableCell.body.font.bold = true;
            }
            if (cell.italic) {
              //console.log('italic');
              tableCell.body.font.italic = true;
            }
          });
        });

        await context.sync();

        // para = paragraphs.getLastOrNullObject();
        // // Format the table (optional)
        // // wordTable.styleBuiltIn = Word.Style.gridTable5DarkAccent1;
        // wordTable.distributeColumns();

        // await context.sync();
        para = paragraphs.getLastOrNullObject();
        para.insertParagraph("", "After");
      }

      if (section.title == "Client Authorization") {
        para = paragraphs.getLastOrNullObject();
        para.spaceAfter = 11;
        para = paragraphs.getLastOrNullObject();
      }

      if (section.title != "Fees and Expenses") {
        para.insertParagraph("", "After");
      }
      para = paragraphs.getLastOrNullObject();

      if (section.list) {
        let items = section.list.items;
        let type = section.list.type;
        if (type != "labels") {
          let wordlist = para.startNewList();
          wordlist.load();
          await context.sync();

          switch (type) {
            case "number":
              wordlist.setLevelNumbering(0, Word.ListNumbering.arabic, [0, "."]);
              break;

            case "letter":
              wordlist.setLevelNumbering(0, Word.ListNumbering.upperLetter, [0, "."]);
              para = paragraphs.getLastOrNullObject();

              break;

            case "hourly":
              break;

            case "filloutForm":
              para.lineSpacing = 24; // Set double spacing for the filloutForm list type
              // wordlist.setLevelNumbering(0, Word.ListNumbering.none, [0, '.']);

              break;
          }

          para.insertText(items[0], "End");
          if (type == "letter") {
            para.insertHtml("<br></br>", Word.InsertLocation.end);
          }
          para = paragraphs.getLastOrNullObject();

          for (let i = 1; i < items.length; i++) {
            para = wordlist.insertParagraph(items[i], "End");
            if (type == "letter" && i < items.length - 1) {
              para.insertHtml("<br></br>", Word.InsertLocation.end);
            }
            para = paragraphs.getLastOrNullObject();
            para.font.name = "Arial Narrow";

            // wordlist.insertParagraph('','After');
            // para = paragraphs.getLastOrNullObject();
          }

          wordlist.insertParagraph("", "After");
          para = paragraphs.getLastOrNullObject();
          // para.spaceAfter = 0;
        } else {
          for (const item of items) {
            para.insertText(item, "End");
            para.font.underline = "Single";

            para.insertParagraph("", "After");
            para = paragraphs.getLastOrNullObject();
            para.font.underline = "None";

            if (section.list.sublist) {
              let subitems = section.list.sublist.items;
              let subtype = section.list.sublist.type;

              let sublist = para.startNewList();
              sublist.load();
              await context.sync();

              para.insertText(subitems[0], "End");
              para = paragraphs.getLastOrNullObject();
              sublist.setLevelNumbering(0, Word.ListNumbering.arabic, [0, "."]);

              for (let k = 1; k < subitems.length; k++) {
                sublist.insertParagraph(subitems[k], "End");
              }

              sublist.insertParagraph("", "After");
              para = paragraphs.getLastOrNullObject();
              para.leftIndent = 0;
            }
          }
          para = paragraphs.getLastOrNullObject();
        }
      } else {
        para = paragraphs.getLastOrNullObject();
      }
      if (section.endText) {
        para.insertText(section.endText, "End");
        para = paragraphs.getLastOrNullObject();
        para.insertParagraph("", "After");
      }
    }
    console.log("done");
  });
}
