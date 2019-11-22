/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */


const DAYS: string[] = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

const MONTHS: string[] = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];

const OPTIONS: any = {
  sheets: {
    repertoire: "Repertoire",
    recitalPlanning: "Recital Planning",
    recitalHistory: "Recital History",
    workflowLog: "Workflow Log",
    workflowOptions: "Workflow Options"
  },

  // Range names to get workflow options from "Workflow Options Sheet"
  fields: {
    performer: "OpPerformer", // Performer Name
    updateDates: "OpUpdateDates", // Y or N
    addToHistory: "OpAddToHistory", // Y or N
    dateColumn: "OpDateColumn", // Leftmost column of performance history columns in the Repertoire sheet
    recitalCount: "OpRecitalCount", // Number of recitals to loop through on the planning sheet
    repField: "OpRepFieldName", // Template for repertoire fields - Ex: Recital{{index}}
    dateField: "OpDateFieldName", // Template for date fields
    v12Field: "OpV12FieldName", // Template for V12 fields
    v2Field: "OpV2FieldName" // Template for V2 fields
  },
  ranges: {}
};

interface CompositionRange {
  // A row from the Recital Planning worksheet:
  // ["1.", "a.", 87, "Carillon de Westminter", "Louis Vierne", 7, "25", "23A"]

  0: string; // number
  1: string; // letter
  2: number; // ID
  3: string; // Title
  4: string; // Composer
  5: number; // Length
  6: string; // TAB
  7: string; // CC
}

class Composition {
  number: string;
  letter: string;
  id: number;
  title: string;
  composer: string;
  length: number;
  tab: string;
  cc: string;

  constructor(range: CompositionRange) {
    this.number = range[0];
    this.letter = range[1];
    this.id = range[2];
    this.title = range[3];
    this.composer = range[4];
    this.length = range[5];
    this.tab = range[6];
    this.cc = range[7];
  }
}

class Recital {
  dateStamp: number;
  dateString: string;
  date: Date;
  performer: string;
  program: string;
  repertoire: Composition[];
  venue12: string;
  venue2: string;

  constructor(dateStamp: number, performer: string, venue12: string, venue2: string, range: CompositionRange[]) {
    this.dateStamp = dateStamp;
    this.date = this.dateStampToDate(this.dateStamp);
    this.dateString = this.dateStamptoString(this.dateStamp);
    this.performer = performer;
    this.venue12 = venue12;
    this.venue2 = venue2;

    this.repertoire = [];

    // Read items into compositions array
    range.forEach((row) => {
      if (row[3] != "") {
        let composition = new Composition(row);
        this.repertoire.push(composition);
      }
    });

    // Filter out unnecessary "a." designations (where there is no "b.")
    this.repertoire.forEach((item, index, array) => {
      // If this is the last piece on the recital and the letter is "a."
      if (index === array.length - 1 && item.letter === "a.") {
        item.letter = "";
      }

      // If this is not the last piece on the recital, the letter is "a.", and the next letter is also "a."
      // (i.e., only one composition in this letter, so no letter necessary)
      else if (index < array.length - 1 && item.letter === "a." && array[index + 1].letter === "a.") {
        item.letter = "";
      }
    });

    this.program = this.programToString();
  } // End constructor

  // Converts Excel datestamp to JS Date
  dateStampToDate(dateStamp: number): Date {
    let date: Date = new Date((dateStamp - (25567 + 2)) * 86400 * 1000);
    return date;
  }

  // Converts Excel datestamp to YYYY-MM-DD string
  dateStamptoString(dateStamp: number): string {
    let date: Date = this.dateStampToDate(dateStamp);
    let dateString: string = date.toISOString().split("T")[0];
    return dateString;
  }

  // Returns the complete program as a string
  // Example:
  //
  // Brian Mathias
  // Monday, November 4, 2019
  // 12:00 - Tabernacle
  // 2:00 - Conference Center
  //
  // 1. Venite! - John Leavitt
  // 2. a. Flute Solo - Thomas Arne
  //    b. My Shepherd Will Supply My Need - Dale Wood
  // 3. a. Come, Come Ye Saints - arr. by organist
  //    b. An old melody - arr. by organist
  // 4. Carillon de Westminster - Louis Vierne

  programToString(): string {
    let dateString =
      DAYS[this.date.getUTCDay()] +
      ", " +
      MONTHS[this.date.getUTCMonth()] +
      " " +
      this.date.getUTCDate() +
      ", " +
      this.date.getUTCFullYear();

    let program: string = "";
    program += this.performer + "\n";
    program += dateString + "\n";

    if (this.venue12 !== "" && this.venue12 !== "None") {
      program += "12:00 - " + this.venue12 + "\n";
    }

    if (this.venue2 !== "" && this.venue2 !== "None") {
      program += "2:00 - " + this.venue2 + "\n";
    }

    program += "\n";

    for (let composition of this.repertoire) {
      let line: string = "";

      if (composition.number !== "") {
        line += composition.number + " ";
      } else {
        line += "   ";
      }

      if (composition.letter !== "") {
        line += composition.letter + " ";
      }

      line += composition.title + " - " + composition.composer + "\n";
      program += line;
    }

    return program;
  }
} // End class

async function filter1() {
  await filter(1);
}

async function filter2() {
  await filter(2);
}

async function filter3() {
  await filter(3);
}

async function filter4() {
  await filter(4);
}

async function filterClear() {
  await filter(-1);
}

async function filter(index: number) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Repertoire");
    const repertoireTable = sheet.tables.getItem("RepertoireList");

    let columns = ["O1", "O2", "O3", "O4"];

    // Clear any current filters on recital order columns
    for (let column of columns) {
      let filter = repertoireTable.columns.getItem(column).filter.clear();
    }

    if (index > 0 && index < 5) {
      let filterColumn = "O" + index.toString();
      let filter = repertoireTable.columns.getItem(filterColumn).filter;
      filter.apply({
        filterOn: Excel.FilterOn.values,
        values: ["1"]
      });
    }

    return context.sync();
  });
}

async function run() {
  await Excel.run(async (context) => {
    let logNode = document.getElementById("log");
    logNode.innerText = "";

    function log(message: string) {
      let node = document.createElement("p");
      node.innerText = message;
      let log = document.getElementById("log");
      log.appendChild(node);
    }

    log("Starting workflow...");

    // Global variables
    let recitalRanges = [];
    let recitals: Recital[] = [];

    // Initialize worksheets
    let sheets: {
      repertoire: Excel.Worksheet;
      recitalPlanning: Excel.Worksheet;
      recitalHistory: Excel.Worksheet;
      workflowLog: Excel.Worksheet;
      workflowOptions: Excel.Worksheet;
    } = {
      repertoire: context.workbook.worksheets.getItem(OPTIONS.sheets.repertoire),
      recitalPlanning: context.workbook.worksheets.getItem(OPTIONS.sheets.recitalPlanning),
      recitalHistory: context.workbook.worksheets.getItem(OPTIONS.sheets.recitalHistory),
      workflowLog: context.workbook.worksheets.getItem(OPTIONS.sheets.workflowLog),
      workflowOptions: context.workbook.worksheets.getItem(OPTIONS.sheets.workflowOptions)
    };

    // Load values of workflow options
    Object.keys(OPTIONS.fields).forEach((key, index, array) => {
      OPTIONS.ranges[key] = sheets.workflowOptions.getRange(OPTIONS.fields[key]);
      OPTIONS.ranges[key].load("values");
    });

    await context.sync();

    // Store values from options fields in OPTIONS object
    Object.keys(OPTIONS.ranges).forEach((key, index, array) => {
      OPTIONS[key] = OPTIONS.ranges[key].values[0][0];
    });

    async function createRecitals() {
      // Loop through the number of recitals specified in workflow options (starting at 1 - not zero!)
      for (let i = 1; i < OPTIONS.recitalCount + 1; i++) {
        // Create field names from templates on workflow options sheet
        let fields: any = {};
        fields.rep = OPTIONS.repField.replace("{{index}}", i); // Recital1, etc.
        fields.date = OPTIONS.dateField.replace("{{index}}", i); // Recital1D etc.
        fields.v12 = OPTIONS.v12Field.replace("{{index}}", i); // Recital1V12, etc.
        fields.v2 = OPTIONS.v2Field.replace("{{index}}", i); // Recital1V2, etc.

        // Get ranges
        let ranges: any = {};
        ranges.rep = sheets.recitalPlanning.getRange(fields.rep);
        ranges.date = sheets.recitalPlanning.getRange(fields.date);
        ranges.v12 = sheets.recitalPlanning.getRange(fields.v12);
        ranges.v2 = sheets.recitalPlanning.getRange(fields.v2);

        // Load values
        ranges.rep.load("values");
        ranges.date.load("values");
        ranges.v12.load("values");
        ranges.v2.load("values");

        // Push ranges object to recitalRanges array
        recitalRanges.push(ranges);
      }

      // Load all queued values
      await context.sync();

      // Loop through recitalRanges - if a recital is present, create Recital object and add it to the recitals array
      recitalRanges.forEach((item, index, array) => {
        // If the recital has a value in the date field
        if (item.date.values[0][0] !== "") {
          let recital = new Recital(
            item.date.values[0][0],
            OPTIONS.performer,
            item.v12.values[0][0],
            item.v2.values[0][0],
            item.rep.values
          );
          recitals.push(recital);
        }
      });
      log(recitals.length + " recitals found.");
    } // End createRecitals()

    // Update Recital Dates
    async function updatePerformanceDates(recitals: Recital[]) {
      const repertoireList = sheets.repertoire.getRange("RepertoireList");
      repertoireList.load("values, rowIndex");
      await context.sync();

      // First row of the table
      let firstRow = repertoireList.rowIndex;

      // First column of dates
      let firstCol: number = OPTIONS.dateColumn;

      // Row index of composition relative to table
      let tableRowIndex: number;

      // Row index of composition relative to worksheet
      let sheetRowIndex: number;

      // A row from repertoire list
      let row;

      for (let recital of recitals) {
        repertoireList.load("values, rowIndex");
        await context.sync();

        for (let composition of recital.repertoire) {
          // Find the row number of the composition and assign it to tableRowIndex
          for (let i = 0; i < repertoireList.values.length; i++) {
            // If the first column contains the ID number of the composition
            if (repertoireList.values[i][0] === composition.id) {
              row = repertoireList.values[i];
              tableRowIndex = i;
              break;
            }
          }

          // First row of RepertoireList table + table index of composition
          sheetRowIndex = firstRow + tableRowIndex;

          let values = [];
          values.push(recital.dateStamp);
          values.push(row[firstCol]);
          values.push(row[firstCol + 1]);
          values.push(row[firstCol + 2]);

          // Excel expects a two-dimensional array to set values for any range
          let arr = [];
          arr.push(values);

          const range = sheets.repertoire.getRangeByIndexes(sheetRowIndex, firstCol, 1, 4);
          range.values = arr;
        }

        await context.sync();
      }
      log("Performance dates updated.");
    } // End updatePerformanceDates()

    async function addToRecitalHistory() {
      for (let recital of recitals) {
        let values = [];
        let formats = [];

        // Set values and number formats for each row

        /// Row 0 - Recital Date
        let recitalDate =
          recital.date.getUTCMonth() + 1 + "/" + recital.date.getUTCDate() + "/" + recital.date.getUTCFullYear();

        values.push([recitalDate, "", "", "", "", ""]);
        formats.push(["m/d/yyyy", "", "", "", "", ""]);

        /// Row 1 - 12:00 Venue
        values.push(["12:00", recital.venue12, "", "", "", ""]);
        formats.push(["@", "@", "", "", "", ""]);

        /// Row 2 - 2:00 Venue
        values.push(["2:00", recital.venue2, "", "", "", ""]);
        formats.push(["@", "@", "", "", "", ""]);

        /// Row 3 - Header Row
        values.push(["ON", "OL", "ID", "Title", "Composer", "Length"]);
        formats.push(["Text", "Text", "Text", "Text", "Text", "Text"]);

        /// Body Rows (Rows 4-(4 + recital.repertoire.length))
        for (let composition of recital.repertoire) {
          values.push([
            composition.number,
            composition.letter,
            composition.id,
            composition.title,
            composition.composer,
            composition.length
          ]);
          formats.push(["@", "", "", "", "", "0.00"]);
        }

        /// Total Row (formula added later)
        values.push(["", "", "", "", "", ""]);
        formats.push(["", "", "", "", "", ""]);

        // Get ranges and add values and number formats to worksheet
        let sheet = context.workbook.worksheets.getItem("Recital History");
        let used = sheet.getUsedRange().getLastRow();
        used.load("address,rowIndex");
        await context.sync();

        // Start new recital 3 rows below the used range
        let range = sheet.getRangeByIndexes(used.rowIndex + 4, 0, values.length, 6);

        // Apply values and number formats to the range
        range.numberFormat = formats;
        range.values = values;

        // Row numbers for getting ranges to apply formatting
        let dateRowIndex = used.rowIndex + 4;
        let v12RowIndex = used.rowIndex + 5;
        let v2RowIndex = used.rowIndex + 6;
        let firstBodyRowIndex = used.rowIndex + 8;
        let lastBodyRowIndex = firstBodyRowIndex + recital.repertoire.length;

        // Format Row 0
        let dateRow = range.getRow(0);
        dateRow.merge();
        dateRow.format.horizontalAlignment = "Right";
        dateRow.format.font.bold = true;
        dateRow.format.font.size = 14;
        dateRow.format.borders.getItem("EdgeBottom").color = "000000";

        // Format Row 1
        let v12 = sheet.getRangeByIndexes(v12RowIndex, 0, 1, 1);
        v12.format.font.bold = true;
        let v12Value = sheet.getRangeByIndexes(v12RowIndex, 1, 1, 5);
        v12Value.merge();

        // Format Row 2
        let v2 = sheet.getRangeByIndexes(v2RowIndex, 0, 1, 1);
        v2.format.font.bold = true;
        let v2Value = sheet.getRangeByIndexes(v2RowIndex, 1, 1, 5);
        v2Value.merge();

        // Format Row 4
        let headerRow = range.getRow(3);
        headerRow.format.fill.color = "000000";
        headerRow.format.font.color = "FFFFFF";
        headerRow.format.font.bold = true;
        headerRow.format.horizontalAlignment = "Center";

        // Format Body Rows
        for (let i = firstBodyRowIndex; i < lastBodyRowIndex; i++) {
          let row = sheet.getRangeByIndexes(i, 0, 1, 6);

          if (i % 2 === 0) {
            row.format.fill.color = "DADADA";
          }

          row.format.borders.getItem("EdgeTop").color = "000000";
          row.format.borders.getItem("EdgeBottom").color = "000000";
          row.format.borders.getItem("EdgeLeft").color = "000000";
          row.format.borders.getItem("EdgeRight").color = "000000";
        }

        // Right Align Composer Names
        let composerColumn = sheet.getRangeByIndexes(firstBodyRowIndex, 4, recital.repertoire.length, 2);
        composerColumn.format.horizontalAlignment = "Right";

        // Add formula and format to time total
        let totalCell = sheet.getRangeByIndexes(lastBodyRowIndex, 5, 1, 1);
        let formula = "=SUM(F" + (firstBodyRowIndex + 1) + ":F" + lastBodyRowIndex + ")";

        totalCell.formulas = [[formula]];
        totalCell.format.font.bold = true;
        totalCell.format.horizontalAlignment = "Right";
        totalCell.numberFormat = [["0.00"]];

        await context.sync();
      }
      log(recitals.length + " recitals added to recital history.");
    }

    async function addToWorkflowLog() {
      let data: any = {};
      let today = new Date();

      // Format today as YYYY-MM-DD
      let todayString = today.getFullYear() + "-";

      if (today.getMonth() + 1 < 10) {
        todayString += "0";
      }
      todayString += today.getMonth() + 1 + "-";

      if (today.getDate() < 10) {
        todayString += "0";
      }
      todayString += today.getDate();

      // Create string of recital dates in MM/DD, MM/DD format
      let recitalDates: string = "";

      for (let i = 0; i < recitals.length; i++) {
        recitalDates += recitals[i].date.getUTCMonth() + 1 + "/" + recitals[i].date.getUTCDate();
        if (i < recitals.length - 1) {
          recitalDates += ", ";
        }
      }

      data.date = todayString;
      data.recitals = recitals;
      let json: string = JSON.stringify(data);

      let logTable = sheets.workflowLog.tables.getItem("WorkflowLog");
      logTable.rows.add(0, [[todayString, recitalDates, json, ""]]);

      log("JSON added to workflow log.");
      return context.sync();
    }

    // Run workflow functions in order

    await createRecitals();

    if (OPTIONS.updateDates === "Y") {
      await updatePerformanceDates(recitals);
    }

    if (OPTIONS.addToHistory === "Y") {
      await addToRecitalHistory();
    }

    await addToWorkflowLog();
    log("Workflow complete.");
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    //console.error(error);
    let log = document.getElementById("log");
    let node = document.createElement("p");
    node.innerText = error;
    node.style.color = "red";
    log.appendChild(node);
  }
}

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
    document.getElementById("filter1").onclick = filter1;
    document.getElementById("filter2").onclick = filter2;
    document.getElementById("filter3").onclick = filter3;
    document.getElementById("filter4").onclick = filter4;
    document.getElementById("filterClear").onclick = filterClear;
  }
});