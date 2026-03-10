/**
 * Google Sheets Data Cleaning (Portfolio Project)
 * Reads data from: Raw_Data
 * Writes cleaned data to: Clean_Data
 * Includes logging, validation, summary stats and UI interaction.
 */


/**
 * Main cleaning logic
 */

function cleanData_withLogging() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const source = ss.getSheetByName("Raw_Data");
  const dest = ss.getSheetByName("Clean_Data") || ss.insertSheet("Clean_Data");
  const logSheet = ss.getSheetByName("Cleaning_Log") || ss.insertSheet("Cleaning_Log");

  if (!source) {
    throw new Error('Sheet "Raw_Data" not found.');
  }

  dest.clearContents();

  const data = source.getDataRange().getValues();

  if (data.length === 0) {
    throw new Error("Source sheet is empty.");
  }

  const nameAliases = ["name", "full name", "customer name", "client name"];
  const emailAliases = ["email", "e-mail", "mail"];
  const salesAliases = ["sales", "amount", "revenue", "value", "total sales"];

  function findHeaderIndex(headerRow, aliases) {
    for (let i = 0; i < headerRow.length; i++) {
      if (aliases.includes(headerRow[i])) {
        return i;
      }
    }
    return -1;
  }

  // --- FIND HEADER ROW AUTOMATICALLY ---
  let headerRowIndex = -1;
  let nameIndex = -1;
  let emailIndex = -1;
  let salesIndex = -1;

  for (let r = 0; r < data.length; r++) {

    const headers = data[r].map(h => String(h).trim().toLowerCase());

    const n = findHeaderIndex(headers, nameAliases);
    const e = findHeaderIndex(headers, emailAliases);
    const s = findHeaderIndex(headers, salesAliases);

    if (n !== -1 && e !== -1 && s !== -1) {
      headerRowIndex = r;
      nameIndex = n;
      emailIndex = e;
      salesIndex = s;
      break;
    }
  }

  if (headerRowIndex === -1) {
    throw new Error("Could not detect header row automatically.");
  }

  let totalRows = data.length - (headerRowIndex + 1);
  let keptRows = 0;

  const output = [];
  output.push(["Name", "Email", "Sales"]);

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Timestamp", "Row", "Reason", "Name", "Email", "Sales"]);
  }

  for (let i = headerRowIndex + 1; i < data.length; i++) {

    const nome = String(data[i]?.[nameIndex] ?? "").trim();
    const email = String(data[i]?.[emailIndex] ?? "").trim();
    const venditeRaw = data[i]?.[salesIndex];

    let reason = "";

    if (!nome || !email) {
      reason = "Missing Name or Email";
    }

    const cleaned = String(venditeRaw ?? "")
      .trim()
      .replace(/[^\d.,-]/g, "");

    const num = Number(cleaned.replace(",", "."));

    if (!reason && isNaN(num)) {
      reason = "Invalid number";
    }

    if (!reason && num < 100) {
      reason = "Below threshold";
    }

    if (reason) {
      logSheet.appendRow([
        new Date(),
        i + 1,
        reason,
        nome,
        email,
        venditeRaw
      ]);
      continue;
    }

    output.push([nome, email, num]);
    keptRows++;
  }

  dest.getRange(1, 1, output.length, 3).setValues(output);

  Logger.log(`Header row detected at: ${headerRowIndex + 1}`);
  Logger.log(`Total rows: ${totalRows}`);
  Logger.log(`Valid rows: ${keptRows}`);

  return {
    totalRows,
    keptRows
  };
}