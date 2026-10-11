function doPost(e) {
  try {
    return handlePost(e);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handlePost(e) {

  const ss = SpreadsheetApp.openById("1tv9myqUU_WjhKTa8Vu1kJChnW9M8ZkEHoXGlwz-rU7A");
  const submissionDate = new Date();

  // Create Month-Year sheet
  const sheetName = Utilities.formatDate(
    submissionDate,
    Session.getScriptTimeZone(),
    "MMMM yyyy"
  );

  let sh = ss.getSheetByName(sheetName);

  // Create sheet + header if not exists
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow([
      "Timestamp",
      "Email ID",
      "Date",
      "Shift",
      "Transporter Name",
      "Transporter Mobile",
      "Vehicle No",
      "Location",
      "Plant",
      "Company",
      "Product",
      "Batch Number",
      "Pack Size (KG)",
      "No. of Bags",
      "Quantity (KG)",
      "Quantity (MT)",
      "Invoice (BSW)",
      "Company Invoice",
      "Loaded Party Name",
      "Loading Start Time",
      "Unloading End Time",
      "Total Work Hours",
      "Remark"
    ]);
    sh.getRange(1, 1, 1, 23).setFontWeight("bold");
  }

  const data = JSON.parse(e.postData.contents);

  const productionDate = Utilities.formatDate(
    new Date(data.date),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  // Split batch numbers into array
  const batchArray = data.batchNumbers
    .split(",")
    .map(b => b.trim())
    .filter(b => b !== "");

  // Append ONE ROW PER BATCH (MAIN CHANGE)
  batchArray.forEach(batchNo => {
    sh.appendRow([
      submissionDate,
      data.email,
      productionDate,
      data.shift,
      data.driv,
      data.drivmn,
      data.truck,
      data.location,
      data.plant,
      data.company,
      data.product,
      batchNo,
      data.packSize,
      data.bags,
      data.qtyKg,
      data.qtyMt,
      data.invoice,
      data.cinvoice,
      data.partyName,
      data.loadStart,
      data.unloadEnd,
      data.workHours,
      data.remark
    ]);
  });

  // Formatting for newly added rows
  const lastRow = sh.getLastRow();
  const rowsAdded = batchArray.length;

  sh.getRange(lastRow - rowsAdded + 1, 1, rowsAdded)
    .setNumberFormat("dd/MM/yyyy HH:mm:ss");

  sh.getRange(lastRow - rowsAdded + 1, 3, rowsAdded)
    .setNumberFormat("dd/MM/yyyy");

  // ================= EMAIL PART =================

  const subject = `Dispatch Form Submission - ${Utilities.formatDate(
    submissionDate,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss"
  )}`;

  const body =
`Dispatch Entry Submitted

Timestamp: ${Utilities.formatDate(submissionDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")}
Email ID: ${data.email}
Date: ${productionDate}
Shift: ${data.shift}
Driver Name: ${data.driv}
Driver Mobile: ${data.drivmn}
Vehicle No: ${data.truck}
Location: ${data.location}
Plant: ${data.plant}
Company: ${data.company}
Product: ${data.product}
Batch Numbers: ${data.batchNumbers}
Pack Size (KG): ${data.packSize}
No. of Bags: ${data.bags}
Quantity (KG): ${data.qtyKg}
Quantity (MT): ${data.qtyMt}
Invoice (BSW): ${data.invoice}
Company Invoice: ${data.cinvoice}
Loaded Party Name: ${data.partyName}
Loading Start Time: ${data.loadStart}
Unloading End Time: ${data.unloadEnd}
Total Work Hours: ${data.workHours}
Remark: ${data.remark}
`;

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    body: body
  });

  // ================= RESPONSE =================

  const output = ContentService.createTextOutput(
    JSON.stringify({
      status: "success",
      message: "Data saved successfully (batch-wise entries created)"
    })
  );

  output.setMimeType(ContentService.MimeType.JSON);
  output.setHeader("Access-Control-Allow-Origin", "*");
  output.setHeader("Access-Control-Allow-Methods", "POST");

  return output;
}
