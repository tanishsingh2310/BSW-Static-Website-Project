function doPost(e) {
  try {
    return handlePost(e);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handlePost(e) {

  const ss = SpreadsheetApp.openById("1kjfAN2S8L91ecGH7P2PCMyyBWXDkzS1POFrt9cbr4tQ");

  const submissionDate = new Date();

  // Month-Year Sheet
  const sheetName = Utilities.formatDate(
    submissionDate,
    Session.getScriptTimeZone(),
    "MMMM yyyy"
  );

  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow([
      "Timestamp",
      "Email ID",
      "Date",
      "Shift",
      "Plant",
      "Company",
      "Product",
      "Mixture No",
      "Labour Party",
      "Start Time",
      "End Time",
      "Total Working Hours",
      "Batch Number",
      "Batch Size (Ton)",
      "Pack Size (KG)",
      "Standard Production ",
      "Actual Production ",
      "Balance Quantity ",
      "Quantity (KG)",
      "Quantity (MT)",
      "Remark"
    ]);
    sh.getRange(1, 1, 1, 21).setFontWeight("bold");
  }

  const data = JSON.parse(e.postData.contents);

  const productionDate = Utilities.formatDate(
    new Date(data.date),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  // ================= SAVE DATA =================

  sh.appendRow([
    submissionDate,
    data.email,
    productionDate,
    data.shift,
    data.plant,
    data.company,
    data.product,
    data.mixture,
    data.labourParty,
    data.productionStartTime,
    data.productionEndTime,
    data.totalWorkingHours,
    data.batch,
    data.batchSizeTon,
    data.packSizeKg,
    data.standardProductionBags,
    data.actualProductionBags,
    data.balanceQuantityBags,
    data.totalQtyKg,
    data.totalQtyMt,
    data.remark
  ]);

  const lastRow = sh.getLastRow();
  sh.getRange(lastRow, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange(lastRow, 3).setNumberFormat("dd/MM/yyyy");

  // ================= EMAIL PART =================

  const subject = `Production Form Submission - ${Utilities.formatDate(
    submissionDate,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss"
  )}`;

  const body =
`Production Entry Submitted

Timestamp: ${Utilities.formatDate(submissionDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")}
Email ID: ${data.email}
Date: ${productionDate}
Shift: ${data.shift}
Plant: ${data.plant}
Company: ${data.company}
Product: ${data.product}
Mixture No: ${data.mixture}

Labour Party: ${data.labourParty}
Start Time: ${data.productionStartTime}
End Time: ${data.productionEndTime}
Total Working Hours: ${data.totalWorkingHours}

Batch: ${data.batch}
Batch Size (Ton): ${data.batchSizeTon}
Pack Size (KG): ${data.packSizeKg}

Standard Production (Bags): ${data.standardProductionBags}
Actual Production (Bags): ${data.actualProductionBags}
Balance Quantity (Bags): ${data.balanceQuantityBags}

Quantity (KG): ${data.totalQtyKg}
Quantity (MT): ${data.totalQtyMt}

Remark: ${data.remark}
`;

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    body: body
  });

  // ================= RESPONSE =================

  const output = ContentService.createTextOutput(
    JSON.stringify({ status: "success", message: "Data saved and email sent" })
  );

  output.setMimeType(ContentService.MimeType.JSON);
  output.setHeader("Access-Control-Allow-Origin", "*");
  output.setHeader("Access-Control-Allow-Methods", "POST");

  return output;
}
