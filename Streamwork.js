function sendApprovalEmail(e) {
  var sheet = e.source.getActiveSheet();
  var row = e.range.getRow();
  var approverColumnIndex = 11; // Approver column index (assuming 11th column)
  var itemDescriptionColumnIndex = 2; // Item Description column index (assuming 2nd column)

  var approverEmail = sheet.getRange(row, approverColumnIndex).getValue();
  var itemDescription = sheet.getRange(row, itemDescriptionColumnIndex).getValue();
  var emailBody = `Please approve the procurement for: ${itemDescription}`;

  MailApp.sendEmail(approverEmail, "Procurement Approval Needed", emailBody);
}

function onEdit(e) {
  sendApprovalEmail(e);
}

function sendReminders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Procurement");
  var data = sheet.getDataRange().getValues();
  var now = new Date();

  for (var i = 1; i < data.length; i++) {
    var status = data[i][9]; // Status column index (assuming 10th column)
    var approverEmail = data[i][10]; // Approver column index (assuming 11th column)
    var itemDescription = data[i][1]; // Item Description column index (assuming 2nd column)
    var orderDate = new Date(data[i][6]); // Order Date column index (assuming 7th column)
    var expectedDeliveryDate = new Date(data[i][7]); // Expected Delivery Date column index (assuming 8th column)

    if (status === 'Ordered' && (now - orderDate > 3 * 24 * 60 * 60 * 1000)) { // Reminder if not approved within 3 days
      var emailBody = `Reminder: Please approve the procurement for: ${itemDescription}`;
      MailApp.sendEmail(approverEmail, "Procurement Approval Reminder", emailBody);
    }
  }
}

function updateStatus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetByName("Procurement");
  var data = sheet.getDataRange().getValues();
  var now = new Date();

  for (var i = 1; i < data.length; i++) {
    var status = data[i][9]; // Status column index (assuming 10th column)
    var actualDeliveryDate = data[i][8]; // Actual Delivery Date column index (assuming 9th column)

    if (status === 'In Transit' && actualDeliveryDate) {
      sheet.getRange(i + 1, 10).setValue('Delivered');
    }
    if (status === 'Delivered' && !actualDeliveryDate) {
      sheet.getRange(i + 1, 10).setValue('Completed');
    }
  }
}


function notifyRiskOwner() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Procurement Risk");
  var data = sheet.getDataRange().getValues();
  var likelihoodColumnIndex = 3; // Replace with actual index of the likelihood column
  var riskOwnerColumnIndex = 6; // Replace with actual index of the risk owner email column
  var riskDescriptionColumnIndex = 1; // Replace with actual index of the risk description column
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][likelihoodColumnIndex] == "High") {
      var ownerEmail = data[i][riskOwnerColumnIndex];
      var riskDescription = data[i][riskDescriptionColumnIndex];
      var emailSubject = "High Risk Notification";
      var emailBody = `The following risk is high: ${riskDescription}`;

      Logger.log(`Sending email to: ${ownerEmail}`);
      Logger.log(`Email subject: ${emailSubject}`);
      Logger.log(`Email body: ${emailBody}`);
      
      MailApp.sendEmail(ownerEmail, emailSubject, emailBody);
    }
  }
}

function onOpen() {
  notifyRiskOwner();
}

function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice Automation");
  if (!sheet) {
    Logger.log('Sheet "Invoice Automation" not found.');
    return;
  }

  var row = e.range.getRow();
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('Row data: ' + data.join(', '));

  var timestamp = new Date(data[0]);
  var invoiceID = data[1];
  var supplierName = data[2];
  var supplierEmail = data[3];
  var invoiceDate = new Date(data[4]);
  var invoiceAmount = data[5];
  var poNumber = data[6];
  var paymentStatus = data[7];
  var approvalDate = new Date(data[8]);
  var comments = data[9];
  var submissionDate = new Date(data[10]);

  var formattedInvoiceAmount = formatCurrency(invoiceAmount);
  sheet.getRange(row, 6).setValue(formattedInvoiceAmount);

  Logger.log('Payment Status: "' + paymentStatus + '"');
  Logger.log('Invoice Amount: "' + invoiceAmount + '"');

  // Approve invoice if conditions met
  if (paymentStatus === 'Pending') {
    sheet.getRange(row, 8).setValue('Approved');
    sheet.getRange(row, 9).setValue(new Date());
    Logger.log('Approving invoice and sending email notification...');
    sendEmailNotification(supplierName, invoiceID, 'Approved', supplierEmail);
  } else {
    Logger.log('Conditions not met for approval for row ' + row);
  }
}

function formatCurrency(amount) {
  // Ensure the amount is treated as a number and format it with the "RM" prefix
  var numericAmount = parseFloat(amount);
  if (isNaN(numericAmount)) {
    Logger.log('Invalid amount: ' + amount);
    return amount; // Return the original value if it's not a valid number
  }
  return 'RM' + numericAmount.toFixed(2); // Format with 2 decimal places
}

function sendEmailNotification(supplier, invoiceID, status,supplierEmail) {
  var emailAddress = supplierEmail; // Replace with your email or dynamic logic
  var subject = 'Invoice ' + invoiceID + ' ' + status;
  var message = 'The invoice from ' + supplier + ' with ID ' + invoiceID + ' has been ' + status + '.';

  Logger.log('Sending email to: ' + emailAddress);
  Logger.log('Subject: ' + subject);
  Logger.log('Message: ' + message);

  MailApp.sendEmail(emailAddress, subject, message);
}

function onFormSubmit2(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Receipt Digitization");
  const row = sheet.getLastRow();
  
  // Extract form data
  const timestamp = sheet.getRange(row, 1).getValue(); // Timestamp
  const receiptId = sheet.getRange(row, 2).getValue(); 
  const receiptFileUrl = sheet.getRange(row, 3).getFormula() || sheet.getRange(row, 3).getValue(); // Receipt Image URL
  const purchaseDate = sheet.getRange(row, 4).getValue(); // Purchase Date
  const supplierName = sheet.getRange(row, 5).getValue(); // Supplier Name
  const amount = sheet.getRange(row, 6).getValue(); //
  const category = sheet.getRange(row, 7).getValue(); // Category
  const uploadedBy = sheet.getRange(row, 8).getValue(); // Uploaded By
  const comments = sheet.getRange(row, 10).getValue(); // Comments

  // Log the receiptFileUrl for debugging
  Logger.log('receiptFileUrl: ' + receiptFileUrl);
  
  // Extract file ID from the URL
  const fileIdMatch = receiptFileUrl.match(/[-\w]{25,}/);
  if (fileIdMatch && fileIdMatch.length > 0) {
    const fileId = fileIdMatch[0];
    Logger.log('fileId: ' + fileId);
    const file = DriveApp.getFileById(fileId);
    
    // Move file to the appropriate category folder
    const folderId = getFolderId(category);
    const folder = DriveApp.getFolderById(folderId);
    file.moveTo(folder);
    
    // Update sheet with additional information
    const uploadDate = new Date();
    sheet.getRange(row, 2).setValue(receiptId); // Adjusted column index
    sheet.getRange(row, 9).setValue(uploadDate); // Adjusted column index
    sheet.getRange(row, 4).setValue(purchaseDate); // Placeholder for Purchase Date (adjusted column index)
    sheet.getRange(row, 5).setValue(supplierName); // Placeholder for Supplier Name (adjusted column index)
    var formattedInvoiceAmount = formatCurrency(amount);
    sheet.getRange(row, 6).setValue(formattedInvoiceAmount);
  } else {
    Logger.log('Error: Unable to extract file ID from URL');
  }
}

function getFolderId(category) {
  const folders = {
    "Office Supplies": "1zpPli1R12C6cNfVW4MVdaIvDIYLXB3P2",
    "IT Equipment": "1TaGM9jZYNT_8O9bamFY37fNj3s2JxaxD",
    // Add more categories and their IDs as needed
  };
  return folders[category];
}