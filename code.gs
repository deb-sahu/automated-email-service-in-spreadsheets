
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Fill');
  menu.addItem('Offer Letter', 'createNewGoogleDocs1').addItem('Letter of Appointment', 'createNewGoogleDocs2').addItem('Letter of Confirmation', 'createNewGoogleDocs3').addItem('Appraisal Letter', 'createNewGoogleDocs4')
  menu.addToUi();
  var ui2 = SpreadsheetApp.getUi();
  ui2.createMenu('Send').addItem('Offer Letter', 'sendPDFForm1').addItem('Letter of Appointment', 'sendPDFForm2').addItem('Letter of Confirmation', 'sendPDFForm3').addItem('Appraisal Letter', 'sendPDFForm4').addToUi();

}



//Offer Letter

function createNewGoogleDocs1() {
  
  const googleDocTemplate = DriveApp.getFileById('1leF_acYbkTk3gv551qFLbT85qYWmIxAy-lUFBmV9f18');
  const destinationFolder = DriveApp.getFolderById('1CGCiKEFKlvKLBUSoR0L1ejxE90FhXK1p');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMPLOYEE MASTER')
  
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row, index){
    if (index === 0 ) return;
    if (row[0]==false) return;
    const copy = googleDocTemplate.makeCopy(`${row[5]} Name of Employee` , destinationFolder)
    const doc = DocumentApp.openById(copy.getId())
    const body = doc.getBody();
    const doj = new Date(row[11]).toLocaleDateString();
    
    body.replaceText('{{Name of Employee}}', row[5]);
    body.replaceText('{{Designation}}', row[6]);
    body.replaceText('{{Department}}', row[7]);
    body.replaceText('{{Branch}}', row[8]);
    body.replaceText('{{DOJ}}', doj);
    doc.saveAndClose();
    /*const pdfContentBlob=doc.getAs(MimeType.PDF);
    pdfFolder.createFile(pdfContentBlob).setName(`${row[1]} Employee Name pdf`)
    destinationFolder.removeFile(copy);
    const url = pdfContentBlob.getUrl(); */
    const url = doc.getUrl();
    sheet.getRange(index + 1, 116).setValue(url)
    
  })
}


function sendPDFForm1()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var row =sheet.getActiveCell().getRow();
  var currcol = 116;
  var lastCell = sheet.getRange(row, currcol);
  var lrow = lastCell.getValue();
  var fileid=lrow.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  sendEmailWithAttachment1(row,fileid);
}


function sendEmailWithAttachment1(row,fileid,copy)
{
  const pdfFolder = DriveApp.getFolderById('1CGCiKEFKlvKLBUSoR0L1ejxE90FhXK1p');
  const file = DriveApp.getFileById(fileid)
  
  var client = getClientInfo(row);
  
  var template = HtmlService.createTemplateFromFile('email-template1');
  template.client = client;
  var message = template.evaluate().getContent();
  
  
  MailApp.sendEmail({
    to: client.email,
    subject: "Offer Letter",
    htmlBody: message,
    attachments: [file.getAs(MimeType.PDF)]
  });
  
  
}

//Appointment Letter

function createNewGoogleDocs2() {
  
  const googleDocTemplate = DriveApp.getFileById('1s17uPTdwi17UNFbgOyv-t3bJ6XuJH-8ei_1ZkoGX-0c');
  const destinationFolder = DriveApp.getFolderById('1zXHiq5FAjaemQewDsdd5rP5tqtvNcYQj');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMPLOYEE MASTER')
  
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row, index){
    if (index === 0 ) return;
    if (row[0]==false) return;
    const copy = googleDocTemplate.makeCopy(`${row[5]} Name of Employee` , destinationFolder)
    const doc = DocumentApp.openById(copy.getId())
    const body = doc.getBody();
    const doj = new Date(row[11]).toLocaleDateString();
    
    body.replaceText('{{Name of Employee}}', row[5]);
    body.replaceText('{{Designation}}', row[6]);
    body.replaceText('{{Department}}', row[7]);
    body.replaceText('{{Branch}}', row[8]);
    body.replaceText('{{DOJ}}', doj);
    doc.saveAndClose();
    /*const pdfContentBlob=doc.getAs(MimeType.PDF);
    pdfFolder.createFile(pdfContentBlob).setName(`${row[1]} Employee Name pdf`)
    destinationFolder.removeFile(copy);
    const url = pdfContentBlob.getUrl(); */
    const url = doc.getUrl();
    sheet.getRange(index + 1, 117).setValue(url)
    
  })
}

function sendPDFForm2()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var row =sheet.getActiveCell().getRow();
  var currcol = 117;
  var lastCell = sheet.getRange(row, currcol);
  var lrow = lastCell.getValue();
  var fileid=lrow.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  sendEmailWithAttachment2(row,fileid);
}


function sendEmailWithAttachment2(row,fileid,copy)
{
  const pdfFolder = DriveApp.getFolderById('1zXHiq5FAjaemQewDsdd5rP5tqtvNcYQj');
  const file = DriveApp.getFileById(fileid)
  
  var client = getClientInfo(row);
  
  var template = HtmlService.createTemplateFromFile('email-template2');
  template.client = client;
  var message = template.evaluate().getContent();
  
  
  MailApp.sendEmail({
    to: client.email,
    subject: "Letter of Appointment",
    htmlBody: message,
    attachments: [file.getAs(MimeType.PDF)]
  });
  
  
}

//Confirmation Letter

function createNewGoogleDocs3() {
  
  const googleDocTemplate = DriveApp.getFileById('1Flt4CLb80pdhLX30qPrjb2Wl26JjePYMYvuRBM2fO6c');
  const destinationFolder = DriveApp.getFolderById('1oHOABooMVG12K6oEVZoqZTqIxO3WLP-e');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMPLOYEE MASTER')
  
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row, index){
    if (index === 0 ) return;
    if (row[0]==false) return;
    const copy = googleDocTemplate.makeCopy(`${row[5]} Name of Employee` , destinationFolder)
    const doc = DocumentApp.openById(copy.getId())
    const body = doc.getBody();
    const doj = new Date(row[11]).toLocaleDateString();
    
    body.replaceText('{{Name of Employee}}', row[5]);
    body.replaceText('{{Designation}}', row[6]);
    body.replaceText('{{Department}}', row[7]);
    body.replaceText('{{Branch}}', row[8]);
    body.replaceText('{{DOJ}}', doj);
    doc.saveAndClose();
    /*const pdfContentBlob=doc.getAs(MimeType.PDF);
    pdfFolder.createFile(pdfContentBlob).setName(`${row[1]} Employee Name pdf`)
    destinationFolder.removeFile(copy);
    const url = pdfContentBlob.getUrl(); */
    const url = doc.getUrl();
    sheet.getRange(index + 1, 118).setValue(url)
    
  })
}

function sendPDFForm3()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var row =sheet.getActiveCell().getRow();
  var currcol = 118;
  var lastCell = sheet.getRange(row, currcol);
  var lrow = lastCell.getValue();
  var fileid=lrow.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  sendEmailWithAttachment3(row,fileid);
}

function sendEmailWithAttachment3(row,fileid,copy)
{
  const pdfFolder = DriveApp.getFolderById('1oHOABooMVG12K6oEVZoqZTqIxO3WLP-e');
  const file = DriveApp.getFileById(fileid)
  
  var client = getClientInfo(row);
  
  var template = HtmlService.createTemplateFromFile('email-template3');
  template.client = client;
  var message = template.evaluate().getContent();
  
  
  MailApp.sendEmail({
    to: client.email,
    subject: "Letter of Confirmation",
    htmlBody: message,
    attachments: [file.getAs(MimeType.PDF)]
  });
  
  
}


//Appraisal Letter

function createNewGoogleDocs4() {
  
  const googleDocTemplate = DriveApp.getFileById('1SLuGrfYORYT-ENxTZI4UOKAI2cePvXDPN15uO8knS74');
  const destinationFolder = DriveApp.getFolderById('1jXfTezBVwiL6-WvWdrY7Xoa17OZi59J4');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMPLOYEE MASTER')
  
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row, index){
    if (index === 0 ) return;
    if (row[0]==false) return;
    const copy = googleDocTemplate.makeCopy(`${row[5]} Name of Employee` , destinationFolder)
    const doc = DocumentApp.openById(copy.getId())
    const body = doc.getBody();
    const doj = new Date(row[11]).toLocaleDateString();
    
    body.replaceText('{{Name of Employee}}', row[5]);
    body.replaceText('{{Designation}}', row[6]);
    body.replaceText('{{Department}}', row[7]);
    body.replaceText('{{Branch}}', row[8]);
    body.replaceText('{{DOJ}}', doj);
    doc.saveAndClose();
    /*const pdfContentBlob=doc.getAs(MimeType.PDF);
    pdfFolder.createFile(pdfContentBlob).setName(`${row[1]} Employee Name pdf`)
    destinationFolder.removeFile(copy);
    const url = pdfContentBlob.getUrl(); */
    const url = doc.getUrl();
    sheet.getRange(index + 1, 119).setValue(url)
    
  })
}

function sendPDFForm4()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var row =sheet.getActiveCell().getRow();
  var currcol = 119;
  var lastCell = sheet.getRange(row, currcol);
  var lrow = lastCell.getValue();
  var fileid=lrow.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  sendEmailWithAttachment4(row,fileid);
}

function sendEmailWithAttachment4(row,fileid,copy)
{
  const pdfFolder = DriveApp.getFolderById('1jXfTezBVwiL6-WvWdrY7Xoa17OZi59J4');
  const file = DriveApp.getFileById(fileid)
  
  var client = getClientInfo(row);
  
  var template = HtmlService.createTemplateFromFile('email-template4');
  template.client = client;
  var message = template.evaluate().getContent();
  
  
  MailApp.sendEmail({
    to: client.email,
    subject: "Appraisal Letter",
    htmlBody: message,
    attachments: [file.getAs(MimeType.PDF)]
  });
  
  
}



function getClientInfo(row)
{
   var sheet = SpreadsheetApp.getActive().getSheetByName('EMPLOYEE MASTER');
   
   var values = sheet.getRange(row,1,row,119).getValues();
   var rec = values[0];
  
  var client = 
      {
        emp_name: rec[5],
        email: rec[17],
        addr:rec[18],
        m_gross:rec[40]


      };
  client.name = client.emp_name
  client.email = client.email
  client.addr = client.addr
  client.m_gross = client.m_gross
  return client;
}
