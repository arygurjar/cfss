function sendEmail(e) {

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loan Data").activate();

  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var lrow = sh.getLastRow();

  var emailTemplates = HtmlService.createTemplateFromFile("Email1");

  var loanId = sh.getRange(lrow, 1).getValue();

  var amount = sh.getRange(lrow, 12).getValue();

  var rate = sh.getRange(lrow, 17).getValue();

  var name = sh.getRange(lrow, 3).getValue();

  var contact = sh.getRange(lrow, 5).getValue();

  var email = sh.getRange(lrow, 7).getValue();

  var term = sh.getRange(lrow, 19).getValue();

  var mPayment = sh.getRange(lrow, 21).getValue();
  
  emailTemplates.loanId = loanId;

  emailTemplates.name = name;

  emailTemplates.contact = contact;

  emailTemplates.email = email;

  emailTemplates.amount = amount;

  emailTemplates.rate = rate;

  emailTemplates.term = term;

  emailTemplates.mPayment = mPayment;

  var subject = "CHOUDHARY FINANCIAL SERVICES";

  var html = emailTemplates.evaluate().getContent();

  MailApp.sendEmail({
    to: email,
    name: name,
    subject: subject,
    htmlBody: html
  });

}

// function onEdit(e) {
//   const sh = e.range.getSheet();
//   if (sh.getName() == 'Loan Data') {
//     const sr = 1;
//     const rg = sh.getRange(sr, 1, sh.getLastRow() - sr + 1, sh.getLastColumn());
//     const vs = rg.getValues();
//     //rg.setBorder(false, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
//     const numcolumns = sh.getLastColumn();
//     vs.forEach((r, i) => {
//       if (r[0]) {
//         sh.getRange(i + sr, 1, 1, numcolumns).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
//       } else {
//         sh.getRange(i + sr, 1, 1, numcolumns).setBorder(false, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
//       }
//     });
//   }
// }
