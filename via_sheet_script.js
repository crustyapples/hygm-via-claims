// Contains functions for generating certificate hashes, and sending generated certificate files to the respective volunteers
// this script is deployed on an internal google sheet at heyyougotmail.com via AppsScript

function processEmail(sheetObject,n) {
  for (var i=2;i<n+1;i++) {
    
    if (sheetObject.getRange(i,11).getValue() == 'Y' && sheetObject.getRange(i,9).getValue() == 'N') {
      var volunteer = 
      {
        name: sheetObject.getRange(i,7).getValue(),
        email: sheetObject.getRange(i,6).getValue(),
        hours: sheetObject.getRange(i,8).getValue(),
        hash: sheetObject.getRange(i,15).getValue()
      }
      
      var template = HtmlService.createTemplateFromFile('email-template');
      template.volunteer = volunteer;
      var message = template.evaluate().getContent();
      var fileName = volunteer.hash + ".pdf"
      var file = DriveApp.getFilesByName(fileName)
      
      MailApp.sendEmail({
        to: volunteer.email,
        subject: "Hi! Thank you for volunteering with Hey, You Got Mail!",
        htmlBody: message,
        attachments: [file.next().getAs("application/pdf")]
      })

      console.log("Email sent to " + volunteer.name);

    }   
  }
}

function main() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var ProcessSheet = ss.getSheetByName('Process');
  var WebsiteSheet = ss.getSheetByName('Website Form');
  var BotSheet = ss.getSheetByName('Herbert Bot');
  var n = WebsiteSheet.getLastRow();
  var m = BotSheet.getLastRow();
  processEmail(WebsiteSheet,n)
  processEmail(BotSheet,m)
}

function MD5 (input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
  var txtHash = '';
  for (i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0';
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}