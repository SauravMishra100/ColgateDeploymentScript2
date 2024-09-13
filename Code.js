function doGet(e) {
    var whiteList = SpreadsheetApp.openById('1kfTu1iUqTYpSyhuLmeyxFKyNdkvOovw-By2e72DY8YQ').getSheetByName("Whitelist").getDataRange().getValues();
    console.log(whiteList);
    var currentUser = new Session.getActiveUser().getEmail();
    console.log(whiteList.includes(currentUser));
    var containsUser = false;
    for(const subList of whiteList){
      for(const indexedUser of subList){
        if(indexedUser == currentUser){
          containsUser = true;
        }
      }
    }
    if(containsUser){
      return HtmlService.createHtmlOutput('WebApp')
    }
    return HtmlService.createHtmlOutputFromFile('NotOnWhitelist');
  }
  
  function AddRecord(row) {
      const user = [new Session.getActiveUser().getEmail(), new Date()];
      var id = '1kfTu1iUqTYpSyhuLmeyxFKyNdkvOovw-By2e72DY8YQ';
      var ss= SpreadsheetApp.openById(id);
      var webAppSheet = ss.getSheetByName("Data");
      const fullRow = row.concat(user);
      webAppSheet.appendRow(fullRow);
      sendEmail(fullRow)
    
  }

function sendEmail(fullRow) {
  console.log(fullRow);
    
    var recipient = [fullRow[7],fullRow[6]];
    let recipients = [];
    for(let i = 0; i < recipient.length; i++){
      if(recipient[i] != ''){
        recipients.push(recipient[i]);
      }
    }
    recipient = recipient.filter(element => !'');
    console.log(recipient);
    if(fullRow[10] == "New User Lenovo"){
      var subject = "Your New Employee Laptop is Ready! "+fullRow[0];
      var htmlTemplate = HtmlService.createTemplateFromFile('emailNewLenovo');

      htmlTemplate.model = fullRow[1];
      htmlTemplate.srNumber = fullRow[2];
      htmlTemplate.adUserName = fullRow[3];
      htmlTemplate.initialPassword = fullRow[4];
      htmlTemplate.address = fullRow[5];
      htmlTemplate.oktaEmail = fullRow[7];
      htmlTemplate.trackingInfo = fullRow[8];
      htmlTemplate.enableAccount = fullRow[9];

      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(recipients, subject, "", {htmlBody: htmlForEmail, cc: 'tc_site_support@colpal.com'});
    } else if(fullRow[10] == "Existing User Lenovo"){
      var subject = "Your New Laptop is Ready! "+fullRow[0];
      var htmlTemplate = HtmlService.createTemplateFromFile('existingLenovo');

      htmlTemplate.model = fullRow[1];
      htmlTemplate.srNumber = fullRow[2];
      htmlTemplate.adUserName = fullRow[3];
      htmlTemplate.initialPassword = fullRow[4];
      htmlTemplate.address = fullRow[5];
      htmlTemplate.oktaEmail = fullRow[7];
      htmlTemplate.trackingInfo = fullRow[8];
      htmlTemplate.enableAccount = fullRow[9];

      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(recipients, subject, "", {htmlBody: htmlForEmail, cc: 'tc_site_support@colpal.com'});
    } else if(fullRow[10] == "New User Mac"){
      var subject = "Your New Employee Laptop is Ready! "+fullRow[0];
      var htmlTemplate = HtmlService.createTemplateFromFile('emailNewMac');

      htmlTemplate.model = fullRow[1];
      htmlTemplate.srNumber = fullRow[2];
      htmlTemplate.adUserName = fullRow[3];
      htmlTemplate.initialPassword = fullRow[4];
      htmlTemplate.address = fullRow[5];
      htmlTemplate.oktaEmail = fullRow[7];
      htmlTemplate.trackingInfo = fullRow[8];
      htmlTemplate.enableAccount = fullRow[9];

      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(recipients, subject, "", {htmlBody: htmlForEmail, cc: 'tc_site_support@colpal.com'});
    } else if(fullRow[10] == "Existing User Mac"){
      var subject = "Your New Laptop is Ready! "+fullRow[0];
      var htmlTemplate = HtmlService.createTemplateFromFile('existingMac');

      htmlTemplate.model = fullRow[1];
      htmlTemplate.srNumber = fullRow[2];
      htmlTemplate.adUserName = fullRow[3];
      htmlTemplate.initialPassword = fullRow[4];
      htmlTemplate.address = fullRow[5];
      htmlTemplate.oktaEmail = fullRow[7];
      htmlTemplate.trackingInfo = fullRow[8];
      htmlTemplate.enableAccount = fullRow[9];

      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(recipients, subject, "", {htmlBody: htmlForEmail, cc: 'tc_site_support@colpal.com'});
    }
}