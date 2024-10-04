const jsonString = HtmlService.createHtmlOutputFromFile("config.json.html").getContent();
const config = JSON.parse(jsonString);
const spreadsheetId = config.spreadSheetID;
const macPDFID = config.macDeploymentPDF;
const lenovoPDFID = config.winDeploymentPDF;

//Initial function of the WebApp
function doGet(e) {
    // Reads the emails from the Whitelist sheet in the spreadSheetID spreadsheet and verifies if the current user accessing the link is on the list or not.
    // If the user is not on the list, they will get a message saying "Access not granted"
    var whiteList = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Whitelist").getDataRange().getValues();
    var currentUser = new Session.getActiveUser().getEmail();
    var containsUser = false;
    for(const subList of whiteList){
      for(const indexedUser of subList){
        if(indexedUser == currentUser){
          containsUser = true;
        }
      }
    }
    if(containsUser){
      return HtmlService.createHtmlOutputFromFile('WebApp');
    }
    return HtmlService.createHtmlOutputFromFile('NotOnWhitelist');
  }

  function getData(){
    var values = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Data").getDataRange().getValues()
    for(var i = 1; i < values.length; i++){
      values[i][12] = values[i][12].toLocaleString("en-US");
    }
    console.log(values);
    return values;
  }


//Adds the details to the Spreadsheet and also calls the email function. Automatically fills out the email of the user using the script and the time submitted.
function AddRecord(row) {
    const user = [new Session.getActiveUser().getEmail(), new Date()];
    var ss= SpreadsheetApp.openById(spreadsheetId);
    var webAppSheet = ss.getSheetByName("Data");
    const fullRow = row.concat(user);
    webAppSheet.appendRow(fullRow);
    sendEmail(fullRow)
}

//Sends the email with the details extracted from the form.
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
    //Email template is chosen based on the deployment chosen by the user.
    //Email templates are in the form of HTML files such as emailNewLenovo.html and existingMac.html
    if(fullRow[10] == "New User Lenovo"){
      var subject = "Your New Employee Laptop is Ready! "+fullRow[0];
      var htmlTemplate = HtmlService.createTemplateFromFile('emailNewLenovo');
      var file = DriveApp.getFileById(lenovoPDFID);


      htmlTemplate.model = fullRow[1];
      htmlTemplate.srNumber = fullRow[2];
      htmlTemplate.adUserName = fullRow[3];
      htmlTemplate.initialPassword = fullRow[4];
      htmlTemplate.address = fullRow[5];
      htmlTemplate.oktaEmail = fullRow[7];
      htmlTemplate.trackingInfo = fullRow[8];
      htmlTemplate.enableAccount = fullRow[9];

      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(recipients, subject, "", {htmlBody: htmlForEmail, cc: 'tc_site_support@colpal.com', attachments: [file.getAs(MimeType.PDF)]});
    } else if(fullRow[10] == "Existing User Lenovo"){
      var subject = "Your New Laptop is Ready! "+fullRow[0];
      var htmlTemplate = HtmlService.createTemplateFromFile('existingLenovo');
      var file = DriveApp.getFileById(lenovoPDFID);

      htmlTemplate.model = fullRow[1];
      htmlTemplate.srNumber = fullRow[2];
      htmlTemplate.adUserName = fullRow[3];
      htmlTemplate.initialPassword = fullRow[4];
      htmlTemplate.address = fullRow[5];
      htmlTemplate.oktaEmail = fullRow[7];
      htmlTemplate.trackingInfo = fullRow[8];
      htmlTemplate.enableAccount = fullRow[9];

      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(recipients, subject, "", {htmlBody: htmlForEmail, cc: 'tc_site_support@colpal.com', attachments: [file.getAs(MimeType.PDF)]});
    } else if(fullRow[10] == "New User Mac"){
      var subject = "Your New Employee Laptop is Ready! "+fullRow[0];
      var htmlTemplate = HtmlService.createTemplateFromFile('emailNewMac');
      var file = DriveApp.getFileById(macPDFID);

      htmlTemplate.model = fullRow[1];
      htmlTemplate.srNumber = fullRow[2];
      htmlTemplate.adUserName = fullRow[3];
      htmlTemplate.initialPassword = fullRow[4];
      htmlTemplate.address = fullRow[5];
      htmlTemplate.oktaEmail = fullRow[7];
      htmlTemplate.trackingInfo = fullRow[8];
      htmlTemplate.enableAccount = fullRow[9];

      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(recipients, subject, "", {htmlBody: htmlForEmail, cc: 'tc_site_support@colpal.com', attachments: [file.getAs(MimeType.PDF)]});
    } else if(fullRow[10] == "Existing User Mac"){
      var subject = "Your New Laptop is Ready! "+fullRow[0];
      var htmlTemplate = HtmlService.createTemplateFromFile('existingMac');
      var file = DriveApp.getFileById(macPDFID);

      htmlTemplate.model = fullRow[1];
      htmlTemplate.srNumber = fullRow[2];
      htmlTemplate.adUserName = fullRow[3];
      htmlTemplate.initialPassword = fullRow[4];
      htmlTemplate.address = fullRow[5];
      htmlTemplate.oktaEmail = fullRow[7];
      htmlTemplate.trackingInfo = fullRow[8];
      htmlTemplate.enableAccount = fullRow[9];

      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(recipients, subject, "", {htmlBody: htmlForEmail, cc: 'tc_site_support@colpal.com', attachments: [file.getAs(MimeType.PDF)]});
    }
}
