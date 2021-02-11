const mailSubject: string = "Hallo World";
const mailBody: string = "This is a test message from GAS";

const SpreadSheetID: string = "1axwEK-8ozcXGm5Xux2A83xYV5pi7aQUv8eHsbIlj8gI";

function getEmailList(
  SpreadSheetID: string
): { address: string; name: string }[] {
  let mailData = [];
  const sheet = SpreadsheetApp.openById(SpreadSheetID);
  const data = sheet.getDataRange().getValues();
  data.shift();
  data.forEach((data) => {
    mailData.push({
      address: data[1],
      name: data[0],
    });
  });
  return mailData;
}

function SendEmail(recipient: string, subject: string, body: string) {
  GmailApp.sendEmail(recipient, subject, body, {
    name: "Slack Mail Notification Bot",
  });
}

function myFunc() {
  const emailList = getEmailList(SpreadSheetID);
  emailList.forEach((emailData) => {
    SendEmail(emailData.address, mailSubject, mailBody);
    Logger.log(`Sent an email to ${emailData.name}`);
  });
}
