const SpreadSheetID: string = "1axwEK-8ozcXGm5Xux2A83xYV5pi7aQUv8eHsbIlj8gI";
const slackAuthToken: string = PropertiesService.getScriptProperties().getProperty(
  "slackOAthToken"
);

interface slackEventResponse {
  // https://api.slack.com/apis/connections/events-api#the-events-api__receiving-events__callback-field-overview
  token: string;
  team_id: string;
  api_spp_id: string;
  type: string;
  authed_users?: [];
  authed_teams?: [];
  authorizations: Object;
  event_id: string;
  event_time: number;
  event: {
    type: "app_mention";
    user: string;
    text: string;
    ts: string;
    channel: string;
    event_ts: string;
  };
}
const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
  SpreadSheetID
).getSheets()[0];

function getEmailList(
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): { address: string; name: string }[] {
  let mailData = [];
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

function addEmail(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  name: string,
  address: string
): { error: boolean; message?: string } {
  let error = false;
  let message = "";
  if (!!sheet.createTextFinder(address).findAll().length) {
    error = true;
    message += "このメールアドレスは既に登録済みです。";
  }
  if (!error) {
    sheet.appendRow([name, address, new Date().toLocaleDateString("ja-jp")]);
  }
  return {
    error: error,
    message: message,
  };
}

function SendEmail(recipient: string, subject: string, body: string) {
  MailApp.sendEmail(recipient, subject, body, {
    name: "Slack Mail Notification Bot",
  });
}

function appMentioned(slackData: slackEventResponse) {
  // https://api.slack.com/events/app_mention
  const emailList = getEmailList(sheet);

  const postedUser: string = getUserInfo(slackData.event.user).ok
    ? getUserInfo(slackData.event.user).profile.display_name_normalized != ""
      ? getUserInfo(slackData.event.user).profile.display_name_normalized
      : getUserInfo(slackData.event.user).profile.real_name_normalized
    : "unknown";
  emailList.forEach((emailData: { address: string; name: string }) => {
    SendEmail(
      emailData.address,
      `A message from ${postedUser}`,
      `${postedUser}: ${slackData.event.text}`
    );
  });
}

function doPost(e) {
  if (e.postData.type == "application/json") {
    const slackData: slackEventResponse = JSON.parse(
      e.postData.getDataAsString()
    );
    if (slackData.token == slackAuthToken)
      if (slackData.event.type == "app_mention") {
        Logger.log("App mentioned");
        appMentioned(slackData);
      }
  }
}

function getUserInfo(id: string) {
  const slackUsersApi: string = "https://slack.com/api/users.profile.get";
  const params: object = {
    method: "post",
    headers: {
      Authorization: `Bearer ${slackAuthToken}`,
    },
    payload: {
      user: id,
    },
  };
  const response = JSON.parse(
    UrlFetchApp.fetch(slackUsersApi, params).getContentText()
  );
  return response;
}
