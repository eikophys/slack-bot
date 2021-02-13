const mailSubject: string = "Hallo World";
const mailBody: string = "This is a test message from GAS";

const SpreadSheetID: string = "1axwEK-8ozcXGm5Xux2A83xYV5pi7aQUv8eHsbIlj8gI";

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
  MailApp.sendEmail(recipient, subject, body, {
    name: "Slack Mail Notification Bot",
  });
}

type slackEventResponseReadOnly = Readonly<slackEventResponse>;

function appMentioned(slackData: slackEventResponse) {
  // https://api.slack.com/events/app_mention
  const emailList = getEmailList(SpreadSheetID);

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
  const slackData: slackEventResponse = JSON.parse(
    e.postData.getDataAsString()
  );
  if (slackData.type == "app_mention") {
    appMentioned(slackData);
  }

  const response = {
    challenge: e,
  };
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function getUserInfo(id: string) {
  const slackUsersApi: string = "https://slack.com/api/users.profile.get";
  const params: object = {
    method: "post",
    headers: {
      Authorization: `Bearer ${PropertiesService.getScriptProperties().getProperty(
        "slackOAthToken"
      )}`,
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
