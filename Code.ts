const SpreadSheetID: string = "1axwEK-8ozcXGm5Xux2A83xYV5pi7aQUv8eHsbIlj8gI";
const slackAuthToken: string = PropertiesService.getScriptProperties().getProperty(
  "slackOAthToken"
);
const slackVerificationToken: string = PropertiesService.getScriptProperties().getProperty(
  "slackVerificationToken"
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
    SendEmail(
      address,
      "メールアドレス登録完了",
      `${name}（${address}）がSlack Email Notification for EPCに登録されました。 \n 配信解除を希望する場合には/unsubscribeを実行してください。\n 質問がある場合には#help-slackをご利用ください。\n \n 栄光学園物理研究部 Slack運営チーム`
    );
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

  const userInfo = getUserInfo(slackData.event.user);
  const postedUser: string = userInfo.ok
    ? userInfo.profile.display_name_normalized != ""
      ? userInfo.profile.display_name_normalized
      : userInfo.profile.real_name_normalized
    : "unknown";
  emailList.forEach((emailData: { address: string; name: string }) => {
    SendEmail(
      emailData.address,
      `A message from ${postedUser}`,
      `${postedUser}: ${slackData.event.text}`
    );
  });
}

interface slashCommandResponse {
  token: string;
  team_id: string;
  team_domain: string;
  enterprise_name: string;
  enterprise_id: string;
  channel_id: string;
  channel_name: string;
  user_id: string;
  user_name: string;
  command: string;
  text: string;
  response_url: string;
  trigger_id: string;
  api_app_id: string;
}

function subscribe(e) {
  const slackData: Readonly<slashCommandResponse> = e.parameter;
  console.info(e);
  console.info(e.parameter);
  const userInfo = getUserInfo(slackData.user_id);
  const name: string =
    userInfo.profile.display_name_normalized != ""
      ? userInfo.profile.display_name_normalized
      : userInfo.profile.real_name_normalized;
  const address: string = userInfo.profile.email;
  const addEmailStatus = addEmail(sheet, name, address);
  const message: string = addEmailStatus.error
    ? addEmailStatus.message
    : `<@${slackData.user_id}> を登録しました`;
  const response = {
    type: addEmailStatus.error ? "ephemeral" : "in_channel",
    text: message,
  };
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function doPost(e) {
  if (e.postData.type == "application/json") {
    const slackData: slackEventResponse = JSON.parse(
      e.postData.getDataAsString()
    );
    if (slackData.token == slackVerificationToken)
      if (slackData.event.type == "app_mention") {
        Logger.log("App mentioned");
        appMentioned(slackData);
      }
  } else if (
    e.postData.type == "application/x-www-form-urlencoded" &&
    e.parameter.token == slackVerificationToken
  ) {
    if ((e.parameter.command = "/subscribe")) {
      return subscribe(e);
    }
  }
}

function getUserInfo(
  id: string
): {
  readonly ok: boolean;
  readonly error?: string;
  readonly profile: {
    avatar_hash: string;
    status_text: string;
    status_emoji: string;
    status_expiration: string;
    real_name: string;
    display_name: string;
    real_name_normalized: string;
    display_name_normalized: string;
    email: string;
    image_original: string;
    team: string;
  };
} {
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
