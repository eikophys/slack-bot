const SpreadSheetID: string = "1axwEK-8ozcXGm5Xux2A83xYV5pi7aQUv8eHsbIlj8gI"
const slackAuthToken: string = PropertiesService.getScriptProperties().getProperty(
    "slackOAthToken"
)
const slackVerificationToken: string = PropertiesService.getScriptProperties().getProperty(
    "slackVerificationToken"
)

interface slackEventResponse {
    // https://api.slack.com/apis/connections/events-api#the-events-api__receiving-events__callback-field-overview
    token: string
    team_id: string
    api_spp_id: string
    type: string
    authed_users?: []
    authed_teams?: []
    authorizations: Object
    event_id: string
    event_time: number
    event: {
        type: string
        // user: string;
        user: { id: string }
        text: string
        ts: string
        channel: string
        event_ts: string
    }
}
const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
    SpreadSheetID
).getSheets()[0]

function getEmailList(
    sheet: GoogleAppsScript.Spreadsheet.Sheet
): { address: string; name: string }[] {
    let mailData = []
    const data = sheet.getDataRange().getValues()
    data.shift()
    data.forEach((data) => {
        mailData.push({
            address: data[1],
            name: data[0]
        })
    })
    return mailData
}

interface errorObj {
    error: boolean
    message?: string
}

function addEmail(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    name: string,
    address: string,
    id: string
): errorObj {
    let error = {
        error: false,
        message: ""
    }

    if (!!sheet.createTextFinder(id).findAll().length) {
        error.error = true
        error.message += "このメールアドレスは既に登録済みです"
        return error
    }
    if (address == "" || address == null || address == "no_text") {
        error.error = true
        error.message = `メールアドレスを指定してください`
        return error
    }

    const re = /^(([^<>()[\]\\,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
    if (!re.test(address)) {
        error.error = true
        error.message = `メールアドレスの形式が正しくありません: ${address}`
        return error
    }

    sheet.appendRow([name, address, new Date().toLocaleDateString("ja-jp"), id])
    SendEmail(
        address,
        "メールアドレス登録完了",
        `${name}（${address}）がSlack Email Notification for EPCに登録されました。 \n 配信解除を希望する場合には/unsubscribeを実行してください。\n 質問がある場合には#help-slackをご利用ください。\n \n 栄光学園物理研究部 Slack運営チーム`
    )
    return {
        error: error.error
    }
}

function deleteEmail(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    address: string
): { error: boolean; message?: string } {
    let error = false
    let message = ""
    const matchCell: GoogleAppsScript.Spreadsheet.Range[] = sheet
        .createTextFinder(address)
        .findAll()
    if (!matchCell.length) {
        error = true
        message += "このメールアドレスは登録されていません"
    }
    return {
        error: error,
        message: message
    }
}

function SendEmail(recipient: string, subject: string, body: string) {
    MailApp.sendEmail(recipient, subject, body, {
        name: "Slack Mail Notification Bot"
    })
}

function appMentioned(slackData: slackEventResponse) {
    // https://api.slack.com/events/app_mention
    const emailList = getEmailList(sheet)

    const userInfo = getUserInfo(slackData.event.user)
    const postedUser: string = userInfo.ok
        ? userInfo.profile.display_name_normalized != ""
            ? userInfo.profile.display_name_normalized
            : userInfo.profile.real_name_normalized
        : "unknown"
    const text = slackData.event.text.replace(/<@.+>/gm, "")
    emailList.forEach((emailData: { address: string; name: string }) => {
        SendEmail(
            emailData.address,
            `${postedUser}からのメッセージ`,
            `${postedUser}: ${text}`
        )
    })
}

function newMemberJoined(id: string) {
    const slackCreateMessageAPI: string =
        "https://slack.com/api/chat.postMessage"
    const blocks = [
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text: ":wave: 物理部Slackへようこそ！"
            }
        },
        {
            type: "header",
            text: {
                type: "plain_text",
                text: ":frame_with_picture: プロフィール",
                emoji: true
            }
        },
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text:
                    "参加したら始めにして欲しいことがプロフィールの設定です。必ずどこかに自分の実名が含まれるようにしてください。\nプロフィール写真もできれば設定して欲しいです。"
            }
        },
        {
            type: "header",
            text: {
                type: "plain_text",
                text: ":pray: ルール",
                emoji: true
            }
        },
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text:
                    "今のところ、厳しいルールはありません。常識的な使い方を期待しています。\n必ず守って欲しいことは *ダイレクトメッセージ* は絶対に使用しないでください。トラブルの原因となります。\n ルールに変更がある場合には<#C01N73R8B7E>でお知らせします。"
            }
        },
        {
            type: "header",
            text: {
                type: "plain_text",
                text: ":hash: チャンネル",
                emoji: true
            }
        },
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text:
                    "現状、画面の左側に<#C01NDH8RG8Z>や<#C01N5EH24SV>が表示されていると思います。これらは *全ての部員* が参加しているチャンネル（グループ）になります。 \n\n ・<#C01N73R8B7E>：全体連絡用 \n ・<#C01N5EH24SV>：部全体で議論したいことを話してください \n ・<#C01NDH8RG8Z>：質問したいことはこちらから\n ・<#C01MVB70CS2>：雑談用チャンネルです（試験的・任意参加）\n "
            }
        },
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text:
                    "以下は各班のチャンネルです。各班が独自に連絡手段を持っている場合にはそちらを優先してかまいません\n・<#C01NY3RPGP2>：工学班\n ・<#C01N88NTN58>：PC班\n ・<#C01NY3VKXUG>：航空力学班\n ・<#C01MTF9RYMV>：環境化学班\n ・<#C01NY3WNL4Q>：数学班\n・<#C01NY3U41Q8>：地学班\n・<#C01MTFWB8CX>：FLL班\n"
            }
        },
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text:
                    "チャンネル作成は誰でもできます！積極的に作って盛り上げていきましょう！新しいチャンネルを作ったら全員がいるチャンネルで宣伝しましょう。"
            }
        },
        {
            type: "header",
            text: {
                type: "plain_text",
                text: ":robot_face: Bot",
                emoji: true
            }
        },
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text:
                    "このアカウントはBotです。\n試しにチャット入力欄で `/subnscribe` と入力してください。「メール通知を受け取る」と表示されるはずです。\nガラケーなどを利用しており、すぐに通知を見られない場合は、こちらからメール通知を受け取れます。（ `/unsubscribe` で登録解除できます）"
            }
        },
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text:
                    "Botにメンション（＠）を送信すると先ほど登録されたユーザー全員にメールが送信されます。全部員に情報を必ず届けたい場合は＠でこのBotを指定しましょう。"
            }
        }
    ]
    const params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: "post",
        headers: {
            Authorization: `Bearer ${slackAuthToken}`
        },
        payload: {
            channel: id,
            blocks: JSON.stringify(blocks)
        }
    }
    UrlFetchApp.fetch(slackCreateMessageAPI, params).getContentText()
    const returnObj = { blocks: blocks }
    return ContentService.createTextOutput(
        JSON.stringify(returnObj)
    ).setMimeType(ContentService.MimeType.JSON)
}

interface slashCommandResponse {
    token: string
    team_id: string
    team_domain: string
    enterprise_name: string
    enterprise_id: string
    channel_id: string
    channel_name: string
    user_id: string
    user_name: string
    command: string
    text: string
    response_url: string
    trigger_id: string
    api_app_id: string
}

function subscribe(e) {
    const slackData: Readonly<slashCommandResponse> = e.parameter
    const userInfo = getUserInfo(slackData.user_id)
    const name: string =
        userInfo.profile.display_name_normalized != ""
            ? userInfo.profile.display_name_normalized
            : userInfo.profile.real_name_normalized
    let address: string = userInfo.profile.email
    if (address == "" || address == null) {
        // Slackからユーザーのメールを取得できない場合には引数から取得する
        address = slackData.text
    }
    const addEmailStatus = addEmail(sheet, name, address, slackData.user_id)
    const message: string = addEmailStatus.error
        ? addEmailStatus.message
        : `<@${slackData.user_id}> を登録しました`
    const response = {
        type: addEmailStatus.error ? "ephemeral" : "in_channel",
        text: message
    }
    return ContentService.createTextOutput(
        JSON.stringify(response)
    ).setMimeType(ContentService.MimeType.JSON)
}

function unsubscribe(e) {
    const slackData: Readonly<slashCommandResponse> = e.parameter
    const deleteEmailStatus = deleteEmail(sheet, slackData.user_id)
    const message: string = deleteEmailStatus.error
        ? deleteEmailStatus.message
        : "登録解除しました"
    const response = {
        type: deleteEmailStatus.error ? "ephemeral" : "in_channel",
        text: message
    }
    return ContentService.createTextOutput(
        JSON.stringify(response)
    ).setMimeType(ContentService.MimeType.JSON)
}

function getUserInfo(
    id: string
): {
    readonly ok: boolean
    readonly error?: string
    readonly profile: {
        avatar_hash: string
        status_text: string
        status_emoji: string
        status_expiration: string
        real_name: string
        display_name: string
        real_name_normalized: string
        display_name_normalized: string
        email: string
        image_original: string
        team: string
    }
} {
    const slackUsersApi: string = "https://slack.com/api/users.profile.get"
    const params: object = {
        method: "post",
        headers: {
            Authorization: `Bearer ${slackAuthToken}`
        },
        payload: {
            user: id
        }
    }
    const response = JSON.parse(
        UrlFetchApp.fetch(slackUsersApi, params).getContentText()
    )
    return response
}

function doPost(e) {
    if (e.postData.type == "application/json") {
        const slackData: slackEventResponse = JSON.parse(
            e.postData.getDataAsString()
        )
        if (slackData.token == slackVerificationToken)
            if (slackData.event.type == "app_mention") {
                Logger.log("App mentioned")
                appMentioned(slackData)
            }
        if (slackData.event.type == "team_join") {
            const slackCreateMessageAPI: string =
                "https://slack.com/api/chat.postMessage"
            UrlFetchApp.fetch(slackCreateMessageAPI, {
                method: "post",
                headers: { Authorization: `Bearer ${slackAuthToken}` },
                payload: {
                    channel: "C01N5EH24SV",
                    text: `:wave: ようこそ、 <@${slackData.event.user.id}>！`
                }
            }).getContentText()
            newMemberJoined(slackData.event.user.id)
        }
    } else if (
        e.postData.type == "application/x-www-form-urlencoded" &&
        e.parameter.token == slackVerificationToken
    ) {
        if (e.parameter.command == "/subscribe") {
            return subscribe(e)
        } else if (e.parameter.command == "/unsubscribe") {
            return unsubscribe(e)
        } else if (e.parameter.command == "/help") {
            return newMemberJoined(e.parameter.user_id)
        }
    }
}
