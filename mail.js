function start() {
  var ss = SpreadsheetApp.getActiveSheet()
  var ssName = ss.getName()

  var map = {
    to     : 2,
    from   : 3,
    cc     : 4,
    bcc    : 5,
    subject: 6,
    body   : 7
  }

  var template = get_template(ssName, map)
  if (validateEmail(template.from) != true) {
    Browser.msgBox("From欄が無効です")
    return
  }
  if (template.subject == '') {
    Browser.msgBox("件名が空白です")
    return
  }
  if (template.body == '') {
    Browser.msgBox("本文が空白です")
    return
  }

  var rcpts = get_recipients(template.to)
  if (rcpts == "ERROR") {
    return
  }

  if (rcpts.length == 0) {
    Browser.msgBox('宛先が１件もありません')
    return
  }


  var result = replace_vars(template.subject, rcpts[0])
  if (result.code == 'ERROR') {
    Browser.msgBox('件名に置換できなかった埋め込みがあります 「' + result.errorTag + '」')
    return
  }
  var xsubject = result.text

  var result = replace_vars(template.body, rcpts[0])
  Logger.log(result)

  if (result.code == 'ERROR') {
    Browser.msgBox('本文に置換できなかった埋め込みがあります 「' + result.errorTag + '」')
    return
  }
  var xbody = result.text

  var prev_map = {
    FROM: template.from,
    CC: template.cc,
    BCC: template.bcc
  }

  var res = preview(template, rcpts, prev_map, xsubject, xbody  )
  if (res == "cancel") {
    Browser.msgBox("メール送信を中止しました" )
    return
  }

  var count = 0
  var errorCnt = 0
  for (var i = 0; i < rcpts.length; i++) {
    var result = replace_vars(template.subject, rcpts[i])
    if (result.code == 'ERROR') {
      Browser.msgBox('件名に置換できなかった埋め込みがあります 「' + result.errorTag + '」')
      return
    }

    var subject = result.text

    var result = replace_vars(template.body, rcpts[i])
    if (result.code == 'ERROR') {
      Browser.msgBox('本文に置換できなかった埋め込みがあります 「' + result.errorTag + '」')
      return
    }
    var body = result.text

    try {
      GmailApp.sendEmail(
        rcpts[i].mail_address, subject, body,
          {
            from: template.from,
            cc: template.cc,
            bcc: template.bcc,
          }
      )
      count++
      log_push("成功", ssName, subject, rcpts[i] )
    }
    catch(err) {
      errorCnt++
      Logger.log(err )
      log_push("失敗", ssName, err, rcpts[i] )
    }
  }


  if (errorCnt > 0) {
    Browser.msgBox(errorCnt + "通のメール送信が失敗しています。送信ログを確認してください。" )
  }
  else {
    Browser.msgBox(count + "通のメール送信を完了しました" )
  }

}



function preview(template, rcpts, map, xsubject, xbody) {
  var buf = ""

  for (key in map) {
    buf = buf + key + ": " + map[key] + "\\n"
  }

  buf = buf + "\\n"
  buf = buf + "件名:" + xsubject + "\\n"
  buf = buf + "本文: =============================================================\\n"

  xbody = xbody.replace(/\n/g, "\\n" )
  buf = buf + xbody +  "\\n"
  buf = buf + "=============================================================\\n"
  buf = buf + "\\n"


  buf = buf + "送信先:" +  "\\n"
  for (var i = 0;  i < rcpts.length; i++) {
    buf = buf + (i + 1) + ")"
    for (key in rcpts[i]) {
      buf = buf + " " + rcpts[i][key]
    }
    buf = buf + "\\n"
  }
  buf = buf + "\\n"

  buf = buf + "上記の内容でメールを送信します。よろしいですか？"

  var res =  Browser.msgBox(buf,Browser.Buttons.OK_CANCEL)
  return res
}


function get_recipients(ssName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName(ssName)

  if (sheet == null) {
    Browser.msgBox("宛先一覧が見つかりません")
    return 'ERROR'
  }

  var keys = []
  var col = 1
  var maExists = 0

  while (1) {
    var key = sheet.getRange(1, col).getValue()
    if (key == "") break

    if (key == 'mail_address') {
      maExists = 1
    }

    keys.push(key)

    col = col + 1
  }

  if (maExists != 1) {
    Browser.msgBox('宛先一覧「' + ssName + '」に項目名「mail_address」が見つかりません')
    return 'ERROR'
  }

  var result = []
  var row = 2
  while (1) {
    var bcnt = 0
    var rcp = {}
    var eidx = 0
    for (var col = 1; col <= keys.length; col++) {
      var value = sheet.getRange(row, col ).getValue()
      Logger.log(row +","+ col +","+ value )

      if (value == "") {
        bcnt++
        if (bcnt == 1) {
          eidx = col-1
          Logger.log ("eidx:" + eidx )
        }
      }
      else {
        if (keys[col-1] == 'mail_address') {
          if (validateEmail(value) != true) {
            Browser.msgBox((row) + "行目のメールアドレスが不正です。処理を中止します。" )
            return "ERROR"
          }
        }
      }

      rcp[keys[col-1]] = value
    }

    if (bcnt == keys.length) {
      break
    }

    if (bcnt != 0 && bcnt < keys.length) {
      Browser.msgBox((row) + "行目の" + keys[eidx] + "が空白です。処理を中止します。" )
      return "ERROR"
    }

    result.push(rcp )

    row = row + 1
  }

  return result
}


function get_template(ssName, map) {
  var col = 2
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName(ssName)

  var template = {}
  for (key in map) {
    var buf = sheet.getRange(map[key], col).getValue()
    template[key] = buf
  }
  return template
}


function replace_vars(temp, rcp) {
  var xbody = temp
  var buf = ""

  for (key in rcp) {
       xbody = xbody.replace(new RegExp('@{' + key + '}', 'g'), rcp[key])
  }

  var errorTag = xbody.match(new RegExp('@\{.+\}', 'g'))
  if (errorTag !== null) {
    return { 'code': 'ERROR', 'errorTag': errorTag }
  }
  return { 'code': 'OK', 'text': xbody }
}


function log_push(status, kind, subject, rcp) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("送信ログ")

  var buf = ""
  for (key in rcp) {
    buf = buf + rcp[key] + ","
  }
  sheet.appendRow([Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"), status, kind, subject, buf ])
}


function validateEmail(email){
  var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/
  return emailPattern.test(email)
}

