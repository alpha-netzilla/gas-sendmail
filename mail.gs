function click() {
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
    Browser.msgBox("FROM mail address is invalid: " + template.from)
    return
  }
  if (template.subject == "") {
    Browser.msgBox("SUBJECT is NULL")
    return
  }
  if (template.body == "") {
    Browser.msgBox("BODY is NULL")
    return
  }

  var rcpts = get_rcpts(template.to)
  if (rcpts == "ERROR") {
    return
  }
  if (rcpts.length == 0) {
    Browser.msgBox("No RCPT addresses")
    return
  }

  var result = replace_vars(template.subject, rcpts[0])
  if (result.code == "ERROR") {
    Browser.msgBox("Failed to replace the value: " + result.errorTag)
    return
  }
  var subject = result.text

  var result = replace_vars(template.body, rcpts[0])
  Logger.log(result)

  if (result.code == "ERROR") {
    Browser.msgBox("Failed to replace the value: " + result.errorTag)
    return
  }
  var body = result.text

  var prev_map = {
    FROM: template.from,
    CC: template.cc,
    BCC: template.bcc
  }

  var callback = preview(template, rcpts, prev_map, subject, body)
  if (callback == "cancel") {
    Browser.msgBox("Canceled to send mail" )
    return
  }

  var count = 0
  var errorCnt = 0
  for (var i = 0; i < rcpts.length; i++) {
    var result = replace_vars(template.subject, rcpts[i])
    if (result.code == "ERROR") {
      Browser.msgBox("Failed to replace the value: " + result.errorTag)
      return
    }
    var subject = result.text

    var result = replace_vars(template.body, rcpts[i])
    if (result.code == "ERROR") {
      Browser.msgBox("Failed to replace the value: " + result.errorTag)
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
      log_push("SUCCESS", subject, rcpts[i].mail_address)
    }
    catch(err) {
      errorCnt++
      Logger.log(err)
      log_push("FAILURE", err, rcpts[i].mail_address)
    }
  }


  if (errorCnt > 0) {
    Browser.msgBox("Faild to send " + errorCnt + " mail" )
  }
  else {
    Browser.msgBox("Succeeded to send " + count + " mail" )
  }

}


function preview(template, rcpts, map, subject, body) {
  var buf = ""

  for (key in map) {
    buf = buf + key + ": " + map[key] + "\\n"
  }

  buf = buf + "\\n"
  buf = buf + "SUBJECT: " + subject + "\\n"
  buf = buf + "BODY:\\n"
  buf = buf + "=============================================================\\n"

  body = body.replace(/\n/g, "\\n" )
  buf = buf + body +  "\\n"
  buf = buf + "=============================================================\\n"
  buf = buf + "\\n"

  buf = buf + "RCPTS:" +  "\\n"
  for (var i = 0;  i < rcpts.length; i++) {
    buf = buf + (i + 1) + ")"
    for (key in rcpts[i]) {
      buf = buf + ", " + rcpts[i][key]
    }
    buf = buf + "\\n"
  }
  buf = buf + "\\n"

  var callback = Browser.msgBox(buf, Browser.Buttons.OK_CANCEL)
  return callback
}


function get_rcpts(ssName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName(ssName)

  if (sheet == null) {
    Browser.msgBox(ssName + " was not found.")
    return "ERROR"
  }

  var keys = []
  var row = 1
  var col = 1

  while (1) {
    var key = sheet.getRange(row, col).getValue().toLowerCase()
    if (key == "") break;
    keys.push(key)
    col = col + 1
  }

  var rcpts = []
  var row = 2

  while (1) {
    var rcpt = {}
    var eidx = 0

    for (var col = 1; col <= keys.length; col++) {
      var value = sheet.getRange(row, col).getValue()
      Logger.log(row + "," + col + "," + value )

      rcpt[keys[col-1]] = value
    }

    if (value == "") return rcpts;

    rcpts.push(rcpt)
    row = row + 1

  }
  return rcpts
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


function replace_vars(text, rcpt) {
  var buf = ""

  for (key in rcpt) {
    text = text.replace(new RegExp("@{" + key + "}", "g"), rcpt[key])
  }

  var errorTag = text.match(new RegExp("@\{.+\}", "g"))
  if (errorTag != null) {
    return { "code": "ERROR", "errorTag": errorTag }
  }
  return { "code": "OK", "text": text }
}


function log_push(status, subject, rcpt) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("log")
  var date = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss")
  sheet.appendRow([date, status, subject, rcpt])
}


function validateEmail(email){
  var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/
  return emailPattern.test(email)
}

