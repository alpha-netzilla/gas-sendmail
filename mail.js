
function start() {
  var sname = SpreadsheetApp.getActiveSheet().getName();

  var dict = get_dict( sname );
  Logger.log( dict );

  
  var temp_map = {
    to: 2,
    from: 3,
    cc: 4,
    bcc: 5,
    subject: 6,
    body: 7
  };
  
  var template = get_template( sname, temp_map );
  if ( validateEmail( template.from ) != true ) {
    Browser.msgBox("From欄が無効です");
    return;
  }
  
  if ( template.subject == '' ) {
    Browser.msgBox("件名が空白です");
    return;
  }
  if ( template.body == '' ) {
    Browser.msgBox("本文が空白です");
    return;
  }
  
  var rcps = get_recipients(template.to);
   if ( rcps == "ERROR" ) {
   return;
   }
   Logger.log( rcps );
   if ( rcps.length == 0 ) {
    Browser.msgBox('宛先が１件もありません');
    return;
   }
   
  
  var result = expand_tags( template.subject, dict, rcps[0] );
  Logger.log( result );
  if ( result.code == 'ERROR' ) {
    Browser.msgBox('件名に置換できなかった埋め込みがあります 「' + result.errtag + '」');
    return;
  }
  var xsubject = result.text;

  var result = expand_tags(template.body, dict, rcps[0] );
  Logger.log( result );
  if ( result.code == 'ERROR' ) {
    Browser.msgBox('本文に置換できなかった埋め込みがあります 「' + result.errtag + '」')
    return;
  }
  var xbody = result.text;

  var prev_map = {
  FROM: template.from,
  CC: template.cc,
  BCC: template.bcc
};
  var res = preview( template, dict, rcps, prev_map, xsubject, xbody  );
  if ( res == "cancel" ) {
    Browser.msgBox( "メール送信を中止しました" );
    return;
  }

  var count = 0;
  var errcnt = 0;
  for ( var i = 0; i < rcps.length; i++ ) {
   
    var result = expand_tags( template.subject, dict, rcps[i] );
    if ( result.code == 'ERROR' ) {
      Browser.msgBox('件名に置換できなかった埋め込みがあります 「' + result.errtag + '」');
      return;
    }  

    var subject = result.text;
  
    var result = expand_tags(template.body, dict, rcps[i] );
    if ( result.code == 'ERROR' ) {
      Browser.msgBox('本文に置換できなかった埋め込みがあります 「' + result.errtag + '」');
      return;
    }  
    var body = result.text;
  
    try {
      GmailApp.sendEmail(
        rcps[i].mail_address, subject, body,
          {
            from: template.from,
            cc: template.cc,
            bcc: template.bcc,
          }
      );
      count++;
      log_push( "成功", sname, subject, rcps[i] );
    }
    catch( err ) {
      errcnt++;
      Logger.log( err );
      log_push( "失敗", sname, err, rcps[i] );
    }
  }


  if ( errcnt > 0 ) {
    Browser.msgBox( errcnt + "通のメール送信が失敗しています。送信ログを確認してください。" );
  }
  else {
    Browser.msgBox( count + "通のメール送信を完了しました" );
  }
  
}



function preview( template, dict, rcps, map, xsubject, xbody ) {

  var buf = "";

  for ( key in map ) {
    buf = buf + key + ": " + map[key] + "\\n";
  }

  buf = buf + "\\n"; 
  buf = buf + "件名:" + xsubject + "\\n";     
  buf = buf + "本文: =============================================================\\n";

  xbody = xbody.replace( /\n/g, "\\n" );
  buf = buf + xbody +  "\\n";
  buf = buf + "=============================================================\\n";
  buf = buf + "\\n"; 


  buf = buf + "送信先:" +  "\\n";
  for ( var i = 0;  i < rcps.length; i++ ) {
    buf = buf + (i + 1) + ")"; 
    for ( key in rcps[i] ) {
      buf = buf + " " + rcps[i][key];
    }
    buf = buf + "\\n";
  }
  buf = buf + "\\n";

  buf = buf + "上記の内容でメールを送信します。よろしいですか？";
   
  var res =  Browser.msgBox( buf,Browser.Buttons.OK_CANCEL);
  return res;
}


function get_dict( p_sname ) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(p_sname);
  
  var result = {};
  var row = 10;
  while ( 1 ) {

    var key = sheet.getRange( row, 1 ).getValue();
    var value = sheet.getRange( row, 2 ).getValue();
    
    // 空白の場合は終了
    if ( key == "" ) break;
 
    result[key] = value;
    row = row + 1;
    
  }
                  
  return result;
  
}



function get_recipients(p_sname) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(p_sname);
  if ( sheet == null ) {
     Browser.msgBox("宛先一覧が見つかりません")
     return 'ERROR'
  }

  var keys = [];
  var col = 1;
  var maExists = 0;
  while ( 1 ) {
  
    var key = sheet.getRange( 1, col ).getValue();
    if ( key == "" ) break;
    
    if ( key == 'mail_address' ) {
      maExists = 1;
    }
    
    keys.push( key );
    
    col = col + 1;
  }
  
  if ( maExists != 1 ) {
    Browser.msgBox('宛先一覧「' + p_sname + '」に項目名「mail_address」が見つかりません');
    return 'ERROR';
  }
  
  
  //
  var result = [];
  var row = 2
  while ( 1 ) {
    
    var bcnt = 0;
    var rcp = {};
    var eidx = 0;
    for ( var col = 1; col <= keys.length; col++ ) {
      var value = sheet.getRange( row, col ).getValue();
      Logger.log( row +","+ col +","+ value );

      if ( value == "" ) {
        bcnt++;
        if ( bcnt == 1 ) {
          eidx = col-1;
          Logger.log ( "eidx:" + eidx ); 
        }
      }
      else {
        if ( keys[col-1] == 'mail_address' ) {
          if ( validateEmail(value) != true ) {
            Browser.msgBox( (row) + "行目のメールアドレスが不正です。処理を中止します。" );
            return "ERROR";
          }
        }
      }
      
      rcp[keys[col-1]] = value;
    }

    if ( bcnt == keys.length ) {
      break;
    }
    
    if ( bcnt != 0 && bcnt < keys.length ) {
         
         Browser.msgBox( (row) + "行目の" + keys[eidx] + "が空白です。処理を中止します。" );
         return "ERROR";
    }
    
    result.push( rcp );
    
    row = row + 1;

}

  return result;
}


function get_template( p_sname, p_map ) {
  
  var C_COL = 2;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(p_sname);

  var result = {};
  for ( key in p_map ) {
    var buf = sheet.getRange( p_map[key], C_COL ).getValue();
    result[key] = buf;
  }

  return result;
  
}
// -------------------------------------------------------------------------------



// -------------------------------------------------------------------------------
// 埋め込み情報の展開
//
   function expand_tags( p_temp, p_dict, p_rcp ) {
   
     var xbody = p_temp;
     var buf = "";
     for ( key in p_dict ) {
       //Logger.log( 'key:' + key + ' val:' + p_dict[key] );
       buf = xbody.replace( new RegExp( '#{' + key + '}', 'g' ), p_dict[key] );
       //Logger.log( "<<" + buf + ">>" );
      xbody = buf;
     }
     var errtag = xbody.match( new RegExp( '#\{.+\}', 'g' ) );
     if ( errtag !== null ) {  
       return { 'code': 'ERROR', 'errtag': errtag }; 
     }
     
     for ( key in p_rcp ) {
       xbody = xbody.replace( new RegExp( '@{' + key + '}', 'g' ), p_rcp[key] );
       //Logger.log( "<<" + xbody + ">>" );
     }
     var errtag = xbody.match( new RegExp( '@\{.+\}', 'g' ) );
     if ( errtag !== null ) {
       return { 'code': 'ERROR', 'errtag': errtag }; 
     }
     
     return { 'code': 'OK', 'text': xbody };
}


function log_push(status, kind, subject, rcp ) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("送信ログ");
  
  var buf = "";
  for ( key in rcp ) {
    buf = buf + rcp[key] + ",";
  }
  sheet.appendRow([Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"), status, kind, subject, buf ]);
}


function validateEmail(email){
  var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;
  return emailPattern.test(email)
}
