const ss = SpreadsheetApp.getActiveSpreadsheet();     //このスプレッドシート
const mainSheet = ss.getSheetByName("main");          //メインシート
const skSheet = ss.getSheetByName("sashikomi");       //差込シート
const logSheet = ss.getSheetByName("log");             //ログシート
const SEND_ERROR_MAIL_ADDRESS = ['メールアドレス']


function pushSendButton() {
  /* スプレッドシートのシートを取得と準備 */

  // const skEndRow = skSheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
  const skEndRow = skSheet.getRange(1,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow(); //シートの宛先メールアドレス列の最終行を取得
  const count = skEndRow - 1;
  // 送信確認
  const select = Browser.msgBox("メールを一括送信します！", count + "件のメールを送信しますが、よろしいですか？", Browser.Buttons.OK_CANCEL);
  if (select == 'ok') {
    try {
      setEmailData(skEndRow);
    }
    catch (e) {
      GmailApp.sendEmail(SEND_ERROR_MAIL_ADDRESS, "【エラー：中断】タイトル", "エラー文");
      Browser.msgBox("エラーが発生したため、処理を中断しました。")
    }
  }
  if (select == 'cancel') {
    Browser.msgBox("送信をキャンセルしました");
  }
}


const setEmailData = (skEndRow) => {
  /* メール基本データの設定 */
  const strFrom = mainSheet.getRange(3, 2).getValue();        // fromメールアドレス
  const strSender = mainSheet.getRange(4, 2).getValue();      // 差出人の名前
  const strSubject = mainSheet.getRange(5, 2).getValue();     // メールタイトル
  const strBody = mainSheet.getRange(6, 2).getValue();        // 本文
  const attachedFileId = mainSheet.getRange(7, 2).getValue(); // 添付ファイルのID
  let attachedFile = undefined
  if (attachedFileId) {
    try {
      attachedFile = DriveApp.getFileById(attachedFileId); // 添付ファイル
    }
    catch (e) {
      Browser.msgBox("ファイルIDが正しくない為、処理を中断します。")
      return false
    }
  }


  sendEmail(skEndRow, strFrom, strSender, strSubject, strBody, attachedFile);
}


function sendEmail(skEndRow, strFrom, strSender, strSubject, strBody, attachedFile) {
  /* 差込文章の作成 */
  logSheet.activate()
  let strSubjectIns = "";   // 差込後のメールタイトル
  let strBodyIns = "";      // 差込後のメール本文
  let executionLog = "";  　// 実行ログ
  let successCount = 0      // 正常処理数
  let errorCount = 0       // エラー処理数

  const date = new Date();

  for (let i = 2; i <= skEndRow; i++) {
    // 差込シートの取得
    const strToEmail = skSheet.getRange(i, 1).getValue(); // 送信先メール
    const strCc = skSheet.getRange(i, 3).getValue();      // CC
    const strBcc = skSheet.getRange(i, 4).getValue();     // BCC
    const sk1 = skSheet.getRange(i, 5).getValue();        // 差込①
    const sk2 = skSheet.getRange(i, 6).getValue();        // 差込②

    // 差込を反映
    strSubjectIns = strSubject.replace(/\$1/g, sk1).replace(/\$2/g, sk2); //タイトル
    strBodyIns = strBody.replace(/\$1/g, sk1).replace(/\$2/g, sk2);       //本文

    /* メール送信 */
    try {
      GmailApp.sendEmail(
        strToEmail,    //toアドレス
        strSubjectIns, //メールタイトル
        strBodyIns,    //本文
        {
          cc: strCc,                //ccアドレス
          bcc: strBcc,              //bccアドレス
          from: strFrom,            //fromアドレス
          name: strSender,          //差出人
          attachments: attachedFile //添付ファイル
        }
      );
      executionLog = "送信完了"
      successCount++
    }
    catch (e) {
      Logger.log('送信エラー')
      executionLog = e
      errorCount++
    }
    finally {
      // ログへの書き込み処理
      // 末尾に追加var
      logSheet.appendRow([Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd hh:mm:ss'), strFrom, strToEmail, executionLog])
    }
  }
  Browser.msgBox(`${successCount}件送信が完了しました。${errorCount}件送信エラーが発生しています。`);
}

