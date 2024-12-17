const SHEET_NAME = "Registrations"; // シート名
const EMAIL_COLUMN = 3;   // メールアドレスの列 (C列)
const PAYMENT_COLUMN = 6; // 支払い予定の列 (F列)
const STAY_COLUMN = 7;    // 宿泊希望の列 (G列)
const THRESHOLD = 100;    // 参加者通知の閾値

// 回答データの自動処理
function processRegistrations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const seenEmails = {}; // メールアドレスの重複チェック用
  let totalParticipants = 0;
  let stayCount = 0; // 宿泊希望者カウント
  let nonStayCount = 0; // 通い参加者カウント

  for (let i = 1; i < data.length; i++) { // 2行目から処理開始
    const email = data[i][EMAIL_COLUMN - 1];
    const paymentStatus = data[i][PAYMENT_COLUMN - 1];
    const stayStatus = data[i][STAY_COLUMN - 1];

    // メールアドレスの重複チェック
    if (seenEmails[email]) {
      sheet.getRange(i + 1, EMAIL_COLUMN).setBackground("red");
      sheet.getRange(i + 1, EMAIL_COLUMN).setComment("重複したメールアドレスです");
    } else {
      seenEmails[email] = true;
    }

    // 支払い予定が未入力の場合、「未設定」を自動入力
    if (!paymentStatus) {
      sheet.getRange(i + 1, PAYMENT_COLUMN).setValue("未設定");
    }

    // 宿泊希望別のカウント
    if (stayStatus === "宿泊する") {
      stayCount++;
    } else if (stayStatus === "宿泊しない") {
      nonStayCount++;
    }

    totalParticipants++;
  }

  // 集計結果をスプレッドシートに表示
  displaySummary(sheet, totalParticipants, stayCount, nonStayCount);

  // 参加者数が閾値を超えたら管理者に通知
  if (totalParticipants >= THRESHOLD) {
    sendNotification(totalParticipants);
  }
}

// 集計結果をスプレッドシートに表示する関数
function displaySummary(sheet, total, stay, nonStay) {
  sheet.getRange("L1").setValue("参加者合計");
  sheet.getRange("L2").setValue(total);
  
  sheet.getRange("M1").setValue("宿泊希望者数");
  sheet.getRange("M2").setValue(stay);
  
  sheet.getRange("N1").setValue("通い参加者数");
  sheet.getRange("N2").setValue(nonStay);
}

// 参加者数の通知メールを送信
function sendNotification(participantCount) {
  const adminEmail = getAdminEmail(); // 管理者メールアドレスをシートから取得
  const subject = "【テニス合宿】参加者数が上限に達しました";
  const body = `現在、合宿参加者が${participantCount}名に達しました。ご確認ください。`;
  MailApp.sendEmail(adminEmail, subject, body);
}

// 管理者メールアドレスを取得する関数
function getAdminEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  return sheet.getRange("K2").getValue(); // K2にメールアドレスが記入されている
}
