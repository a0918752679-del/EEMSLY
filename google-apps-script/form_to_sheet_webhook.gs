/**
 * Google Apps Script：表單送出後轉拋到 Zeabur 系統。
 *
 * 使用方式：
 * 1. 開啟綁定 Google Sheet 的 Apps Script。
 * 2. 貼上本檔案內容。
 * 3. 修改 ZEABUR_WEBHOOK_URL 與 APP_PASSCODE。
 * 4. 建立觸發器：onFormSubmit，事件來源選「試算表」，事件類型選「表單提交時」。
 *
 * 注意：
 * 本系統主要同步方式是「系統從 Google Sheet 拉回」。
 * 此 Script 是選配，用於需要即時將表單資料送到系統的情境。
 */

const ZEABUR_WEBHOOK_URL = 'https://你的-zeabur-網址/api/sync/pull';
const APP_PASSCODE = '你的內部驗證碼';

function onFormSubmit(e) {
  const payload = {
    trigger: 'google_form_submit',
    submittedAt: new Date().toISOString(),
    namedValues: e.namedValues || {}
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-app-passcode': APP_PASSCODE
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // 表單送出後，請系統從 Google Sheet 拉回最新資料。
  UrlFetchApp.fetch(ZEABUR_WEBHOOK_URL, options);
}
