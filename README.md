# 新北市環保局｜聲音照相超標車輛裁罰案件審核系統

角色設定：新北市環保局委外顧問公司「聯韻聲學環保顧問股份有限公司」。

本系統用於局內長官審核聲音照相超標車輛案件，並銜接後續「逕行告發」與「通知到檢」行政流程追蹤。系統以科技簡約風格呈現案件狀態，長官可透過勾選核可改變案件狀態，承辦端可追蹤公文文號、發文日期、裁處字號、送達日期、繳款、通檢期限與結案情形。

## 已內建功能

- 告發、通檢兩類案件分流管理。
- 固定格式 Excel 匯入與匯出。
- 局內長官勾選核可後，案件自動轉為「行政處理中」。
- 退回修正與退回原因紀錄。
- 行政流程追蹤欄位：
  - 公文文號
  - 發文日期
  - 裁處字號
  - 裁罰金額
  -   - 送達日期
  - 繳款狀態
  - 通檢期限
  - 到檢狀態
  - 結案日期
  - 行政備註
- 每筆案件保留更新紀錄。
- Google Sheet 雙向同步：
  - 從 Google Sheet 拉回資料
  - 將系統資料推送到 Google Sheet
  - 測試 Google Sheet 連線成功按鈕
- Zeabur Docker 部署設定。

## 目錄

```text
.
├── server.js
├── package.json
├── Dockerfile
├── zeabur.json
├── columns.json
├── .env.example
├── public/
│   ├── index.html
│   ├── styles.css
│   └── app.js
├── data/
│   └── cases.json
├── docs/
│   └── ZEABUR部署與操作手冊.md
└── google-apps-script/
    └── form_to_sheet_webhook.gs
```

## 本機測試

```bash
npm install
cp .env.example .env
npm start
```

開啟：

```text
http://localhost:3000
```

## Zeabur 快速部署

1. 將整包上傳到 GitHub。
2. Zeabur 新增專案，選擇該 GitHub repository。
3. Deployment method 選 Dockerfile。
4. 設定環境變數。
5. 建議掛載 Volume：
   - Mount path：`/data`
   - 環境變數：`DATA_DIR=/data`
6. 部署完成後，進入外網網址。

## 必填環境變數

最小可用設定：

```env
PORT=8080
DATA_DIR=/data
APP_PASSCODE=自行設定
ADMIN_DELETE_PASSCODE=69677323
```

Google Sheet 串接設定：

```env
GOOGLE_SHEET_ID=你的GoogleSheetID
GOOGLE_SERVICE_ACCOUNT_JSON={"type":"service_account",...}
```

若 Zeabur 不方便貼多行 JSON，可改貼 Base64：

```env
GOOGLE_SERVICE_ACCOUNT_JSON_BASE64=base64後的service_account_json
```

## Google Sheet 分頁

建議建立以下分頁：

- 告發
- 通檢
- 案件追蹤紀錄

系統啟動同步時，也會嘗試自動建立缺少的分頁。

## 固定格式

系統固定以你提供的 Excel 為基礎：

- 分頁「告發」欄位：承辦複核、是否累犯、違規日期、前次告發、日數、1年內再次違反、車號、車種、量測日期、稽查時間、量測位置地點、便於複製(行政區)、便於複製(路段)、背景修正後分貝、管制標準、金額、稽查單號、告發單號、簽日期、簽文號、廠牌、出廠年月、車主姓名、證號、生日、郵遞區號、戶籍地。
- 分頁「通檢」欄位：備註、車號、車種、量測日期、稽查時間、量測位置地點、背景修正後分貝、管制標準、廠牌、出廠年月、車主姓名、證號、生日、郵遞區號、戶籍地。
- 系統另外附加審核與行政追蹤欄位。

## 操作流程

1. 顧問公司匯入告發或通檢案件。
2. 系統進入「待長官審核」狀態。
3. 長官勾選「核可」。
4. 系統自動轉為「行政處理中」。
5. 承辦人員進入「行政追蹤」填寫公文文號、發文日期、裁處字號、送達日期與繳款等。
6. 完成後填入結案日期或設定已結案狀態。
7. 可匯出 Excel 或同步回 Google Sheet。

## 安全提醒

此封包已避免放入原始 Excel 的個資樣本；`data/cases.json` 僅放匿名示範資料。正式部署時請：

- 務必設定 `APP_PASSCODE`。
- Zeabur 專案權限限縮給內部人員。
- Google Sheet 僅共用給必要帳號與 Service Account。
- 不要把 Service Account JSON 放到 GitHub。


## V1.2 功能修正

- 車號重複超標案件會自動標示「加重」，並於匯出欄位「重複超標註記」帶出。
- 從 Google Sheet 同步時，Google Sheet 同步規則：告發案件須填「稽查單號」或「告發單號」（有「案件編號」也可）；通檢案件須填「案件編號」。缺少必要識別欄位者會略過，不會自動產生案件編號。
- 案件列表新增「修改」功能，可調整案件編號、狀態、來源欄位與加重註記。
- 案件列表新增「刪除」功能；刪除需要管理權限密碼，預設為 `69677323`，也可用 `ADMIN_DELETE_PASSCODE` 環境變數調整。

## V1.4 Google Sheet 欄位修正

- 告發案件 Google Sheet 可使用「稽查單號」與「告發單號」作為案件識別，欄名若有空白或換行會自動修正。
- 告發案件仍可保留「案件編號」欄位；若未填，系統不自動亂編號，會改以稽查單號或告發單號作為案件主鍵。
- 通檢案件 Google Sheet 必須填「案件編號」；未填者略過，不自動產生案件編號。
- 相容舊欄名「稽查編號」，同步時會自動對應為「稽查單號」。
