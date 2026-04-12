# SAP (Sovereign Asset Protocol)

> 基於 Google Apps Script 的資產追蹤、交易所同步、風險監控與戰略報告系統。

SAP 是一套建構在 Google Apps Script (GAS) 與 Google Sheets 上的個人資產自動化系統。它會整合多個交易所餘額、更新市場價格與匯率、寫入統一資產台帳，並產出槓桿風險、資產配置、再平衡與每日收盤報告。

這個公開 repo 只應保存產品程式碼與必要的部署 workflow。執行期憑證、GAS 專案 ID、試算表 ID、Tunnel URL、本地 bridge 設定與私人施工文件，都必須留在 git 之外。

## 目前能力

- 將多交易所資產同步到統一台帳。
- 交易所同步失敗時保留上次成功結果，避免失敗同步洗掉既有資料。
- 在交易所 API 支援的範圍內追蹤現貨、資金帳戶、理財、質押、借貸抵押品與借貸負債。
- 更新資產價格，支援交易所價格、公開來源與 Google Finance fallback。
- 維護 `價格暫存`，並可自動把資產配置中新出現的 ticker 加入價格目錄。
- 維護 `參數設定` 匯率矩陣，在價格更新與策略報告前更新必要 currency pair。
- 產出資產配置、LTV、質押健康度、狙擊區、再平衡與每日收盤報告。
- 提供 Google Sheets 選單、排程入口、健康檢查、Webhook 與 Discord / Email 通知整合。

## 交易所同步範圍

交易所同步模組集中在 `Sync_*.js`。

| 交易所 | 目前同步範圍 |
| --- | --- |
| Binance | Spot、Funding、Simple Earn flexible / locked、Futures account-level 資訊、Flexible Loan 資訊，實際範圍取決於權限與帳戶模式。 |
| OKX | Spot / Funding 類資產、Earn / Staking、Flexible Loan 抵押品與負債。 |
| Bitget | Spot、Funding、Savings / Earn、Crypto Loan 抵押品與負債。 |
| Bybit | Unified Wallet、衍生品 instrument metadata、Earn / product metadata、借貸欄位；若 GAS IP 被擋，需透過 local bridge。 |
| Pionex | Spot balance。Bot / order 資料屬 optional，取決於 API 權限。 |
| BitoPro | Spot balance，以及 BitoPro 原生資產價格支援。 |

交易所 API 的可用資料會受到 API 權限、帳戶模式、地區限制與交易所文件變動影響。非必要來源失敗時應記錄 warning，不應覆蓋或刪除上次成功台帳。

## 試算表模型

GAS 專案預期綁定一份 Google Sheets 工作簿。下列只列工作表名稱，不包含任何私人試算表 ID。

| 工作表 | 用途 |
| --- | --- |
| `Unified Assets` | 統一資產台帳，由 `Lib_SyncManager.js` 寫入。 |
| `Sync_Status` | 每個交易所最近嘗試、最近成功、狀態與訊息。 |
| `System_Logs` | Apps Script 執行 log。 |
| `價格暫存` | 價格目錄，供策略公式與 `GET_PRICE(ticker)` 使用。 |
| `參數設定` | 匯率矩陣，供 `MY_GOOGLEFINANCE(from, to)` 與部分直接引用使用。 |
| `Balance Sheet` | 策略輸入表，包含資產、負債、Layer、質押群組與配置標籤。 |
| `Key Market Indicators` | 市場指標輸入表，包含部分直接引用。 |

工作表結構由 `Lib_WorkbookContracts.js` 與 `Util_HealthCheck.js` 檢查。任何結構調整後，都應先跑 `runSystemHealthCheck()`。

## 自動化流程

`runAutomationMaster()` 是主要排程入口。

```text
syncCurrencyPairs()
  -> updateAllFxRates()
  -> updateAllPrices()
  -> syncAllAssets_()
  -> runStrategicMonitor()
```

重要行為：

- 匯率先於價格更新，因為部分價格 fallback 會讀取 `參數設定`。
- 價格先於資產同步與策略報告更新。
- 匯率或價格發生 fatal error 時，主流程會停止，避免用明顯不可靠的資料產生行動建議。
- 交易所同步採 guarded commit，避免失敗來源造成局部覆寫。
- Daily Close 使用同一套 master workflow，前置流程失敗時不會繼續快照與日報。

## 匯率系統

`Sync_FxRates.js` 負責維護 `參數設定` 匯率矩陣。

- 必要 currency pair 由 `MY_GOOGLEFINANCE(from, to)` 公式掃描取得，並固定包含 `USD/TWD`。
- 缺少的來源幣別 row 或目標幣別 column 會自動補上。
- 既有矩陣 pair 可 opportunistic 更新，但只有必要 pair 在無可用舊值時會變成 fatal。
- 匯率透過 Google Sheets `GOOGLEFINANCE("CURRENCY:...")` 寫入專屬 `Temp` 區塊取得。
- 讀取公式結果時有小型 retry loop，降低 Google Sheets 暫時顯示 `Loading`、空白或尚未重算造成的誤判失敗。
- `參數設定!B1/A4/B4` 被視為固定 USD/TWD 語意契約，因為工作簿內可能直接引用 `'參數設定'!$B$4`。

`runSystemHealthCheck()` 會檢查必要 pair、無效匯率、過期 timestamp、optional matrix 維護狀態，以及固定 USD/TWD 位置語意。

## 價格系統

`Fetch_CryptoPrice.js`、`Fetch_StockPrice.js`、`Fetch_BitoProPrice.js` 與 `Repo_PriceCache.js` 負責維護 `價格暫存`。

- `GET_PRICE(ticker)` 是 GAS custom function。
- 資產配置中新出現的 ticker 可以自動加入價格目錄。
- 若某列需要手動鎖價，可把 `Updated` 設成未來時間，避免排程覆蓋。
- 股票價格可透過 Google Finance sheet formula 取得。
- BITO 等交易所原生資產可使用交易所專屬 price fetcher。

## 部署方式

本 repo 透過 GitHub Actions 部署到 GAS。

- Workflow 位於 `.github/workflows/deploy.yml`。
- 只有 push 到 `master` 會觸發部署。
- Workflow 會在執行期從 GitHub Secrets 產生 `.clasprc.json` 與 `.clasp.json`。
- 最後以 `clasp push --force` 部署到 Apps Script。

下列 GitHub Secret 名稱會出現在 workflow 中，名稱本身不是秘密：

- `CLASP_CREDENTIALS_JSON_BASE64`
- `GAS_PROJECT_JSON_BASE64`

不要提交產生出的 clasp 設定檔，也不要提交原始 credential JSON。

## 執行期設定

執行期值應放在 Apps Script Script Properties 或 GitHub Secrets，不應放進 source code。

Script Properties 常見類別：

- 管理者與通知設定。
- 各交易所 API key、secret、passphrase。
- local bridge 的共享 tunnel URL 與 relay password。
- staging-only guarded test 設定。

設定名稱可以出現在程式碼中；真正的值不得 commit、不得寫進 README、不得放進公開文件，也不得 hardcode 在 GAS 檔案裡。

## Local Bridge

部分交易所 API 可能拒絕 GAS IP。此專案可以讓部分請求透過本地 bridge 與共享 tunnel 轉送。

安全要求：

- Tunnel URL 視為敏感操作端點。
- Bridge endpoint 必須有 shared password 或等價驗證。
- 不提交 bridge runtime 檔案、本地設定、log、產生出的 URL 或機器路徑。
- local bridge 版本應可獨立回復，讓正式 GAS 可以在新版本失敗時回到已知可用行為。

公開 repo 可以描述 bridge 概念，但不應揭露實際 endpoint。

## 資安邊界

永遠不要提交：

- `.clasp.json`、`.clasprc.json`、OAuth client secret 或產生出的 clasp credential。
- 交易所 API key、API secret、passphrase、webhook secret、password、Discord token。
- GAS Web App URL、Cloudflare Tunnel URL、試算表 ID 或本機路徑。
- `docs/` 底下的私人施工文件。
- local bridge 資料夾、log、程序輸出、臨時 audit、截圖或產生物。

建議流程：

- 本地私有排除項目放在 `.git/info/exclude`。
- 公開 `.gitignore` 維持簡潔，避免把私有環境結構寫得太細。
- 任何曾經 commit 或公開貼出的 credential / endpoint 都應輪替。
- 每次 push 前檢查 `git status`、`git diff --cached`，並做敏感字掃描。

## 常用入口

常用 GAS 函式：

- `runAutomationMaster()`：完整排程流程。
- `runDailyCloseRoutine()`：每日收盤流程。
- `runSystemHealthCheck()`：工作簿、同步、bridge、價格與匯率診斷。
- `updateAllPrices()`：更新價格目錄。
- `updateAllFxRates()`：更新匯率矩陣。
- `syncCurrencyPairs()`：掃描並補齊必要 currency pair。
- `runStrategicMonitor()`：產生策略報告。
- `getBinanceBalance()`、`getOkxBalance()`、`getBitgetBalance()`、`getBybitBalance()`、`getPionexBalance()`、`getBitoProBalance()`：交易所同步入口。

常用試算表 custom function：

- `GET_PRICE(ticker)`
- `MY_GOOGLEFINANCE(fromCurrency, toCurrency)`

## 驗證方式

本地語法檢查：

```powershell
Get-ChildItem -Path . -Filter *.js | ForEach-Object {
  node --check $_.FullName
}
```

push 前 git 檢查：

```powershell
git status --short --branch
git diff --check
git diff --cached --check
```

GAS 部署後 smoke test 建議順序：

1. 執行 `updateAllFxRates()`。
2. 執行 `updateAllPrices()`。
3. 執行 `runSystemHealthCheck()`。
4. 執行 `runAutomationMaster()`。

Guarded staging tests 預設會被擋下，除非 staging Script Properties 明確開啟。測試不應寫入正式工作表。

## 專案結構

```text
/
|-- Core_*.js                 # 主流程、log、策略引擎
|-- Event_*.js                # 試算表選單與 webhook 入口
|-- Fetch_*.js                # 市價抓取
|-- Lib_*.js                  # 同步與 workbook contract 共用邏輯
|-- Repo_*.js                 # 工作表 repository abstraction
|-- Sync_*.js                 # 交易所、currency pair 與 FX sync
|-- Test_*.js                 # guarded staging tests
|-- Util_*.js                 # 設定、憑證、健康檢查、通知
|-- Config.js                 # 非秘密常數
|-- appsscript.json           # GAS manifest
`-- .github/workflows/        # GitHub Actions 部署 workflow
```

## Author

Larry Kang
