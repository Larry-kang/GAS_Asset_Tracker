# BTC Dual-Factor Regime 開發計畫書

日期：2026-05-21

範圍：`Core_StrategicEngine.js` 的 BTC 週期判定、Crypto LTV 建議、Layer 動態配置

狀態：施工中（V1 helper / regime 接線已完成，待 GAS 實跑驗證）

---

## 1. 背景

目前 GAS 的 BTC 策略核心仍高度依賴 `BTC_MM`：

- `buildContext()` 會讀取 `Key Market Indicators` 的 `BTC_MM`
- `BTC_MM` 會影響 L1/L2/L3 的動態目標配置
- 報告中的 BTC 週期定位與 Crypto 目標 LTV 也主要由 `BTC_MM` 決定

這個設計簡潔，但有一個週期性風險：當熊市時間拉長時，`BTC_200DMA` 會下修，導致 `BTC_MM = BTC_Price / BTC_200DMA` 被動回升。此時 BTC 價格可能仍距離高點很遠，但系統會因為 MM 回升而提早把市場視為中性。

專案其實已經有 `SAP_Base_ATH`，而且 Martingale Sniper 已經用它計算 BTC 距離定錨高點的跌幅。這次要做的不是新增一堆 sheet 欄位，而是把既有資訊正式納入 BTC regime 判定。

---

## 2. 設計目標

第一階段目標是把 BTC 判定從單因子模型升級為：

```text
BTC_MM + ATH Drawdown + Crypto LTV + Cash Buffer
```

核心成果：

- 在程式內由 `Current_BTC_Price` 與 `SAP_Base_ATH` 推導 `btcDrawdownFromATH`
- 建立集中式 `getBtcRegime_()` helper
- 統一驅動：
  - BTC 週期定位文字
  - Crypto 目標 LTV 區間
  - L1/L2/L3 動態配置目標
  - 高 LTV guardrail 訊息
- 保留 `Alloc_L1_Target / Alloc_L2_Target / Alloc_L3_Target / Alloc_L4_Target` 的人工 override 優先權
- 不改 live Finance sheet 結構

---

## 3. 非目標

本階段不做：

- 不新增 `Key Market Indicators` keys
- 不新增 `BTC_Drawdown_From_ATH / BTC_Regime / BTC_LTV_Target / BTC_Restock_Mode` 等輸出型欄位
- 不新增自動下單或交易執行
- 不改 `BTC_MARTINGALE` 預算與 level 設定
- 不重構整個策略引擎
- 不改 GAS workbook contract 的欄位結構

如果後續真的需要把 regime 寫回 sheet，應作為第二階段資料輸出設計，而不是第一階段策略輸入。

---

## 4. 現況整理

### 4.1 現有輸入

`Repo_KeyMarketIndicatorsView.js` 已讀取：

- `SAP_Base_ATH`
- `Current_BTC_Price`
- `BTC_MM`
- `Alloc_L1_Target`
- `Alloc_L2_Target`
- `Alloc_L3_Target`
- `Alloc_L4_Target`

### 4.2 現有 drawdown 使用點

`BTC Martingale Sniper` 已用：

```javascript
const currentDrop = (context.market.btcPrice - context.market.sapBaseATH) / context.market.sapBaseATH;
```

但這個 drawdown 目前只影響狙擊訊號，沒有進入一般 BTC 週期定位、LTV 目標或配置目標。

### 4.3 現有 MM 使用點

目前 MM 主要影響：

- `assetGroups` 的動態 target
- 市場情報中的 `週期定位`
- `目標 LTV (Crypto) (建議)`

問題是這三段邏輯各自寫一份 threshold，未來容易漂移。

---

## 5. V1 規格

### 5.1 派生欄位

在 `buildContext()` 內建立：

```javascript
market.btcDrawdownFromATH = null;

if (market.btcPrice > 0 && market.sapBaseATH > 0) {
  market.btcDrawdownFromATH = (market.btcPrice - market.sapBaseATH) / market.sapBaseATH;
}
```

定義：

- `-0.30` 代表距離定錨高點下跌 30%
- 若 `SAP_Base_ATH` 缺失或無效，回傳 `null`
- `SAP_Base_ATH` 在 V1 視為「策略定錨高點」，不強制要求它一定是歷史最高價

### 5.2 Regime helper

新增內部 helper：

```javascript
function getBtcRegime_(btcMM, btcDrawdown, cryptoLTV, survivalRunway) {
  // returns normalized regime object
}
```

回傳結構：

```javascript
{
  regime: "ACCUMULATE",
  phaseLabel: "強力累積區",
  action: "NORMAL_RESTOCK",
  restockAllowed: true,
  targetLtvMin: 0.25,
  targetLtvMax: 0.30,
  allocationBias: "L1_ACCUMULATE",
  severity: "INFO",
  reason: "BTC_MM < 1.0"
}
```

V1 先將此物件放在：

```javascript
context.market.btcRegime
```

後續所有顯示與建議都從這個物件取值。

### 5.3 Regime 判定順序

判定順序必須先處理生存風險，再處理機會區。

| 優先序 | 條件 | Regime | 行動 |
|---:|---|---|---|
| 1 | `cryptoLTV >= 0.50` | `HARD_CAP_BREACH` | 強制去槓桿 |
| 2 | `cryptoLTV >= 0.45` | `DEFCON_1` | 停止所有買入，優先還款 |
| 3 | `btcMM >= 2.10` | `BUBBLE` | De-risk / 降 LTV |
| 4 | `btcMM >= 1.50` | `DE_RISK` | 停止加槓桿，現金流還款 |
| 5 | `btcMM < 0.80` 或 `btcDrawdown <= -0.50` | `PANIC_ACCUMULATE` | 僅在 LTV 安全時低點累積 |
| 6 | `btcMM < 1.00` | `ACCUMULATE` | 正常累積 |
| 7 | `btcMM < 1.15` 且 `btcDrawdown <= -0.30` | `EXTENDED_ACCUMULATE` | 降低強度延長累積 |
| 8 | `btcMM < 1.50` | `NEUTRAL` | 固定 DCA，不新增借款買入 |
| 9 | 其他 | `MANUAL_REVIEW` | 手動檢查 |

備註：

- `btcDrawdown === null` 時，不允許觸發 `EXTENDED_ACCUMULATE`
- `PANIC_ACCUMULATE` 不代表無限制提高 LTV，仍受 LTV guardrail 約束
- `Survival > Profit` 是最高原則

### 5.4 LTV target 規格

V1 使用區間，不再只顯示單一 target：

| Regime | Target LTV |
|---|---:|
| `PANIC_ACCUMULATE` | `0.30 - 0.40` |
| `ACCUMULATE` | `0.25 - 0.30` |
| `EXTENDED_ACCUMULATE` | `0.20 - 0.30` |
| `NEUTRAL` | `0.20 - 0.25` |
| `DE_RISK` | `0.00 - 0.10` |
| `BUBBLE` | `0.00` |
| `DEFCON_1` | `0.00` |
| `HARD_CAP_BREACH` | `0.00` |

`restockAllowed` 依 LTV 另行 gate：

```text
cryptoLTV < 32.5%: 可允許一般 restock
32.5% <= cryptoLTV < 35%: 只允許現金流 DCA
35% <= cryptoLTV < 45%: 停止新增借款
45% <= cryptoLTV < 50%: DEFCON_1，停止所有買入
cryptoLTV >= 50%: Hard cap breach，強制去槓桿
```

### 5.5 Allocation target 規格

人工 override 仍最高優先：

```text
Alloc_L*_Target exists -> use manual target
else -> use btcRegime allocation mapping
```

V1 mapping：

| Regime | L1 | L2 | L3 |
|---|---:|---:|---:|
| `PANIC_ACCUMULATE` | `0.75` | `0.15` | `0.10` |
| `ACCUMULATE` | `0.70` | `0.20` | `0.10` |
| `EXTENDED_ACCUMULATE` | `0.68` | `0.22` | `0.10` |
| `NEUTRAL` | `0.65` | `0.25` | `0.10` |
| `DE_RISK` | `0.60` | `0.30` | `0.10` |
| `BUBBLE` | `0.50` | `0.30` | `0.20` |
| `DEFCON_1` | 不追配置 | 優先降債 | 優先保留現金 |
| `HARD_CAP_BREACH` | 不追配置 | 優先降債 | 優先保留現金 |

若 regime 是 `DEFCON_1` 或 `HARD_CAP_BREACH`，再平衡建議不應輸出「補強 L1」作為首要行動，避免在風險紅線時仍建議買入 BTC。

---

## 6. 報告輸出規格

### 6.1 市場情報

原本：

```text
- Mayer Multiple: 1.02
- 週期定位: 正常累積區
```

改為：

```text
- Mayer Multiple: 1.02
- 距離 ATH: -34.8%
- BTC Regime: Extended Accumulate
- 週期定位: 熊市均線下修延長累積區
```

若 drawdown 不可用：

```text
- 距離 ATH: N/A（缺少 SAP_Base_ATH）
```

### 6.2 生存指標

原本只顯示單一 target：

```text
- 目標 LTV (Crypto) (建議): 25%
```

改為：

```text
- 目標 LTV (Crypto): 20% - 30%
- 活性 LTV (Active): 28.4%
- 總體 LTV (Global): 18.1%
- Restock Mode: EXTENDED / reduced size only
```

若超過 guardrail：

```text
- Restock Mode: OFF
  > Active LTV 已超過 35%，停止新增借款，只允許還款或現金流 DCA。
```

---

## 7. 實作計畫

### Phase 1：規格與測試樣本

狀態：已完成

工作：

- 確認 regime threshold
- 確認不新增 sheet key
- 列出 expected behavior examples

### Phase 2：集中式 helper

狀態：已完成

修改：

- `Core_StrategicEngine.js`

工作：

- 新增 `calculateBtcDrawdownFromATH_()`
- 新增 `getBtcRegime_()`
- 新增 `getBtcAllocationTargets_()`
- 將 `context.market.btcRegime` 接入 `buildContext()`

### Phase 3：接入現有策略輸出

狀態：已完成

修改：

- `Core_StrategicEngine.js`

工作：

- `assetGroups` 動態配置改吃 regime
- `generatePortfolioSnapshot()` 改吃 regime
- Crypto LTV target 改成區間
- 高 LTV guardrail 顯示明確 action

### Phase 4：測試與驗證

狀態：進行中

工作：

- `node --check Core_StrategicEngine.js`
- 建立或補充 strategy helper 測試
- 用情境表驗證：
  - `MM 1.08 + drawdown -35% -> EXTENDED_ACCUMULATE`
  - `MM 1.08 + drawdown -10% -> NEUTRAL`
  - `MM 0.95 + LTV 34% -> ACCUMULATE but DCA_ONLY`
  - `MM 0.75 + LTV 46% -> DEFCON_1`
  - `MM 2.2 + LTV 12% -> BUBBLE`

目前已完成：

- `node --check Core_StrategicEngine.js` 通過
- 本地 Node VM 情境測試通過：
  - `MM 1.08 + drawdown -35% -> EXTENDED_ACCUMULATE`
  - `MM 1.08 + drawdown -10% -> NEUTRAL`
  - `MM 0.95 + LTV 34% -> ACCUMULATE / DCA_ONLY`
  - `MM 0.75 + LTV 46% -> DEFCON_1`
  - `MM 2.2 + LTV 12% -> BUBBLE`

---

## 8. 驗收標準

功能驗收：

- `BTC_MM` 不再是唯一 BTC 週期判定來源
- `SAP_Base_ATH` 缺失時系統仍可正常 fallback
- 人工 `Alloc_L*_Target` override 行為不變
- `cryptoLTV >= 45%` 時不再輸出任何 BTC 買入優先訊息
- `BTC_MM 1.0 - 1.15` 且 drawdown 深時，能顯示延長累積區

品質驗收：

- 不新增 sheet key
- 不改 workbook contract
- 不破壞 Martingale Sniper 現有觸發
- `node --check Core_StrategicEngine.js` 通過
- 報告文字能清楚說明 regime reason

---

## 9. 開放問題

1. `SAP_Base_ATH` 是否要長期定義為「歷史 ATH」，還是「策略定錨高點」？
2. `EXTENDED_ACCUMULATE` 的 L1 target 要用 `68%` 還是維持現有 `65%`？
3. `PANIC_ACCUMULATE` 是否應要求 `survivalRunway >= 6` 才允許 restock？
4. Martingale Sniper 是否要受 `cryptoLTV >= 35%` 的新增借款禁令約束？

本文件的預設答案：

- V1 將 `SAP_Base_ATH` 視為策略定錨高點
- V1 採 `EXTENDED_ACCUMULATE = L1 68%`
- V1 在 `survivalRunway < 6` 時將 restock 降為 `DCA_ONLY`
- V1 建議 Martingale 訊號仍可顯示，但 action 必須標示 LTV guardrail
