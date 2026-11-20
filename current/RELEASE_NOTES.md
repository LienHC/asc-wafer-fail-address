# v13.4

## 2026-11-20 — VBA 載入修復與執行驗證

- **修復**：開啟範本時出現「載入 WRPerformance 時發生錯誤」。以原生 Excel 重建 VBA 專案，只替換 `xl/vbaProject.bin`；保留全部 19 個模組的程式邏輯，以及原有工作表、按鈕與設定。
- **Source Row**：保留選取 Fail Analysis／Chip Fail Details 的來源列，跳到 Fail Address 對應列的功能。
- **測試工具**：修正 W/R 原生測試對 `Option Private Module` 的呼叫方式與過期列號；靜態檢查接受 Excel 自動調整的識別字大小寫與空白，仍精確比對字串常值。
- **原生 Excel 16.0 驗證通過**：W/R 與 Scan W-R 報表、3 WT／2 Pattern、15 種精確組合、灰底排除、重跑、零分母、報表失效清理及 Source Row 選取事件。
- **儲存重開通過**：報表數值與圖表保持完整；靜態檢查確認非 VBA 活頁簿內容未變更。
- **驗證範圍**：採合成資料，尚未以公司正式 ASC 執行完整匯入。
- **合併紀錄**：[PR #2](https://github.com/LienHC/asc-wafer-fail-address/pull/2) 已合併至 `main`；修復提交 [72d6228](https://github.com/LienHC/asc-wafer-fail-address/commit/72d622812231006eac27e8f24a102056830edb6d)，合併提交 [42d4860](https://github.com/LienHC/asc-wafer-fail-address/commit/42d486098c49a5b28496b418caf0d74daceda814)。

---

# v13.3

新增 `Scan W-R Performance`。每個已產生 Scan Map 的 WT 都有七欄百分比、各欄 Fail Chips 顆數和 DC Pass Count。灰底 Chip 從分子和分母排除。

`W&R Performance` 保留 W、R、RW、X 獨立計數，並列 15 種精確組合的顆數、百分比和圖表。

---

# v13.2

新增 W&R Performance 工作表，依 Scan、March 等 Pattern 分表，並列已匯入 WT 的 W、R、RW、X Chip 百分比。

每顆 Chip 在同一 WT、Pattern、類別最多計一次。分母為全部 Raw Chip 扣除 WT1 Cat C10、C11、C14、C15、C16、C17。右側四欄只顯示百分比。Lot 使用前六碼，Wafer 顯示 #01 格式。

匯入並分析與 Refresh Maps 自動更新。Calibri 11 置中、兩層表頭與分類配色。

已以原生 Excel 驗證 1～3 份合成 ASC 完整匯入、3-to-1 重跑、六個 Config 指定巨集、Wafer Maps 雙擊至 BLK Pattern Distribution、DC Cat 排除、去重、零分母及儲存重開。

---

# Wafer ASC 分析 v0.26.0

[下載正式 VBA Excel](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.26.0/ASC_Wafer_Analysis_Map_Layout.xlsm) · [下載64顆Chip Demo](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.26.0/ASC_Wafer_Analysis_Map_Layout_Demo64.xlsm) · [下載完整套件](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.26.0/ASC_Wafer_Analysis_Map_Layout_Latest.zip)

- **Map版面**：縮放70%、欄寬6、Map區列高22，字體加粗；四宮格右下數量改RGB(128,128,128)灰字。導覽按鈕配合欄寬縮排，不重疊。
- **四宮格標題**：例如 `X768#11 BL=0 WT1 Scan Fail`；March／WT2／WT3依對應Pattern與WT命名。Lot先取前6碼，再去掉前2碼，Wafer補足兩位。
- **Fail BLK Count標題**：例如 `WT1 X768#11 1.20v~0.56v Scan Fail BLK`。實際查找該Pattern Function區間最高／最低電壓，不只取首尾測試點。缺電壓顯示`V?~V?`。
- **BL=0僅為標題文字**：統計仍包含BL0～15，所有Fail數字維持不變。
- **四宮格藍底**：Cat固定採WT1；Cat1～4且當前WT／Pattern終點完整256個BLK FBC全部為0時，整Chip填RGB(0,176,240)。缺資料不當0、Cat0／5以上不依此條件填藍；已有Fail數字時優先保留。
- 保留v0.25.0的BLK電壓雙擊、Observed W/R分析、Source Row、合併Map與一鍵PPT，以及Fail Address凍結與格線。
- 原生Excel驗證：384個Count與1536個四格數值不變；384組獨立Cat／終點FBC條件、Cat0～5邊界、非首尾站電壓極值、標題、數字優先與儲存重開均通過。

全部VBA已內嵌，執行不需Python；採Excel2016相容VBA。詳見 [Map版面與標題規則](MAP_LAYOUT.md)。

---

以下為歷史版本說明；最新版以上方v0.26.0為準。

# Wafer ASC 分析 v0.25.0

[下載正式 VBA Excel](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.25.0/ASC_BLK_Voltage_Analysis.xlsm) · [下載64顆Chip Demo](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.25.0/ASC_BLK_Voltage_Analysis_Demo64.xlsm) · [下載完整套件（含VBA及模擬ASC）](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.25.0/ASC_BLK_Voltage_Analysis_Latest.zip)

- **Scan／March Map雙擊Chip → Chip Fail Details**：16×16 BLK圖按欄往下排列，第一欄BLK0～15，第二欄BLK16～31，右下BLK255。
- 格內顯示依Config順序、第一筆Function FBC>0的測試電壓；不是完整BLK Vmin，也不把TNO自行排序。
- 底色顯示已捕捉Address的Observed W/R傾向：Write worse／Read worse／Equal／Mixed／Insufficient。只有雙側有效的Monitor失效電壓才比較；缺紀錄不當Pass。星號表示部分捕捉樣本仍無法比較。
- **再雙擊BLK**：查看Scan／March電壓-FBC圖、各TNO原始數據、逐Address W/R比較及Source Row跳轉。沿用Hard Fail、前WT全部Function及Scan→March排除規則。
- **Wafer Maps合併導覽**：各WT／Pattern的Fail BLK、W&R四宮格及FailBLK Loss二宮格集中同頁；原始來源隱藏保留，原Config五個按鈕保留。
- **一鍵PPT**：寬37公分單頁，上排六張Fail BLK圖，下排六張分Pattern四宮格，各下排圖右下方另放縮小二宮格，不重疊，使用點陣圖。需安裝桌面版PowerPoint。[PPT範例](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.25.0/ASC_Wafer_Maps_Demo64.pptx)
- Fail Analysis整合Chip／BLK／Monitor Address明細；Import Info整合路徑、座標與執行訊息。Fail Address固定第1列及A:G，保留細格線與TNO註解。
- 原生Excel驗證：2,304筆獨立ASC BLK預期值、9組Chip情境、W/R混合及缺資料、真實雙擊事件、Source Row、Scan Pass後March Fail、儲存重開均通過。前期6,809個Address/TNO旗標回歸亦保留紀錄。

全部VBA已內嵌，執行不需Python或手動匯入程式碼；採Excel2016相容VBA，已以本機桌面版Excel驗證。iPad可檢視已輸出檔案；VBA互動需桌面Excel。公司實際ASC／Excel2016環境仍請以正式資料確認。

操作細節：[BLK電壓分析](current/BLK_VOLTAGE_ANALYSIS.md) · [合併Map與PPT](current/MAP_DASHBOARD_PPT.md)。WL／BL／IO三維圖仍屬規劃，未加入本版。

---

以下為歷史版本說明；最新版以上方v0.25.0為準。

# v0.18.0 - Stop Sign endpoint Fail Maps

- Config G1 Stop Sign is red/white; G Pattern Name selects the same-row Function TNO in D/E/F as the cumulative endpoint. Monitor A/B/C never participates in Map Judge.
- Count only endpoint BLKs with FBC>0, never re-add earlier TNOs. Green0 Pass, red1..256 Fail BLK Count, gray No Data, yellow Data Error.
- Raw paired X/Y inventory includes DUT/DC/ISB/no-FBC records and is shared across WT1/2/3. Map geometry is data-driven, with one three-WT sheet per Pattern.
- Add Chip Locations and BLK Fail Count outputs; J19 rebuilds Maps without Raw reimport. J13 remains Address-only; J16 imports FBC and builds Maps.
- Fix native VBA EOF call: explicit Call FinishRecord prevents a line-start label from dropping the last record. Native tests cover both complete and incomplete final records.
- Passed native Excel16 tests for synthetic64/128 Chips, two Patterns, three WTs, endpoint501-only0/3/256 counts, missing/DC-only gray, parser-error yellow, managed Map cleanup and Address preservation. Also passed original Address regression and Office2016 format checks. Company ASC/environment remain untested.
