# Excel 範本下載與更新紀錄

更新日期：2026-11-23

## 最新版：SRAM Extra BLK Failure Analysis v13.4

**[下載 v13.4 修復版 Excel（.xlsm）](https://raw.githubusercontent.com/LienHC/asc-wafer-fail-address/main/current/templates/SRAM%20Extra%20BLK%20Failure%20%20Analysis_v13.4.xlsm)**

[檢視檔案](SRAM%20Extra%20BLK%20Failure%20%20Analysis_v13.4.xlsm) · [完整更新紀錄](../RELEASE_NOTES.md) · [操作說明](../README.md) · [W/R 報表說明](../docs/WR_PERFORMANCE.md)

完整 VBA 已內嵌。下載後以 Windows 桌面版 Excel 開啟，依公司規範允許此活頁簿的巨集，再使用「匯入並分析」或「Refresh Maps」。日常執行不需 Python，也不需手動匯入 VBA。

檔名為 `SRAM Extra BLK Failure  Analysis_v13.4.xlsm`；`Failure` 與 `Analysis` 之間保留兩個空白。此下載連結指向 `main` 上已合併修復的檔案。

## 2026-11-20 更新內容

- 修復開啟 Excel 時「載入 WRPerformance 時發生錯誤」。
- 以原生 Excel 重建 VBA 專案，只替換 VBA 封裝；保留全部 19 個模組的程式邏輯、原有工作表、按鈕及設定。
- 保留 W/R、Scan W-R 報表與 Source Row 來源列跳轉功能。
- 修正 W/R 測試入口與過期列號，新增報表儲存重開檢查。
- 修復已透過 [PR #2](https://github.com/LienHC/asc-wafer-fail-address/pull/2) 合併至 `main`。

## 驗證結果

本機 Excel 16.0 原生測試已通過：

- W/R 與 Scan W-R 報表：3 WT／2 Pattern、15 種精確組合、灰底 Chip 排除與計數。
- 重跑、零分母與報表失效清理。
- Fail Analysis／Chip Fail Details 的 Source Row 選取跳轉，包含無效選取與事件狀態恢復。
- 儲存後重新開啟，報表數值與圖表保持完整。
- 靜態核對全部 19 個模組及非 VBA 活頁簿內容。

驗證採合成資料，尚未以公司正式 ASC 執行完整匯入；Excel 2016 實機相容性亦未另行驗證。

## Address Data Auto output v1.8

[下載 v1.8（.xlsm）](Address%20Data%20Auto%20output_v1.8.xlsm?raw=true)

Fail Address 的 G 欄 `BL=1` 沿用原 Cat 數值，H 欄 `BL=0` 顯示所選模式的分類，TNO 從 I 欄開始。

- **First Fail Cat（預設）** 從第一個出現 Fail Bit 的站點取得 `Config!G` 的 Cat。
- **C20 in WT2/WT3** 若 WT1 Cat 不屬於 C1／C2，且該 Address 首次出現在 WT2／WT3，H 欄輸出 20。WT1 已出現的 Address 沿用首次 Fail Cat。

在 `Config!K12` 選擇模式後執行批次匯出。已用 WT1／WT2 模擬 ASC 確認兩種模式分別輸出 9 與 20。

## 其他範本與歷史版本

| 檔案 | 下載 | 說明 |
| --- | --- | --- |
| Address Data Auto output v1.6 | [下載 .xlsm](https://raw.githubusercontent.com/LienHC/asc-wafer-fail-address/main/current/templates/Address%20Data%20Auto%20output_v1.6.xlsm) | 獨立工具，使用自己的版本編號 |
| SRAM Analysis v13.3 | [下載歷史版](https://raw.githubusercontent.com/LienHC/asc-wafer-fail-address/main/current/templates/SRAM%20Extra%20BLK%20Failure%20%20Analysis_v13.3.xlsm) | 新增 Scan W-R Performance；本次 VBA 載入修復請使用 v13.4 |
| SRAM Analysis v13.2 | [下載歷史版](https://raw.githubusercontent.com/LienHC/asc-wafer-fail-address/main/current/templates/SRAM%20Extra%20BLK%20Failure%20%20Analysis_v13.2.xlsm) | 新增 W&R Performance |
| SRAM Analysis v13.1 | [下載歷史版](https://raw.githubusercontent.com/LienHC/asc-wafer-fail-address/main/current/templates/SRAM%20Extra%20BLK%20Failure%20%20Analysis_v13.1.xlsm) | 保留供歷史比對 |

[返回專案首頁](../../README.md)
