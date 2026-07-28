# FitLog AI

> 永森健康管理系統 — AI 食物分析與健身追蹤（網頁 App + Google Apps Script 後端）

![FitLog AI 預覽](https://rock903400-byte.github.io/wind/assets/fitlog-ai.webp)

## 功能

- 🔍 **AI 食物分析**：文字描述或拍照上傳，GPT-4o Vision 模型即時辨識食物並估算熱量，附營養師級建議
- ⚖️ **體重追蹤**：記錄體重變化，自動計算 TDEE 與每日攝取目標
- 📊 **每日儀表板**：今日攝取量、七日趨勢圖、連續記錄天數
- 🔐 **帳號系統**：密碼登入、嘗試鎖定機制、多使用者支援
- 📱 **手機最佳化**：響應式設計，以手機操作為主

## 架構

| 層級 | 技術 |
|---|---|
| **Production 後端** | Google Apps Script + Google Sheets |
| **AI 分析** | GitHub Models API — GPT-4o Vision（文字＋圖片） |
| **Standalone 版** | HTML / CSS / JavaScript（SPA）+ LocalStorage（GitHub Pages） |

Production 版透過 Google Apps Script Web App 部署，後端連接 Google Sheets 作為資料庫，
前端透過 GAS API 進行登入、紀錄、AI 分析等操作。GitHub Pages 版為純本機獨立版本，
免登入、免後端，功能較精簡（無 AI 分析）。

## Demo

https://rock903400-byte.github.io/FitLog-AI/

## License

MIT
