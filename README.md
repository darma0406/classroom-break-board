# classroom-break-board

適合國小班級的即時下課狀態與作業管理看板。

## 一、功能介紹

- 即時下課狀態
- 平板管理
- 今日提醒
- 一般叮嚀
- 自動倒數
- Google Sheet 作業同步
- 未交作業（紅色）
- 未訂正作業（黃色）
- 多裝置同步
- 本機備份
- JSON 匯出

## 二、快速開始

1. 開啟 GitHub Pages。
2. 輸入班級代碼。

   建議格式：

   `rpes114501`

3. 進入看板。

## 三、不使用 Google Sheet

如果暫時不連 Google Sheet，也可以直接用看板。

- 手動調整學生數量
- 手動修改姓名
- 快速貼上名單

範例：

```text
1 王小明
2 林小美
```

## 四、如何連結 Google Sheet

1. 建立 Google Sheet。
2. 開啟：`擴充功能 → Apps Script`
3. 刪除原始碼。
4. 在系統裡按：`複製 Apps Script 程式碼`
5. 貼上。
6. 部署 Web App。
7. 複製 `/exec` 網址。
8. 貼回看板。

## 五、Google Sheet 規則

Google Sheet 會自動同步作業狀態。

- 紅色：未交作業
- 黃色：未訂正作業

系統會自動同步，不需要每次手動整理。

## 六、顏色說明

- 綠色：可下課
- 紅色：有未交作業
- 黃色：有未訂正作業

## 七、平板管理

平板管理會即時同步。

- 使用平板
- 禁用平板

## 八、備份

- 系統會自動 local backup
- 可匯出 JSON

## 九、常見問題

Q：為什麼設定按鈕打不開？

A：先確認頁面是否完整載入，並檢查瀏覽器是否擋住彈出視窗或腳本錯誤。

Q：為什麼 Google Sheet 沒同步？

A：先確認 Apps Script 網址已貼好，並且 Web App 已正確部署。

Q：為什麼 Apps Script 不能執行？

A：通常是程式碼未貼完整、部署權限不正確，或 `/exec` 網址不是最新版本。

Q：如何換班級？

A：到設定頁輸入新的班級代碼，再切換進入即可。

Q：如何備份？

A：系統會自動做本機備份，也可以在設定中匯出 JSON。

## 十、部署資訊

- GitHub Pages
- Firebase Realtime Database
- Google Apps Script

## 十一、風格

- 簡單
- 老師看得懂
- 避免工程術語
- 使用 markdown
- 適合 GitHub 首頁閱讀
