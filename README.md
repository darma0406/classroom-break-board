# 下課狀態看板

這個專案是舊版 GitHub Pages 下課狀態看板。它保留原本的 Google Sheet 顏色掃描 fallback，同時可優先讀取 Classroom OS 寫入的 `_ClassroomOS_BreakBoard` mirror sheet。

## 本機啟動

```bash
npm start
```

開啟：

```text
http://localhost:4173
```

也可以使用 GitHub Pages 正式網址。

## Classroom OS Mirror Read 架構

正式資料流：

```text
Classroom OS
-> Apps Script doPost
-> Google Sheet _ClassroomOS_BreakBoard
-> Apps Script doGet/getStatus
-> GitHub Pages index.html fetch
-> data normalization
-> student card render
```

`_ClassroomOS_BreakBoard` 使用下列表頭：

- 座號
- 姓名
- 未交
- 訂正
- 未完成
- 可下課
- 禁用平板
- 同步時間

讀取端會依表頭名稱找欄位，不依賴固定欄位位置。

## Apps Script 檔案

本 repo 提供兩份 Apps Script 來源：

- `apps-script.gs`: 舊版 board API 參考檔。
- `apps-script-classroom-os-mirror-merged.gs`: 可貼入 Google Apps Script 的整合版 board 讀取程式。

整合版重點：

- `getStatus()` 優先讀 `_ClassroomOS_BreakBoard`。
- mirror 有有效學生資料時，回傳 `source: "ClassroomOSMirror"` 與 `classroomOSMirror: true`。
- mirror 不存在或無有效學生資料時，fallback 到 legacy color scan。
- 保留 `manualStatus`、`classCalm`、`updateStatus()`、`updateAllStatus()`。
- 不宣告 top-level `doPost(e)`，避免破壞 Classroom OS 寫入 sync。

## Google Apps Script 部署步驟

目前 Google Apps Script 專案若同時包含 `board.gs` 與 `webapp.gs.gs`，請照以下方式處理：

1. 開啟 Google Apps Script 專案。
2. 將 `apps-script-classroom-os-mirror-merged.gs` 的內容貼到 `board.gs`，覆蓋原本 `board.gs`。
3. 保留 `break-board-sync.gs`。
4. 保留 `webapp.gs.gs` 裡已成功運作的 Classroom OS 寫入 `doPost(e)`。
5. 確認 Apps Script 專案中只有一個正式 top-level `doPost(e)`。
6. 建立新部署版本，建議使用下一個版本號，例如 Version 6。
7. Web App URL 可維持原本 `/exec` URL。

如果仍需要舊版 POST action，請不要在 `board.gs` 新增第二個 top-level `doPost(e)`。應由唯一的 top-level `doPost(e)` 做 payload routing，或呼叫整合檔中的 `legacyBoardDoPostHandler(e)`。

## 測試方式

先在 Apps Script 編輯器執行：

```javascript
testClassroomOsMirrorRead()
```

預期回傳：

- `students` 有資料
- `updatedAt` 來自 `_ClassroomOS_BreakBoard` 的同步時間
- 若缺少 `座號` 或 `姓名` 表頭，會 throw error

部署後測試：

```text
https://script.google.com/macros/s/AKfycby5h0QWJ5TSAR7IVE1s7QrgYoNAKbRiIu6HmJXdyPofyaF55rHvuXIpxhC8S7jRl76rfg/exec?action=getStatus
```

正確 JSON 必須包含：

```json
{
  "source": "ClassroomOSMirror",
  "classroomOSMirror": true
}
```

再測試正式看板：

```text
https://darma0406.github.io/classroom-break-board/?class=rpes114503&v=classroom-os-mirror-final
```

開啟後按「更新名單」。

## Legacy Fallback

若 `_ClassroomOS_BreakBoard` 不存在或沒有有效學生資料，`getStatus()` 會 fallback 到原本的 Google Sheet 顏色掃描：

- 紅色作業格 -> `missingAssignments`
- 黃色作業格 -> `missingCorrections`
- `incompleteAssignments` 補空陣列
- `canBreak` 補 `true`
- `tabletBlocked` 補 `false`
- `source` 回傳 `LegacyColorScan`
- `classroomOSMirror` 回傳 `false`

這個 fallback 用來保護舊版正式看板，不應移除。
