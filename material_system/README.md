# Excel 物料版本比對與更新系統 MVP

這是一個可在本機啟動的第一階段 MVP，針對 `QP-QA-43-08原物料&包材物質清單2026.xlsx` 的矩陣格式與 `WI-QA-199-01-有害物質調查表Ver23.xlsx` 的調查表格式實作。

## 啟動

```powershell
& 'C:\Users\sammi.hung\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' .\material_system\app.py
```

開啟：

```text
http://127.0.0.1:8000
```

## 已完成

- 上傳總表與調查表
- Excel 工作表、列數、欄數與前 20 列預覽
- Composite Key 設定
- 依參考格式解析 `原物料`、`包材` 與 `有害物質調查表`
- 新增、刪除、修改、相同、衝突比對
- 分頁差異確認
- 批次接受新版、保留舊版、確認新增、標記停用、暫不處理
- 人工確認後產生新版總表
- 匯出 Search Report
- SQLite 保存檔案、版本、比對、確認與稽核紀錄

## 測試方式

1. 上傳 `QP-QA-43-08原物料&包材物質清單2026.xlsx` 作為總表。
2. 輸入調查表欄位，例如廠商代號 `0001102290`、化學物質類別 `原物料`。
3. 上傳 `WI-QA-199-01-有害物質調查表Ver23 0608.xlsx` 或其他調查表。
4. 選擇總表工作表後按「開始版本比對」。
5. 在差異表勾選資料並選擇處理方式。
6. 全部非衝突差異確認後，按「套用修改並匯出」。

## 目前限制

- MVP 使用 Python 標準庫 HTTP Server，未安裝 FastAPI / Next.js 時仍可直接執行。
- 調查表目前視為單一物料更新，物料名稱由調查表 `品名Item No.` 解析。
- 若 Composite Key 衝突，系統會阻止套用，需要先調整資料或 Key 設定。
