# 護理排班系統（Streamlit）

以表單輸入條件，自動產生 Excel 班表。

## 儲存庫內主要檔案

| 檔案 | 說明 |
|------|------|
| `streamlit_app.py` | 網頁介面（Streamlit Cloud 入口） |
| `nurse_scheduler_v2.py` | 排班核心與 Excel 匯出 |
| `config.sample.json` | 預設／儲存用設定 |
| `requirements.txt` | Python 依賴 |
| `.streamlit/config.toml` | Streamlit 設定 |

本機若需要「雙擊啟動」或範例圖，可自行放在資料夾內，**不必**放進 GitHub（避免與雲端無關、或體積較大的檔案污染 repo）。

## 本機執行

需安裝 Python 3.10+，於專案目錄執行：

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## 部署到 Streamlit Community Cloud（網址開啟）

1. 將此專案推上 **GitHub** 公開儲存庫（Private 需付費方案）。
2. 登入 [Streamlit Community Cloud](https://streamlit.io/cloud)。
3. **New app** → 連結你的 GitHub → 選擇此 repo。
4. 設定：
   - **Branch**：`main`（或你實際使用的分支）
   - **Main file path**：`streamlit_app.py`
5. 部署完成後會得到 `https://<你的應用名稱>.streamlit.app` 網址。

### 推送到 GitHub（在本機專案目錄）

若尚未建立遠端儲存庫，請先在 GitHub 網頁 **New repository** 建立空 repo（不要勾 README）。

```bash
git init
git add .
git commit -m "Initial commit: nurse schedule Streamlit app"
git branch -M main
git remote add origin https://github.com/<你的帳號>/<repo名稱>.git
git push -u origin main
```

之後在 Streamlit Cloud 重新 **Redeploy** 即可更新網站。

## 注意事項

- 排班完成後請按頁面上的 **「下載 Excel」**，雲端無法像本機一樣自動開啟檔案總管裡的 xlsx。
- `config.sample.json` 在雲端冷啟後會還原成 repo 內版本（屬正常現象）。
- 若需長期保存每位使用者的設定，需另行接資料庫或雲端儲存（可再擴充）。

## 無法開啟時請檢查

### 本機

- 請用：`python -m streamlit run streamlit_app.py`（或雙擊 `啟動排班系統.bat`）。
- **不要**用 `python streamlit_app.py` 直接跑（不會正常出現網頁）；若誤用，已內建會嘗試自動改呼叫 `streamlit run`。

### Streamlit Cloud

- 免費版 App 一陣子沒人使用會**休眠**，第一次開可能要等 30～90 秒，請重新整理。
- 到 Cloud 後台該 App → **Logs**，看紅字是否為套件安裝失敗或 `config.sample.json` 找不到。
- **Main file path** 必須是：`streamlit_app.py`（在 repo 根目錄）。
- 若出現 **Error during processing dependencies**：請確認 repo 根目錄有 **`runtime.txt`**（指定 Python 3.11）與 **`requirements.txt`**（勿含中文註解，避免 pip 編碼錯誤）；修改後 **push** 並在 Cloud **Redeploy**。
