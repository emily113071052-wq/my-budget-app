# 💰 我的個人記帳本

Streamlit + Google Sheets 的個人記帳系統。

## 功能
- 支出 / 收入 / 轉帳（跨幣別自動抓匯率換算 TWD）
- 類別 → 子類別 → 商家/品項 → 備註 四層記錄
- 帳戶資產走勢圖、總資產組成圓餅圖、單一帳戶檢視
- 每月預算進度條與超支警示
- 訂閱 / 定期支出到期提醒
- 明細篩選、搜尋、直接編輯 / 刪除

## 本機執行
```
pip install -r requirements.txt
streamlit run app.py
```
需在 `.streamlit/secrets.toml` 放入 Google 服務帳戶金鑰（`gcp_service_account` 區塊）。
此檔已被 .gitignore 排除，**請勿上傳**。

## 部署（Streamlit Community Cloud）
1. 把這個 repo 推上 GitHub（私人）
2. 到 https://share.streamlit.io 用 GitHub 登入 → New app → 選這個 repo 和 app.py
3. 在 App settings → Secrets 貼上本機 `.streamlit/secrets.toml` 的完整內容
4. Settings → Sharing 設為私人，只允許自己的帳號檢視
