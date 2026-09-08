# 法規鑑別追蹤工具

用 [全國法規資料庫](https://law.moj.gov.tw)（law.moj.gov.tw）追蹤法規修正、比對新舊條文差異，
並把每一條的適用性、守規性評估、鑑別人員、是否結案等紀錄存下來，隨時可以編輯、
匯出成 Excel（欄位對齊 HPA01-02B 法規鑑別表格式）。

## 資料夾內容

| 檔案 | 用途 |
|---|---|
| `law_tool_app.py` | 主程式（Flask 網頁版工具本體） |
| `moj_law_diff.py` | 命令列版工具（進階／批次比對用，非必要） |
| `requirements.txt` | 需要的 Python 套件 |
| `Procfile` | 告訴雲端主機要怎麼啟動這個服務（部署用，本機執行不需要） |
| `start_law_tool.bat` | 本機使用：雙擊即可開啟工具 |
| `first_time_setup.bat` | 本機使用：換到新電腦時執行一次，安裝套件 |
| `law_tool.db` | 資料庫（第一次啟動後自動產生，不會被上傳到 GitHub） |

---

## 方法一：在自己電腦上執行（原本的用法）

```
pip install -r requirements.txt
python law_tool_app.py
```

或直接雙擊 `start_law_tool.bat`（第一次用新電腦要先跑一次 `first_time_setup.bat`）。
執行後會自動開瀏覽器開啟 `http://127.0.0.1:5050`，資料存在同資料夾的 `law_tool.db`。

---

## 方法二：部署到 Railway，變成隨時可用的網址

這樣不管在哪台電腦，開瀏覽器輸入同一個網址就能用，不用再依賴某一台電腦、
也不用每次手動啟動程式。

### ⚠️ 先看這個：一定要接「永久儲存空間（Volume）」

Railway（以及大多數雲端主機）預設的檔案系統是**暫時性的**：服務只要重新部署、
重新啟動，容器裡的檔案就會被清空還原成最初的樣子。這支工具的資料庫
`law_tool.db` 存的是你所有追蹤的法規、鑑別評估紀錄，**如果沒有另外接上
「永久儲存空間」，這些資料在下一次重新部署或重啟時會完全消失**。

所以部署步驟裡，「加 Volume、設定 `LAW_TOOL_DB_PATH` 指到 Volume 路徑」這一步
絕對不能跳過，一定要在你開始正式使用（新增法規、填鑑別評估）之前就設定好。

### 部署步驟

1. **把這個資料夾推上 GitHub**（建議設成 Private 私人倉庫）
   - 到 github.com 建立一個新的 repository（例如叫 `law-tool`），不要勾選自動產生
     README（這個資料夾已經有一個了）
   - 在這個資料夾裡執行：
     ```
     git init
     git add .
     git commit -m "法規鑑別追蹤工具"
     git branch -M main
     git remote add origin https://github.com/<你的帳號>/law-tool.git
     git push -u origin main
     ```

2. **在 Railway 建立新專案**
   - 登入 [railway.app](https://railway.app) → New Project → **Deploy from GitHub repo**
   - 選剛剛推上去的 `law-tool` 這個 repo
   - Railway 會自動偵測到 `requirements.txt` 和 `Procfile`，用 gunicorn 啟動，
     不用額外設定啟動指令

3. **加一個 Volume（永久儲存空間）**
   - 進到這個服務的設定頁 → **Volumes** → **New Volume**
   - Mount Path 填一個路徑，例如 `/data`
   - 儲存空間 1GB 就綽綽有餘（這個資料庫只存文字，不會很大）

4. **設定環境變數**
   - 到服務的 **Variables** 分頁，新增：
     - `LAW_TOOL_DB_PATH` = `/data/law_tool.db`　（路徑要跟上一步 Mount Path 一致，
       檔名 `law_tool.db` 可以自己取，重點是資料夾要對到 Volume）
   - `PORT` 這個變數 Railway 會自動提供，不用自己設

5. **重新部署（Redeploy）一次**，讓新的環境變數生效

6. **產生對外網址**
   - 到 **Settings → Networking**，點 **Generate Domain**，Railway 會給你一個
     像 `law-tool-production.up.railway.app` 的網址
   - 之後在任何電腦的瀏覽器打開這個網址，就是同一份工具、同一份資料

### 之後要更新程式時

以後在自己電腦改完程式碼、要更新到線上版本，只要：
```
git add .
git commit -m "說明這次改了什麼"
git push
```
Railway 偵測到 GitHub 有新的 commit 就會自動重新部署（因為 Volume 是分開的，
重新部署不會影響已經存進去的資料）。

### 關於這個網址的隱私

這次設定沒有加登入密碼保護，網址本身沒有公開分享、也不會被搜尋引擎收錄，
一般情況下不會有不相關的人連進來。但如果之後想要多一層保護（例如同事的電腦
被別人借用，或想避免網址被誤傳出去），之後可以再加一個簡單的帳號密碼保護，
是很小的改動，有需要的話可以再請我加。
