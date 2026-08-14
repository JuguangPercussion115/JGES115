# 🏆 莒光打擊樂團 - 選拔成績看板自動化系統

本專案為莒光打擊樂團選拔賽事之全自動成績處理系統。系統會自動從 Google 雲端硬碟下載各年級最新成績與加減分紀錄、完成積分與排名運算，並自動生成與發布線上成績看板。

---

## 🌐 線上成績看板網址

👉 **[點此開啟 🏆 全國賽_選拔排名表](https://juguangpercussion115.github.io/JGES115/html/)**

---

## 🚀 日常成績更新說明（行政人員操作）

當老師在 Google 雲端硬碟更新學生試奏分數或加減分表格後，可以透過以下方式讓線上網頁同步更新：

### 方案 A：☁️ 雲端全自動更新（最推薦，手機/電腦皆可操作）

1. **定時自動跑**：系統已設定**每天晚上 18:00（台灣時間）**自動連線雲端硬碟更新網頁，平時無需任何人工操作。
2. **隨時手動一鍵更新**：
   1. 開啟本專案的 [GitHub 頁面](https://github.com/JuguangPercussion115/JGES115)。
   2. 點擊上方的 **「Actions」** 分頁。
   3. 在左側選單點選 **「自動更新成績看板」**。
   4. 點擊右側的 **「Run workflow」** 綠色按鈕 ➔ 再按一次 **「Run workflow」**。
   5. 等待約 30 秒出現綠色勾勾（🟢），線上網頁即更新完成！

---

### 方案 B：💻 本地電腦執行（備用方案）

若要在您自己的 Windows 電腦上手動執行：
1. 雙擊執行資料夾中的 **`Run_Pipeline.bat`**。
2. 視窗會自動下載最新 Excel、計算成績、產生網頁，並推送到 GitHub 網頁上。

> 💡 *第一次在新的電腦執行時，請先雙擊 `InstallEnv.bat` 安裝環境。*

---

## ⚙️ 系統設定與檔案說明 (`ini/competition.ini`)

若需要新增年級、修改曲目或排除學生，請開啟 `ini/competition.ini` 檔案修改：

```ini
[DEFAULT]
產出html檔名 = html/index.html
產出html標題 = 🏆🏆🏆 2027全國賽_選拔排名表 🏆🏆🏆

; 設定參與計算的年級（請用半角逗號隔開）
grade = 115,116,117

; 各年級與加減分檔案之 Google 雲端硬碟下載連結
adj_url = [https://drive.google.com/uc?export=download&id=](https://drive.google.com/uc?export=download&id=)...
115_url = [https://drive.google.com/uc?export=download&id=](https://drive.google.com/uc?export=download&id=)...
116_url = [https://drive.google.com/uc?export=download&id=](https://drive.google.com/uc?export=download&id=)...
117_url = [https://drive.google.com/uc?export=download&id=](https://drive.google.com/uc?export=download&id=)...

; 評比曲目名稱（須與 Excel 欄位一致）
曲目資訊 = Do Re Mi 之歌-2,Do Re Mi 之歌-3,...

; 不參賽學生（格式：班級_姓名，多個請用半角逗號隔開）
remove_student = 406_陳在緯
--
📌 首次使用設定（只需設定一次）
步驟 0：準備 GitHub 帳號與數位金鑰（Token）
白話解釋：這就像是在 GitHub 註冊帳號，並申請一把「自動更新網頁的通行證鑰匙」（Personal Access Token），讓電腦程式可以幫你把最新網頁自動放到網路上。

做法：

註冊並登入 GitHub 網站。

到個人設定（Settings -> Developer settings -> Personal access tokens）申請一組金鑰 Token。

權限請勾選 repo（完全控制專案）。

申請完成後，複製那一長串金鑰代碼（請先妥善儲存）。

步驟 1：安裝電腦執行環境
白話解釋：讓你的電腦安裝好計算成績與產生網頁所需的工具包。

做法：在資料夾中找到 InstallEnv.bat 檔案，直接「連點滑鼠左鍵兩下」執行，等待跳出的視窗跑完即可。

步驟 2 & 3：設定成績檔案網址與連線資訊
白話解釋：告訴系統去哪裡下載最新的成績 Excel 檔，以及把製作好的網頁放到哪個網站。

做法：

點進 ini 資料夾，用「記事本」開啟 competition.ini 設定檔。

貼上 Excel 下載連結：確認各年級（如 115_url、116_url、117_url）的雲端硬碟連結是否正確。

貼上網站與金鑰資訊：

github_url：填入你的 GitHub 專案網址（例如：https://github.com/JuguangPercussion115/JGES115）

github_token：填入步驟 0 複製的那一長串金鑰字串（例如：ghp_xxxxxxxxx）。

🚀 日常更新成績（每次更新成績時執行）

步驟 4：一鍵執行「成績計算」與「自動上傳網頁」
白話解釋：這是平時最常用的動作！只要雲端 Excel 的成績有更新，點這個檔案就會全自動完成所有工作。

做法：

直接「連點滑鼠左鍵兩下」執行 Run_Pipeline.bat。

電腦會自動下載 Excel 檔 ➔ 計算總分與排名 ➔ 產生網頁 ➔ 自動同步上傳到網站。

🌐 查看成果
步驟 5：檢查產生的網頁
本機備份網頁：會在你的電腦 html 資料夾內自動產生 index.html 檔案。

公開線上網頁（發布給家長/學生看）：
https://juguangpercussion115.github.io/JGES115/html/

End
--
--
For engineering reference :
0. Apply GitHub account and create personal access token
Personal access token should have full control of private repositories

1. Install enviroment package
Run "InstallEnv.bat"

2. Update URL and config in ".\ini\competition.ini"

3. Check gitgub repo URL and personal access token file in .ini
ex.
github_url = https://github.com/JuguangPercussion115/JGES115
github_token = ini/PAT.json

4. Double click "Run_Pipeline.bat" or run with windows cmd for detail.

5. Output HTML file to .\html folder and auto piblish to Github as public web site once success
HTML URL on Github : https://juguangpercussion115.github.io/JGES115/html/
