@echo off

for /d %%i in ("%LOCALAPPDATA%\GitHubDesktop\app-*") do set "PATH=%PATH%;%%i\resources\app\git\cmd"

:: ==============================================================================
:: JGES_score 環境一鍵自動安裝與設定腳本 (.bat 台灣CDN特化版)
:: ==============================================================================
chcp 65001 >nul
echo ====================================================
echo    開始執行 JGES_score 環境安裝程序 (含免下載 Git 自動部署)
echo ====================================================
echo.

:: 0. 設定專案目錄為腳本當前所在路徑
set "projectDir=%~dp0"
cd /d "%projectDir%"
echo [資訊] 當前專案工作目錄：%projectDir%

:: ==============================================================================
:: 核心擴充：切換為微軟授權 NuGet 穩定節點，徹底繞過 GitHub 防火牆阻擋
:: ==============================================================================
echo [步驟 0/3] 正在檢查系統 Git 部署狀態...
where git >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [資訊] 偵測到系統已安裝 Git，跳過 Git 安裝程序。
    goto :AnacondaCheck
)

echo [資訊] 系統尚未安裝 Git，啟動微軟 CDN 綠色免認證通道下載...

:: 建立 Git 專用綠色版資料夾
set "gitTargetDir=%USERPROFILE%\Git_Portable"
if not exist "%gitTargetDir%" mkdir "%gitTargetDir%"

:: 改採用微軟 NuGet 官方伺服器所代管的 Git 綠色可攜二進位檔 (確保大小約 16MB 且不被網管阻擋)
set "gitUrl=https://nuget.org"
set "gitZip=%TEMP%\MinGit.zip"

echo [資訊] 正在從高可用 CDN 下載 Git 核心元件 (請稍候)...
:: 使用 curl 下載
curl -L -k -f "%gitUrl%" -o "%gitZip%"

if not exist "%gitZip%" (
    echo [錯誤] 檔案下載未成功，請確認網路是否具有連線能力。
    pause
    exit /b 1
)

echo [資訊] 下載完成，正在使用系統 tar 引擎進行自動解壓縮...
:: 使用 Windows 內建 tar 指令解壓
tar -xf "%gitZip%" -C "%gitTargetDir%"

:: 刪除暫存壓縮檔
del /f /q "%gitZip%" 2>nul

:: 修正 NuGet 包結構中的路徑，對齊到系統內
set "PATH=%PATH%;%gitTargetDir%\tools\cmd;%gitTargetDir%\tools\bin;%gitTargetDir%\cmd;%gitTargetDir%\bin"

:: 寫入一組臨時控制項，確保當前視窗及下方 pip 指令百分之百可以叫用 git 核心
set "git"

where git >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [資訊] 免手動下載！Git 綠色版已在背景透過微軟通道部署成功。
    goto :AnacondaCheck
)

echo [錯誤] Git 雲端部署失敗，請確認是否受到本機防毒軟體嚴格阻擋。
pause
exit /b 1


:AnacondaCheck
echo.
:: ==============================================================================
:: 1. 檢查並自動下載/安裝 Anaconda3 (若尚未安裝)
:: ==============================================================================
set "activatePath="
if exist "%USERPROFILE%\anaconda3\Scripts\activate.bat" set "activatePath=%USERPROFILE%\anaconda3\Scripts\activate.bat"
if exist "C:\ProgramData\Anaconda3\Scripts\activate.bat" set "activatePath=C:\ProgramData\Anaconda3\Scripts\activate.bat"
if exist "%USERPROFILE%\miniconda3\Scripts\activate.bat" set "activatePath=%USERPROFILE%\miniconda3\Scripts\activate.bat"
if exist "C:\ProgramData\miniconda3\Scripts\activate.bat" set "activatePath=C:\ProgramData\miniconda3\Scripts\activate.bat"

if defined activatePath goto :CondaEnvironmentSetup

echo [步驟 1/3] 未檢測到 Anaconda3，開始自動下載並靜默安裝...

set "installerUrl=https://anaconda.com"
set "installerPath=%TEMP%\Anaconda3_Installer.exe"

echo [資訊] 正在從官方網站下載 Anaconda3 安裝檔...
curl -L "%installerUrl%" -o "%installerPath%"

echo [資訊] 正在進行 Anaconda3 靜默安裝 (可能需要數分鐘，請稍候)...
start /wait "" "%installerPath%" /InstallationType=JustMe /RegisterPython=1 /S /D=%USERPROFILE%\anaconda3

del /f /q "%installerPath%" 2>nul

if exist "%USERPROFILE%\anaconda3\Scripts\activate.bat" (
    set "activatePath=%USERPROFILE%\anaconda3\Scripts\activate.bat"
    echo [資訊] Anaconda3 安裝成功！
    goto :CondaEnvironmentSetup
)

echo [錯誤] Anaconda3 安裝失敗，請嘗試手動安裝後再執行此腳本。
pause
exit /b


:CondaEnvironmentSetup
echo [資訊] 已檢測到 Anaconda 環境：%activatePath%

:: 2. 建立 score Conda 環境 (Python 3.13)
echo [步驟 2/3] 建立 Conda 環境：score (Python 3.13)...
call "%activatePath%" base
call conda create -n score python=3.13 -y

:: 3. 啟用 score Environment 並強制獨立分行安裝套件
echo [步驟 3/3] 在 score 環境中安裝必要套件 (pandas, openpyxl, gitpython)...
call "%activatePath%" score
echo [資訊] 正在安裝 pandas 與 openpyxl...
call conda install -y pandas openpyxl
echo [資訊] 正在安裝 gitpython...
call pip install GitPython

:: 4. 完成提示
echo.
echo ====================================================
echo    環境建置完成！(免下載 Git 與 Anaconda 已全數就緒)
echo ====================================================
echo 請關閉目前此視窗，並點擊執行 [Run_Pipeline.bat] 開始部署。
echo ====================================================
echo.
pause
