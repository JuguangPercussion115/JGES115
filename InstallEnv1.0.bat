@echo off
:: ==============================================================================
:: JGES_score 環境一鍵自動安裝與設定腳本 (.bat 企業穩定版)
:: ==============================================================================
chcp 65001 >nul
echo =========================================
echo    開始執行 JGES_score 環境安裝程序   
echo =========================================
echo.

:: 0. 設定專案目錄為腳本當前所在路徑
set "projectDir=%~dp0"
cd /d "%projectDir%"
echo [資訊] 當前專案工作目錄：%projectDir%

:: 1. 檢查並自動下載/安裝 Anaconda3 (若尚未安裝)
set "activatePath="
if exist "%USERPROFILE%\anaconda3\Scripts\activate.bat" set "activatePath=%USERPROFILE%\anaconda3\Scripts\activate.bat"
if exist "C:\ProgramData\Anaconda3\Scripts\activate.bat" set "activatePath=C:\ProgramData\Anaconda3\Scripts\activate.bat"
if exist "%USERPROFILE%\miniconda3\Scripts\activate.bat" set "activatePath=%USERPROFILE%\miniconda3\Scripts\activate.bat"
if exist "C:\ProgramData\miniconda3\Scripts\activate.bat" set "activatePath=C:\ProgramData\miniconda3\Scripts\activate.bat"

if defined activatePath goto :CondaEnvironmentSetup

echo [步驟 1/3] 未檢測到 Anaconda3，開始自動下載並靜默安裝...

set "installerUrl=https://anaconda.com"
set "installerPath=%TEMP%\Anaconda3_Installer.exe"

echo [資訊] 正在從官方網站使用 curl 下載 Anaconda3 安裝檔...
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
:: 移除 && 串接，改用 call 呼叫環境後獨立執行 conda 核心指令
call "%activatePath%" base
call conda create -n score python=3.13 -y

:: 3. 啟用 score 環境並安裝必要套件 (pandas, openpyxl, gitpython)
echo [步驟 3/3] 在 score 環境中安裝必要套件 (pandas, openpyxl, gitpython)...
:: 核心修正：明確切換到 score 環境，並分行強制執行安裝
call "%activatePath%" score
echo [資訊] 正在安裝 pandas 與 openpyxl...
call conda install -y pandas openpyxl
echo [資訊] 正在安裝 gitpython...
call pip install GitPython

:: 4. 完成提示
echo.
echo =========================================
echo    環境建置完成！                     
echo =========================================
echo 若要啟用 score 環境，請執行以下命令：
echo conda activate score
echo.
pause
