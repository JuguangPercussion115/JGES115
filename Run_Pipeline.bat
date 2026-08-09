@echo off
rem ==========================================
rem LINK TO GITHUB DESKTOP FOR GIT CORE
rem ==========================================
for /d %%i in ("%LOCALAPPDATA%\GitHubDesktop\app-*") do set "PATH=%PATH%;%%i\resources\app\git\cmd"

rem ==========================================
rem ACTIVATE CONDA ENVIRONMENT
rem ==========================================
call "%USERPROFILE%\anaconda3\Scripts\activate.bat" score

rem ==========================================
rem RUN CALCULATION AND AUTO deployment
rem ==========================================
"%USERPROFILE%\anaconda3\envs\score\python.exe" _Champion.py

pause
