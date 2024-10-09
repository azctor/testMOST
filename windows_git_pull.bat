@echo off
REM 進入當前目錄下的虛擬環境
call windowsenv\Scripts\activate

REM 執行當前目錄下的 git_pull.py 腳本，並傳入當前目錄作為儲存庫路徑
python git_pull.py .

REM 停留在畫面，讓你看到結果
pause
