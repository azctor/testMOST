@echo off
:: 切換到目前 .bat 檔案所在的目錄
cd /d "%~dp0"

:: 啟動虛擬環境，假設虛擬環境位於當前目錄下的 `myenv` 資料夾
call myenv\Scripts\activate.bat

:: 執行 Python 腳本來更新專案，假設 git_pull.py 在當前目錄中
python git_pull.py .

:: 關閉虛擬環境
deactivate

:: 提示更新完成
echo 更新完成
pause
