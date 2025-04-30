@echo off
setlocal enabledelayedexpansion

rem Set UTF-8 code page
chcp 65001 > nul

rem Configure colors
color 0A

rem Set window title
title DOC 檔案轉換工具

:menu
cls
echo ==================================
echo     Word 文件轉換工具
echo ==================================
echo.
echo  --- DOC 轉 DOCX 選項 ---
echo  1. 標準轉換 (保留原始檔案)
echo  2. 轉換並刪除原始檔案
echo  3. 不建立備份的轉換
echo  4. 不建立備份並刪除原始檔案
echo.
echo  --- DOCX 轉 TXT 選項 ---
echo  5. 將 DOCX 轉換為 TXT (保留原始檔案)
echo  6. 將 DOCX 轉換為 TXT 並刪除原始檔案
echo.
echo  --- 其他選項 ---
echo  7. 自定義參數
echo  8. 離開
echo.
echo ==================================
echo.

set /p choice=請輸入選項 (1-8): 

if "%choice%"=="1" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ConvertDocToDocx.ps1"
    goto end
)
if "%choice%"=="2" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ConvertDocToDocx.ps1" -DeleteOriginal
    goto end
)
if "%choice%"=="3" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ConvertDocToDocx.ps1" -Backup:$false
    goto end
)
if "%choice%"=="4" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ConvertDocToDocx.ps1" -DeleteOriginal -Backup:$false
    goto end
)
if "%choice%"=="5" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ConvertDocxToTxt.ps1"
    goto end
)
if "%choice%"=="6" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ConvertDocxToTxt.ps1" -DeleteOriginal
    goto end
)
if "%choice%"=="7" (
    cls
    echo 自定義參數
    echo =================
    echo.
    echo 請選擇要執行的轉換:
    echo 1. DOC 轉 DOCX
    echo 2. DOCX 轉 TXT
    set /p conv_type=請選擇轉換類型 (1-2): 
    
    set script_file="%~dp0ConvertDocToDocx.ps1"
    if "!conv_type!"=="2" (
        set script_file="%~dp0ConvertDocxToTxt.ps1"
    )
    
    echo.
    set /p source_dir=請輸入來源目錄 (留空使用當前目錄): 
    
    echo.
    set /p del_original=是否刪除原始檔案? (y/n): 
    
    echo.
    set /p create_backup=是否建立備份? (y/n): 
    
    echo.
    set /p silent_mode=是否使用安靜模式? (y/n): 
    
    set command=powershell -ExecutionPolicy Bypass -File !script_file!
    
    if not "!source_dir!"=="" (
        set command=!command! -SourceDir "!source_dir!"
    )
    
    if /i "!del_original!"=="y" (
        set command=!command! -DeleteOriginal
    )
    
    if /i "!create_backup!"=="n" (
        set command=!command! -Backup:$false
    )
    
    if /i "!silent_mode!"=="y" (
        set command=!command! -Silent
    )
    
    echo.
    echo 執行命令: !command!
    echo.
    echo 按任意鍵執行...
    pause > nul
    
    !command!
    goto end
)
if "%choice%"=="8" (
    goto exit
)

echo.
echo 無效的選項，請重新選擇。
echo.
pause
goto menu

:end
echo.
echo 轉換完成!
echo.
echo 按任意鍵返回主選單...
pause > nul
goto menu

:exit
echo.
echo 感謝使用，再見!
echo.
pause
exit 