@echo off
setlocal

cd /D D:\wallpapers\32-9

:: Create numbered list of files in a temporary file
set "tempFile=%temp%\%~nx0_fileList_%time::=.%.txt"
dir /b /s /a-d %1 | findstr /n "^" >"%tempFile%"

:: Count the files
for /f %%N in ('type "%tempFile%" ^| find /c /v ""') do set cnt=%%N

call :openRandomFile

:: Delete the temp file
del "%tempFile%"

exit /b

:openRandomFile
set /a "randomNum=(%random% %% cnt) + 1"
for /f "tokens=1* delims=:" %%A in (
  'findstr "^%randomNum%:" "%tempFile%"'
) do start "" "C:\Users\colac\Documents\script wallpaper change\wallpaper-change.bat" %%B
exit /b