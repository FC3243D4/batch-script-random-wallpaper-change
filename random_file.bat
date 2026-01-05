@echo off
setlocal

cd /D YOUR-WALLPAPERS-FOLDER

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
) do start "" "REPO-FOLDER\wallpaper-change.bat" %%B
exit /b