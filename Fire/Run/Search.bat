@echo off
setlocal enabledelayedexpansion

echo Enter search string:
set /p searchString=

echo Searching for files containing: %searchString%

for /r %%f in (*.*) do (
  findstr /i "%searchString%" "%%f" >nul && echo File: %%f
)

pause
