@echo off

rem Get the temporary files folder path
set temp_folder=%temp%

rem Delete all files and folders in the temporary files folder
for /f "delims=" %%f in ('dir /b %temp_folder%') do (
  del /q /f "%temp_folder%\%%f"
)

rem Delete all empty folders in the temporary files folder
rd /s /q "%temp_folder%"