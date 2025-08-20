@echo off
@echo off
set curTime=%time:~0,2% %time:~3,2%
set CUR_DATE=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%
set CUR_Mo=%DATE:~4,2%
set CUR_Year=%DATE:~10,4%
move "C:\Users\Peter\Downloads\Copy of total_report.pdf" "D:\Committee\01 Inform\%CUR_DATE%_IScour_%curTime%.pdf"
move "C:\Users\Peter\Downloads\Copy of total_report.xlsx" "D:\Committee\01 Inform\%CUR_DATE%_IScour_%curTime%.xlsx"
start Acrobat.exe "D:\Committee\01 Inform\%CUR_DATE%_IScour_%curTime%.pdf"

