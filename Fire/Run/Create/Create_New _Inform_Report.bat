@echo off
set CUR_DATE=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%
set CUR_Mo=%DATE:~10,4%-%DATE:~4,2%

copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Iscour_Report.xlsm" "D:\Committee\01 Inform\Reports\%CUR_DATE%_Iscour_Report.xlsm"
start Excel "D:\Committee\01 Inform\Reports\%CUR_DATE%_Iscour_Report.xlsm"