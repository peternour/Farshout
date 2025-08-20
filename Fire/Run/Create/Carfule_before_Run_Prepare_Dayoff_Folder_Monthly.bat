@echo off
set CUR_DATE=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%
set CUR_MO=%DATE:~4,2%-%DATE:~10,4%
set CUR_MOnth=%DATE:~4,2%
set CUR_Year=%DATE:~10,4%
mkdir "D:\Days_Off\Day_Off_Archive\%CUR_Year%\%CUR_MOnth%"
copy "D:\Days_Off\Forms\Report.docx" "D:\Days_Off\Day_Off_Archive\%CUR_Year%\%CUR_MOnth%\%CUR_DATE%_Days_Off_Report.docx"
copy "D:\Days_Off\Forms\Attendance.xlsm" "D:\Days_Off\Attendence Record\%CUR_Year%\%CUR_MOnth%.xlsm"
copy "C:\Users\Peter\Desktop\Working_On\AttendanceDetails_Empty.xlsm" "C:\Users\Peter\Desktop\Working_On\%CUR_MO%_Days_Off\%CUR_DATE%.xlsm"
start "Word" "D:\Days_Off\Day_Off_Archive\%CUR_Year%\%CUR_MOnth%\%CUR_DATE%_Days_Off_Report.docx"
