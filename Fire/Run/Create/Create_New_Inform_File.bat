echo off
::Setting Variables
set curTime=%time:~0,2% %time:~3,2% %time:~6,2%
set CUR_DATE=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%
set CUR_Mo=%DATE:~4,2%
set CUR_Year=%DATE:~10,4%
::Create New folder Named With Current Month
if not exist "D:\COMMITTEE FORMES\Iscour_List\%CUR_Year%\%CUR_Mo%" mkdir "D:\COMMITTEE FORMES\Iscour_List\%CUR_Year%\%CUR_Mo%"
::Copy New Inform File Named With Current date and time to the new folder and start it
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\2023-11-01_4-.xlsx" "D:\COMMITTEE FORMES\Iscour_List\%CUR_Year%\%CUR_Mo%\%CUR_DATE%_4-%curTime%.xlsx"
start excel "D:\COMMITTEE FORMES\Iscour_List\%CUR_Year%\%CUR_Mo%\%CUR_DATE%_4-%curTime%.xlsx"
::Start Inquiry Site To inform


Exit /b
