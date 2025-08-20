@echo off
REM Get the current date in YYYY-MM-DD format.
set CUR_DATE=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%
REM CREATE FOLDER IN COMMITTEE WITH CURRENT DATE 
rem Ahly
mkdir "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Ahly"
REM COPY EDY EXCEL SHEET FROM COMMITTEE FORMES TO CURRENT FOLDER WITH CURRENT DATE
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Ahly.xlsm" "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Ahly\%CUR_DATE%-Ahly.xlsm"
start EXCEL "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Ahly\%CUR_DATE%-Ahly.xlsm"

rem Spanish

REM CREATE FOLDER IN COMMITTEE WITH CURRENT DATE 
mkdir "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Spanish"
REM COPY EDY EXCEL SHEET FROM COMMITTEE FORMES TO CURRENT FOLDER WITH CURRENT DATE
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Spanish.xlsm" "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Spanish\%CUR_DATE%-Spanish.xlsm"
start EXCEL "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Spanish\%CUR_DATE%-Spanish.xlsm"

rem EDY

REM CREATE FOLDER IN COMMITTEE WITH CURRENT DATE 
mkdir "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Edy"
REM COPY EDY EXCEL SHEET FROM COMMITTEE FORMES TO CURRENT FOLDER WITH CURRENT DATE
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Edy.xlsm" "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Edy\%CUR_DATE%-Edy.xlsm"
start EXCEL "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Edy\%CUR_DATE%-Edy.xlsm"

rem Misr-mid

REM CREATE FOLDER IN COMMITTEE WITH CURRENT DATE 
mkdir "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Misr"
REM COPY EDY EXCEL SHEET FROM COMMITTEE FORMES TO CURRENT FOLDER WITH CURRENT DATE
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Misr_Med.xlsm" "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Misr\%CUR_DATE%-Misr_Med.xlsm"
start EXCEL "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Misr\%CUR_DATE%-Misr_Med.xlsm"

rem Misr-short

REM CREATE FOLDER IN COMMITTEE WITH CURRENT DATE 
::mkdir "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Misr"
REM COPY EDY EXCEL SHEET FROM COMMITTEE FORMES TO CURRENT FOLDER WITH CURRENT DATE
::copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Misr_Short.xlsm" "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Misr\%CUR_DATE%-Misr_Short.xlsm"
::start EXCEL "D:\Committee\03 Disburise\%CUR_DATE%-Disburise\%CUR_DATE%-Misr\%CUR_DATE%-Misr_Short.xlsm"

