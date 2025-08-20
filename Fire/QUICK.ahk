#NoTrayIcon
Gui, Show, w740 h300, Control Panel
FormatTime, xx,, yyyy
;CREATE==============================================================================
Gui, Add, Button, x12 y19 w140 h20 gopenrevisionFile,           Open Revision TXT
Gui, Add, Button, x12 y39 w140 h20 grunrevisionreport,          Run Revision Report
Gui, Add, Button, x12 y59 w140 h20 gcreatDisburseFolder,        New Full Disburse Folder
Gui, Add, Button, x12 y79 w140 h20 gcreatDayoffFolder,          Creat Dayoff Folder
Gui, Add, Button, x12 y99 w140 h20 gcreatcommittee,             Initialize Committee Folder
Gui, Add, Button, x12 y119 w140 h20 gcreatnew-committee,        New Committee
Gui, Add, Button, x12 y139 w140 h20 gdisburse-edy,              Disburse Edy
Gui, Add, Button, x12 y159 w140 h20 gdisburse-ahly,             Disburse Ahly
Gui, Add, Button, x12 y179 w140 h20 gdisburse-alex,             Disburse Alex
Gui, Add, Button, x12 y199 w140 h20 gdisburse-misr,             Disburse Misr
Gui, Add, Button, x12 y219 w140 h20 gdisburse-spanish,          Disburse Spanish
;SLECTICTIVE FOLDERS==============================================================
Gui, Add, Button, x182 y19 w150 h20  gdayoffFolder,             Dayoff Folder %xx%
Gui, Add, Button, x182 y39 w150 h20  gselectDayoffFolder,       Dayoff Folder
Gui, Add, Button, x182 y59 w150 h20 gselectContractsFolder,     Contract Folder
Gui, Add, Button, x182 y79 w150 h20 gsearchiscore,              Search In Iscore
Gui, Add, Button, x182 y99 w150 h20 gsearchforiscore,           Search In Iscore History
Gui, Add, Button, x182 y119 w150 h20 gsearchforDayoffBalance,   DayOff Balance
;OPEN===============================================================================
Gui, Add, Button, x362 y19 w150 h20 gforms,                     Forms
Gui, Add, Button, x362 y39 w150 h20 gopeniscore,                Open Iscore_TXT
Gui, Add, Button, x362 y59 w150 h20 gopeniscorereport,          Run Inform-->Iscore Reprots
Gui, Add, Button, x362 y79 w150 h20 gopenInformFolder,          Iscour Folder
Gui, Add, Button, x362 y99 w150 h20 gopencommitteeFolder,       Committee Folder
Gui, Add, Button, x362 y119 w150 h20 gdisburseFolder,           Disburse Folder
Gui, Add, Button, x362 y139 w150 h20 gcurrentContractFolder,    Current Contract Folder
Gui, Add, Button, x362 y159 w150 h20 giscoursite,               Iscour Site
Gui, Add, Button, x362 y179 w150 h20 gemptyCards,               Empty Cards
Gui, Add, Button, x362 y199 w150 h20 gkaramSheet,               karam R Sheet

;DO====================================================================================
Gui, Add, Button, x542 y19 w160 h20 gtanfeez,                   Tanfeez
Gui, Add, Button, x542 y39 w160 h20 gdueInstallmentsReport,     Due Installments
Gui, Add, Button, x542 y59 w160 h20 gmoveInformFile,            Move Inform File
;=======================================================================================
Gui, Add, Text, x-12 y79 w200 h2 ,
Gui, Add, Text, x-12 y321 w170 h5 ,
Gui, Add, Text, x-12 y670 w170 h2 ,
Gui, Add, GroupBox, x537 y4 w180 h460 +Center, Do
Gui, Add, GroupBox, x7 y4 w160 h460 +Center, Create
Gui, Add, GroupBox, x357 y4 w170 h460 +Center, Open
Gui, Add, GroupBox, x177 y4 w170 h460 +Center, Selective Folders
return

;CREATE====================================================================================================================================

;Open Revision TXT File--------------------------------------------------------------

openrevisionFile:
Run, "D:\Committee\01 Inform\Reports\Gur_Report.txt"
Run, "D:\Committee\01 Inform\Reports\Client_Report.txt"

;FormatTime, ll,, yyyy-MM-dd
;FormatTime, xx,, yyyy             ;CURRENT YEAR
;FormatTime, tt,, MM               ;CURRENT MONTH
;FormatTime, ff,, HH mm ss
;IfNotExist, D:\COMMITTEE FORMES\Iscour_List\%xx%\%tt%\             ; Check if the folder exists
;{
;    FileCreateDir, D:\COMMITTEE FORMES\Iscour_List\%xx%\%tt%    ; Create the folder if it doesn't exist
;}
;FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\2023-11-01_4-.xlsx, D:\COMMITTEE FORMES\Iscour_List\%xx%\%tt%\%ll%_4-%ff%.xlsx
;run, D:\COMMITTEE FORMES\Iscour_List\%xx%\%tt%\%ll%_4-%ff%.xlsx
return

;Run Revision report--------------------------------

runrevisionreport:
Run,"D:\Fire\Py_Scripts\ISCORE_REPORT.py"
;FormatTime, ff,, HHmmss       ;CURRENT TIME
;FormatTime, CUR_DATE,,yyyy-MM-dd
;FileCreateDir, D:\Committee\01 Inform\Reports
;FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Iscour_Report.xlsx,  D:\Committee\01 Inform\Reports\%CUR_DATE% %ff%_Iscour_Report.xlsx
;run,D:\Committee\01 Inform\Reports\%CUR_DATE% %ff%_Iscour_Report.xlsx
return

;creat Full Disburse Folder--------------------------------------------------------------

creatDisburseFolder:
    FormatTime, CUR_DATE,,yyyy-MM-
	FormatTime, xx,,HHmmss
    InputBox, day, Enter Disburse Day !!!
	FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Ahly
	FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Edy
	FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Misr
	FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Spanish
	FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Alex
	FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Ahly.xlsm,         D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Ahly\%CUR_DATE%%day%_%xx%_Ahly.xlsm
	FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Edy.xlsm,          D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Edy\%CUR_DATE%%day%_%xx%_Edy.xlsm
	FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Misr_Med.xlsm,     D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Misr\%CUR_DATE%%day%_%xx%_Misr_Med.xlsm
	FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Spanish.xlsm,      D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Spanish\%CUR_DATE%%day%_%xx%%day%_Spanish.xlsm
	FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Alex.xlsm,         D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-Alex\%CUR_DATE%%day%_%xx%%day%_Alex.xlsm
return

;creat Disburse Edy Folder--------------------------------------------------------------
    disburse-edy:
    FormatTime, CUR_DATE,,yyyy-MM-
	FormatTime, xx,,HHmmss
    inputBox, day, Enter Disburse Day !!!
    FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Edy
    FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Edy.xlsm,  D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Edy\%CUR_DATE%%day%_%xx%_ايدي في ايدك.xlsm
    return

;creat Disburse Ahly Folder--------------------------------------------------------------
disburse-ahly:
    FormatTime, CUR_DATE,,yyyy-MM-
	FormatTime, xx,,HHmmss
    inputBox, day, Enter Disburse Day !!!
    FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Ahly
    FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Ahly.xlsm, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Ahly\%CUR_DATE%%day%_%xx%_البنك الاهلي.xlsm
    return
;creat Disburse Alex Folder--------------------------------------------------------------
disburse-alex:
    FormatTime, CUR_DATE,,yyyy-MM-
	FormatTime, xx,,HHmmss
    inputBox, day, Enter Disburse Day !!!
    FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Alex
    FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Alex.xlsm, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Alex\%CUR_DATE%%day%_%xx%_بنك الاسكندرية.xlsm
    return
;creat Disburse Misr Folder--------------------------------------------------------------
disburse-misr:
    FormatTime, CUR_DATE,,yyyy-MM-
	FormatTime, xx,,HHmmss
    inputBox, day, Enter Disburse Day !!!
    FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Misr
    FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Misr_Med.xlsm,  D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Misr\%CUR_DATE%%day%_%xx%-_بنك مصر_متوسط.xlsm
     FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Misr_Short.xlsm,  D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Misr\%CUR_DATE%%day%_%xx%_بنك مصر_قصير.xlsm
return
;creat Disburse Spanish Folder-----------------------------------------------------------
disburse-spanish:
    FormatTime, CUR_DATE,,yyyy-MM-
	FormatTime, xx,,HHmmss
    inputBox, day, Enter Disburse Day !!!
    FileCreateDir, D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Spanish
    FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Spanish.xlsm,  D:\Committee\03 Disburise\%CUR_DATE%%day%-Disburise\%CUR_DATE%%day%-%xx%-Spanish\%CUR_DATE%%day%-%xx%_الاسباني.xlsm
    return

;Creat Days Off Folder--------------------------------------------------------------

creatDayoffFolder:
InputBox, month, MM, Please Enter Month Number!!
FormatTime, xx,, yyyy             ;CURRENT YEAR

IfNotExist, D:\Days_Off\Day_Off_Archive\%xx%\%month%\             ; Check if the folder exists
{
    FileCreateDir,  D:\Days_Off\Day_Off_Archive\%xx%\%month%\    ; Create the folder if it doesn't exist
}
FileCopy, D:\Days_Off\Forms\Report.docx, D:\Days_Off\Day_Off_Archive\%xx%\%month%\%month%-%xx%.docx
FileCopy, D:\Days_Off\Forms\Attendance.xlsm, D:\Days_Off\Day_Off_Archive\%xx%\%month%\%month%.xlsm
FileCopy, D:\Days_Off\Forms\AttendanceDetails_Form.xlsm, D:\Days_Off\Day_Off_Archive\%xx%\%month%\Attendence-%month%-%xx%.xlsm
return

;initialize Main Committee Folder--------------------------------------------------------------

creatcommittee:
FormatTime, CUR_DATE,,yyyy-MM-dd
FormatTime, xx,,HHmmss
FileCreateDir, D:\Committee\01 Inform
FileCreateDir, D:\Committee\01 Inform\Reports
FileCreateDir, D:\Committee\03 Disburise
FileCreateDir, D:\Committee\05 Neg-List
FileCreateDir, D:\Committee\05 Neg-List\05-01 Clients
FileCreateDir, D:\Committee\05 Neg-List\05-02 Guarantors
FileCreateDir, D:\Committee\09 Scanner
FileCreateDir, D:\Committee\09 Scanner\Others
FileCreateDir, D:\Committee\09 Scanner\طلبات نجع حمادي للمراجعه_1
FileCreateDir, D:\Committee\10 Inbox
;===COPING FILES TO NEW CREATED FOLDERS---------------------->
FileCopy, D:\COMMITTEE FORMES\Neg_List_Form\Neg_List_Report.xlsx,           D:\Committee\05 Neg-List\%CUR_DATE% _Neg_List_Report.xlsx
FileCopy, D:\COMMITTEE FORMES\Neg_List_Form\Neg_List_Report.docx,           D:\Committee\05 Neg-List\%CUR_DATE% _Neg_List_Report.docx
FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\G-Tabel.xlsm,              D:\Committee\05 Neg-List\%CUR_DATE%_G-Tabel.xlsm
FileCopy, D:\COMMITTEE FORMES\cashed first 4 monthes.xlsm,                  D:\Committee\%CUR_DATE%_CASHED_FRIST_4_MONTHES.xlsm
FileCopy, D:\COMMITTEE FORMES\cashed_filter.xlsx,                           D:\Committee\%CUR_DATE%_CASHED_FILTER.xlsx
FileCopy, D:\COMMITTEE FORMES\2020-12-03_كروت_الصرف.xlsx,                   D:\Committee\%CUR_DATE%_INSTALMENTS.xlsx
FileAppend,, D:\Committee\01 Inform\Reports\Client_Report.txt
FileAppend,, D:\Committee\01 Inform\Reports\Gur_Report.txt
FileAppend,, D:\Committee\01 Inform\Reports\Iscore.txt


return

;Creat New Committee Folders==============================================

creatnew-committee:
InputBox, day, Enter Sub-Com Commite Minutes Day !!!
FormatTime, CUR_DATE,,yyyy-MM-
FormatTime, xx,,ddHHmmss
FileCreateDir, D:\Committee\Com-%xx%
FileCreateDir, D:\Committee\Com-%xx%\02 Com-Reports
FileCreateDir, D:\Committee\Com-%xx%\02 Com-Reports\Card_Holder_Sheet
FileCreateDir, D:\Committee\Com-%xx%\04 Reinforcement
FileCreateDir, D:\Committee\Com-%xx%\06 Helping-Files\Approved Loans-%CUR_DATE%%day%
FileCreateDir, D:\Committee\Com-%xx%\06 Helping-Files\Issued Loans-%CUR_DATE%%day%
FileCreateDir, D:\Committee\Com-%xx%\07 Clients-Contract-%xx%


;===COPING FILES TO NEW CREATED FOLDERS------------------------->

FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Ahly.xlsm,                                  D:\Committee\Com-%xx%\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%%day%_البنك الاهلي.xlsm
FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Edy.xlsm,                                   D:\Committee\Com-%xx%\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%%day%_ايدي في ايدك.xlsm
FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Misr_Mid.xlsm,                              D:\Committee\Com-%xx%\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%%day%_بنك مصر_متوسط.xlsm
FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Misr_short.xlsm,                            D:\Committee\Com-%xx%\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%%day%_بنك مصر_قصير.xlsm
FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_spanish.xlsm,                               D:\Committee\Com-%xx%\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%%day%_الاسباني.xlsm
FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Alex.xlsm,                                  D:\Committee\Com-%xx%\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%%day%_بنك الاسكندرية.xlsm
FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Ta3zez.docx,                                           D:\Committee\Com-%xx%\04 Reinforcement\%CUR_DATE%%day%_Ta3zez.docx
FileCopy, D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Ma7dar.xlsm,                                           D:\Committee\Com-%xx%\02 Com-Reports\%CUR_DATE%%day%_Com_Ma7dar.xlsm
FileCopy, D:\COMMITTEE FORMES\Contract\Update_Aug_2025\CLIENT_CONTRACT_AUG_2025_UPDATE.xltm,  D:\Committee\Com-%xx%\07 Clients-Contract-%xx%\%CUR_DATE%%day%-%xx%-Com Clients Contract_Aug-2025-update.xltm
return

;SELECTECTIVE FOLDERS===========================================================================================================================
;Open Days Off Folder--------------------------------------------------------------

dayoffFolder:
FormatTime, xx,, yyyy  ;CURRENT YEAR
InputBox, month, Open Folder MM, in Year %xx% Which Month You Wanna Open!!!
 Run,  D:\Days_Off\Day_Off_Archive\%xx%\%month%\
return

;Select Days Off Folder--------------------------------------------------------------

selectDayoffFolder:
FormatTime, xx,, yyyy
FormatTime, yy,, MM
InputBox, yr, Enter Year !!!
Sleep, 200
InputBox, mo, Enter Month !!!
Run, D:\Days_Off\Day_Off_Archive\%yr%\%mo%
return

;Select Contract Folder-------------------------------------------------------------

selectContractsFolder:
FormatTime, xx,, yyyy
FormatTime, yy,, MM
InputBox, yr, Enter Year !!!
Sleep, 200
InputBox, mo, Enter Month !!!
Run, \\karam\عقود\%yr%\%mo%
return
;==Search For Iscour
searchiscore:
run, "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
sleep,1000
send, ^+f
send, +{Tab}
send, !{Down}
send, {Down 5}
Send, {Enter}
sleep, 1500
Send, +{tab}
Send, +{tab}
SendInput, D:\Committee\{Numpad0}{Numpad1} Inform
Send, {enter}
sleep, 1500
Send, {Tab}
send, +{Tab}
send, +{Tab}
send, +{Tab}
send, {Enter}
send, {Tab 3}
return
;===Search For Isoor File================================================================

searchforiscore:
Sleep, 200
InputBox, YY, Enter Year, from 11-2023 !!!
Sleep,200
InputBox, MM, Enter Month, from 11-2023 !!!
sleep, 200
run, "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
sleep,1300
send, ^+f
sleep, 200
send, +{Tab}
sleep, 200
send, !{Down}
sleep, 200
send, {Down 5}
sleep, 200
Send, {Enter}
sleep, 1500
Send, +{tab}
sleep, 200
Send, +{tab}
Sleep, 1300
SendInput, D:\@_ME_Archive\%YY%\%MM%\Committee\{Numpad0}{Numpad1} Inform
Sleep, 1300
Send, {enter}
sleep, 1500
Send, {Tab}
sleep, 200
send, +{Tab}
sleep, 200
send, +{Tab}
sleep, 200
send, +{Tab}
sleep, 200
send, {Enter}
sleep, 200
send, {Tab 3}
return

;===Open Dayoff  Balance-----------------------------

searchforDayoffBalance:
InputBox, YY, YYYY, Enter Year !!!
Run,D:\Days_Off\Days_Off_balance\%YY%
run, D:\Days_Off\Days_Off_balance\%YY%\Days off_balance_Naghammadi_%YY%.xlsm
return


;OPEN============================================================================================

;Forms--------------------------------------------------------------

Forms:
Run,"D:\Project_X\index.html"
return

;Open Txt file to--------------------------------------------------------------

openiscore:
Run,"D:\Committee\01 Inform\Reports\Iscore.txt"
;SetWorkingDir, D:\Committee\05 Neg-List\05-01 Clients
;Loop, Files, *.xlsx
;{
;   FileGetAttrib, Attributes, %A_LoopFileFullPath%
;  If !(Attributes & FILE_ATTRIBUTE_DIRECTORY)
; {
;    Run, %A_LoopFileFullPath%
;   Break
;}
;}
return

;Run iscore script--------------------------------------------------------------

openiscorereport:
Run,"D:\Fire\Py_Scripts\ISCORE_CLIENT_&_GUARANTOR.py"
;SetWorkingDir, D:\Committee\05 Neg-List\05-02 Guarantors ; Replace with the actual folder path
;Loop, Files, *.xlsx ; Adjust the file extension if needed (e.g., *.xls for older Excel files)
;{
    ;FileGetAttrib, Attributes, %A_LoopFileFullPath%
    ;If !(Attributes & FILE_ATTRIBUTE_DIRECTORY) ; Check if it's not a directory
    ;{
        ; First non-directory file found, open it and exit the loop
        ;Run, %A_LoopFileFullPath%
        ;Break
    ;}
;}
return

;Open Inform Folder--------------------------------------------------------------

openInformFolder:
FormatTime, ll,, yyyy-MM-dd
FormatTime, xx,, yyyy             ;CURRENT YEAR
FormatTime, tt,, MM               ;CURRENT MONTH
Run,D:Iscour_List\%xx%\%tt%
return

;Commitee Folder--------------------------------------------------------------

opencommitteeFolder:
Run,"D:\Committee"
return

;Disburse Folder--------------------------------------------------------------

disburseFolder:
run, D:\Committee\03 Disburise
return

;Open Current Contract Folder

currentContractFolder:
FormatTime, mo,, MM
FormatTime, yy,, yyyy
Run, \\karam\عقود\%yy%\%mo%
return

;I Scour Site===

iscoursite:
run, Chrome.exe "https://newinquiry.emff-eg.com/Pages/Security/Login"
Sleep, 800
WinWaitActive, EMFF-Home - Google Chrome
sleep, 800
Send, {Tab}
sleep, 100
SendInput, AUEEDQ{Numpad4}
sleep, 100
Send, {Tab}
sleep, 100
SendInput, AUEED{Numpad4}
sleep, 100
Send, {Tab}
send, {Space}
Run, "D:\Fire\Move inform file auto to com.ahk"
return

;Open Empty Cards Sheet==============

emptyCards:
run, D:\COMMITTEE FORMES\Miza_Cards\Miza_Daiy_Disburse\Miza_Recevied_Report\Empty_Cards.xlsm
return

;Karam Review Sheet--------------------------------------------------------------
karamSheet:
Run, "D:\Docs\Karam_REview.xlsm"
return

;DO==============================================================================================================

;Tanfeez--------------------------------------------------------------

tanfeez:
WinActivate, LoanTracker 3.1
send, {alt}
send, {Left}
send, {Down 4}
Send, {Left}
Send, {Left}
send, {Down 4}
MouseMove, 1006, 258
MouseClick,
WinWaitActive, LoanTracker 3.1
MouseMove, 1423, 657
Sleep 400
MouseClick,
Sleep 400
MouseMove, 1230, 658
Sleep 400
MouseClick,
Sleep 400
MouseMove, 637, 659
Sleep 1000
MouseClick,
Sleep 400
Send, {Space}
Sleep 400
Send, {Space}
Sleep, 500
MouseMove, 835, 111
Sleep, 200
MouseClick
WinActivate, ahk_exe AutoHotkeyU32.exe
Send, {Down}
return

;Due Installments Report--------------------------------------------------------------

dueInstallmentsReport:
FormatTime xx,,MMyyyy
WinActivate, LoanTracker 3.1
MouseMove 1381, 460
MouseClick
sleep 1500
WinWaitActive LoanTracker 3.1
SendInput 01%xx%
Send {tab}
SendInput 28%xx%
Send {tab}
MouseMove 957, 546
sleep 1000
MouseClick
sleep 1000
MouseMove 894, 440
MouseClick
sleep 4000


;MouseMove, 797, 374
;Sleep 100
;MouseClick
return

;Move Inform File--------------------------------------------------------------

moveInformFile:
FormatTime, ll,, yyyy-MM-dd   ;CURRENT DATE
FormatTime, ff,, HH mm ss       ;CURRENT TIME

; Define source and destination folders
sourceFolder := "C:\Users\Peter\Downloads"

; Define file prefix and new file name
filePrefix := "AUEED"
newFileName := "%ii%_IScour_%ff%.pdf"  ; Set to "" to keep original name

; Loop through files in the source folder
Loop, Files, %sourceFolder%\%filePrefix%*.pdf
{
    ; Construct full source and destination paths
    sourceFilePath := A_LoopFileFullPath
    destinationFilePath := destinationFolder . "\" . (newFileName ? newFileName : A_LoopFileName)

    ; Move and rename the file
    FileMove, %sourceFilePath%, D:\Committee\01 Inform\%ll%_IScour_%ff%.pdf
    Run,D:\Committee\01 Inform\%ll%_IScour_%ff%.pdf
    return

    ; Check if the move operation succeeded
    ;if (ErrorLevel)
   ; {
   ;     MsgBox, Error moving file: %sourceFilePath%
   ; }
   ; else
   ; {
   ;     MsgBox, File moved successfully to: D:\Committee\01 Inform\
  ;  }

    ; Exit after the first file (optional if you only want to move one file)
  ;  break
}

; Optional: Exit the script when done
;ExitApp












return
