#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
 ;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force

;***************************************My Daily Use Shortcuts************************************

;{START E-MAIL TITLE>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    ;{START DAILY DISBURSE

;EDY
:*:@ed::
FormatTime, xx,, MM-yyyy
InputBox, inputDate,dd, من فضلك اكتب يوم صرف العميل , 450, 230, 150
SendInput, صرف ميزه ايدي نجع حمادي يوم %inputDate%-%xx%
return

;MISR
:*:@mi::
FormatTime, xx,, MM-yyyy
InputBox, inputDate,dd, من فضلك اكتب يوم صرف العميل , 450, 230, 150
SendInput, صرف ميزه مصر متوسط نجع حمادي يوم %inputDate%-%xx%
return

;SPANISH
:*:@sp::
FormatTime, xx,, MM-yyyy
InputBox, inputDate,dd, من فضلك اكتب يوم صرف العميل , 450, 230, 150
SendInput, صرف ميزه تمكين السيدات نجع حمادي يوم %inputDate%-%xx%
return

;AHLY
:*:@ah::
FormatTime, xx,, MM-yyyy
InputBox, inputDate,dd, من فضلك اكتب يوم صرف العميل , 450, 230, 150
SendInput, صرف ميزه الاهلي نجع حمادي يوم %inputDate%-%xx%
return

;ALEX
:*:@x::
FormatTime, xx,, MM-yyyy
InputBox, inputDate,dd, من فضلك اكتب يوم صرف العميل , 450, 230, 150
SendInput, صرف ميزه بنك الاسكندرية نجع حمادي يوم %inputDate%-%xx%
return

;}END DAILY DISBURISE

    ;{START MONTHLY REPORTS
    ;Risk management
:*:@rc::
InputBox, month, MM, Please Enter Month Number!!
FormatTime, yy,, yyyy
SendInput, نتيجة الاستعلام في القوائم السلبية عملاء عن شهر %month%-%yy%
return

;Inform In Negative Lists Gr
:*:@rg::
InputBox, month, MM, Please Enter Month Number!!
FormatTime, yy,, yyyy
SendInput, نتيجة الاستعلام في القوائم السلبية ضامنين عن شهر %month%-%yy%
return

;Days Off Report
:*:@off::
FormatTime, yy,, yyyy
InputBox, month, MM, Please Enter Month Number!!
SendInput, اجازات نجع حمادي عن شهر %month%-%yy%
return

;MONTHLY REPORTS
:*:@mr::
FormatTime, yy,, yyyy
send, Monthly_Reports_%savedInput%-%yy%
return

;}END MONTHLY REPORTS
;}END E-MAIL TITLE>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
;{START FILE RENAME>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    ;{START EDY
::+ea::
FormatTime, xx,,dd-MM-yyyy
Send,  قروض معتمدة ايدي في ايدك شريحة [أ]-%xx%
return
::+eb::
FormatTime, xx,,dd-MM-yyyy
Send, قروض معتمدة ايدي في ايدك شريحة [ب]-%xx%
return
::+ec::
FormatTime, xx,,dd-MM-yyyy
Send, قروض معتمدة ايدي في ايدك شريحة [ج]-%xx%
return
::+e::
FormatTime, xx,,dd-MM-yyyy
send, معتمدة ايدي في ايدك-%xx%
return
::-e::
FormatTime, xx,,dd-MM-yyyy
Send, قروض مصدرة ايدي في ايدك-%xx%
return
;}END EDY
    ;{START AHLY
::+aa::
FormatTime, xx,,dd-MM-yyyy
SendInput, قروض معتمدة البنك الاهلي شريحة [أ]-%xx%
return

::+ab::
FormatTime, xx,,dd-MM-yyyy
SendInput, قروض معتمدة البنك الاهلي شريحة [ب]-%xx%
return

::+ac::
FormatTime, xx,,dd-MM-yyyy
SendInput, قروض معتمدة البنك الاهلي شريحة [ج]-%xx%
return
::+a::
FormatTime, xx,,dd-MM-yyyy
send, قروض معتمدة البنك الاهلي-%xx%
return
::-a::
FormatTime, xx,,dd-MM-yyyy
SendInput, قروض مصدرة البنك الاهلي-%xx%
return
    ;}END AHLY
    ;{START MISR
::+ma::
FormatTime, xx,,dd-MM-yyyy
Send, قروض معتمدة بنك مصر متوسط شريحة [أ]-%xx%
return

::+mb::
FormatTime, xx,,dd-MM-yyyy
Send, قروض معتمدة بنك مصر متوسط شريحة [ب]-%xx%
return

::+mc::
FormatTime, xx,,dd-MM-yyyy
Send, قروض معتمدة بنك مصر متوسط شريحة [ج]-%xx%
return
::+m::
FormatTime, xx,,dd-MM-yyyy
send, قروض معتمدة بنك مصر-%xx%
return
::-m::
FormatTime, xx,,dd-MM-yyyy
Send, قروض مصدرة بنك مصر-%xx%
return
    ;}END MISR
    ;{START SPANISH
::+sa::
FormatTime, xx,,dd-MM-yyyy
Send, قروض معتمدة مشروع الاسباني تمكين السيدات اجتماعيا واقتصاديا شريحة [أ]-%xx%
return

::+sb::
FormatTime, xx,,dd-MM-yyyy
Send, قروض معتمدة مشروع الاسباني تمكين السيدات اجتماعيا واقتصاديا شريحة [ب]-%xx%
return

::+sc::
FormatTime, xx,,dd-MM-yyyy
Send, قروض معتمدة مشروع الاسباني تمكين السيدات اجتماعيا واقتصاديا شريحة [ج]%xx%
return

::+s::
FormatTime, xx,,dd-MM-yyyy
send, قروض معتمدة مشروع الاسباني تمكين السيدات اجتماعيا واقتصادي-%xx%
return
::-s::
FormatTime, xx,,dd-MM-yyyy
SendInput, قروض مصدرة مشروع الاسباني تمكين السيدات اجتماعيا واقتصاديا-%xx%
return
    ;}END SPANISH
    ;{START TKAFUL
::+t::
FormatTime, xx,,dd-MM-yyyy
send, قروض معتمدة مشروع تعزيز القدرات المؤسسية-%xx%
return
::-t::
FormatTime, xx,,dd-MM-yyyy
SendInput, قروض مصدرة مشروع تعزيز القدرات المؤسسية-%xx%
return
    ;}END TKAFUL
    ;{START ALEXANDRIA
::+xa::
FormatTime, xx,,dd-MM-yyyy
send, قروض معتمدة مشروع بنك الاسكندرية شريحة [أ]-%xx%
return

::+xb::
FormatTime, xx,,dd-MM-yyyy
send, قروض معتمدة مشروع بنك الاسكندرية شريحة [ب]-%xx%
return

::+xc::
FormatTime, xx,,dd-MM-yyyy
send, قروض معتمدة مشروع بنك الاسكندرية شريحة [ج]-%xx%
return
::+x::
FormatTime, xx,,dd-MM-yyyy
send, معتمدة مشروع بنك الاسكندرية-%xx%
return
::-x::
FormatTime, xx,,dd-MM-yyyy
send, قروض مصدرة مشروع بنك الاسكندرية-%xx%
return
     ;}END ALEXANDRIA
;}END FILE RENAME>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
;{START AID FILES>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Ins::
FormatTime, xx,, dd-MM-yyyy
FormatTime, ll,, -MM-yyyy
FormatTime, tt,, hhmmss
InputBox, selected_option, select an option, Select Option: ,  , 450, 230, 315,370
Send,{F2}
SendInput, AID FILE_%xx%-%tt%[%selected_option%]
return
;}END AID FILES
;{START TIME STAMP>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
!x::
FormatTime, xx,, dd-MM-yyyy hh mm ss
SendInput, %xx%
return
    ;{START LAST REVIEWED
!+x::
FormatTime, xx,, dd-MM-yyyy HHmmss
SendInput, Last Reviewed On, %xx%
return
    ;}END LAST REVIEWED
    ;{START ACTIVATED WITH CURRENT DATE-TIME
+Ins::
FormatTime xx,,hh:mm:ss
FormatTime yy,,dd-MM-yyyy
SendInput Activated: %yy%  %xx%
return
    ;}END ACTIVATED WITH CURRENT DATE-TIME
    ;{START Loan Tracker Backup Title In Monthly Closing
:*:+backup::
FormatTime, xx,, dd-MM-yyyy
SendInput, Nag Backup After Closing %xx%
return
:*:-backup::
FormatTime, xx,, dd-MM-yyyy
SendInput, Nag Backup Before Closing %xx%
return
    ;}END Loan Tracker Backup Title In Monthly Closing
;}END TIME STAMP
;{START FAST ACCESS TO FOLDERS>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
;Contract Folder On Karam Pc [Alt]+[k]
!k::
run, \\karam\عقود
return

;Contract Folder On Karam (Current Month)[Alt][Shift][k]
!+k::
FormatTime, yy,, yyyy
FormatTime, mo,, MM

IfNotExist, \\karam\عقود\%yy%\%mo%\         ; Check if the folder exists
{
    FileCreateDir, \\karam\عقود\%yy%\%mo%\  ; Create the folder if it doesn't exist
}

Run, explorer.exe \\karam\عقود\%yy%\%mo%\   ; Open the folder in Explorer
return

;Clients ID Folder [Alt][i]
!i::
Run,D:\Clients_ID
sleep 300
Send, ^f
return

;Clients ID Folder (Current Month) [Alt][Shift][i]
!+i::
FormatTime, dd,, yyyy_MM_dd
Run,D:\Clients_ID\%dd%
return

;Open Docs Folder
!d::
run, D:\Committee\03 Disburise
return

!+d::
FormatTime, dd,, yyyy-MM-dd
run, D:\Committee\03 Disburise\%dd%-Disburise
return
;}END FAST ACCESS TO FOLDERS
;{START VOLUME CONROLE>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
+NumpadAdd::
send {Volume_Up }
return
+NumpadSub::
send {Volume_Down}
return
;}END VOLUME CONROLE
;{START COPY DATA FROM LOANTRACKER>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    ;{START COPY DATA FROM GAR
Pause::
Send {home}
Send +{End}
Send ^c
send {tab 3}
Sleep 600
Send ^c
send {tab 2}
Sleep 600
Send ^c
Sleep, 300
WinActivate, ahk_exe EXCEL.EXE

return
    ;}END COPY DATA FROM GAR
    ;{START COPY FROM CLIENT ADDRESS TO PROJECT ADDRESS IN LOAN TRACKER
:*:==::
MouseMove 1094, 570
MouseClick
Send {Home}
Send +{End}
Send ^c
MouseMove 1184, 679
MouseClick
Send ^v
return
;Add the start - end of colsing month
;-::
;Send 01062025
;return
;=::
;Send 30062025
;return


    ;}END COPY FROM CLIENT ADDRESS TO PROJECT ADDRESS IN LOAN TRACKER
;}END COPY GAR DATA FROM LOANTRACKER
;{START SAVED FOR LATER USE>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
; p::
; MouseGetPos, x, y ; Get current mouse position
; Tooltip, Mouse position: %x% - %y% ; Display mouse coordinates in a tooltip
;return
;~ ;Copy From Clients Screen For Contracts (PrintScreen)
;~ #IfWinActive ahk_class SunAwtFrame
;~ PrintScreen::
;~ send {tab}
;~ send ^c
;~ Sleep 1000
;~ send {tab}
;~ send ^c
;~ MouseMove 1122, 687
;~ MouseClick
;~ Send {Home}
;~ Send +{End}
;~ Sleep 1500
;~ Send ^c
;~ Send {Tab}
;~ Sleep 100
;~ Send {tab}
;~ sleep 500
;~ send ^c
;~ Sleep 600
;~ MouseMove 739, 722
;~ MouseClick
;~ return
;~ #IfWinActive
;}END SAVED FOR LATER USE