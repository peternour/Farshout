#Persistent
#include D:\Fire\ahk\AutoHotInterception\Lib\AutoHotInterception.ahk
AHI := new AutoHotInterception()
keyboardId := AHI.GetKeyboardId(0x03F0, 0x0024)
AHI.SubscribeKeyboard(keyboardId, true, Func("KeyEvent"))
return

KeyEvent(code, state)
{
;ToolTip % "Keyboard Key - Code: " code ", State: " state

;get mouse position>>>numlock<<<=======================================================
if (state=1) & (code=325)
{
MouseGetPos, xpos, ypos
MsgBox, The cursor is at X%xpos% Y%ypos%.
}

;switch between desktops>>>arrow right<<< =======================================================
if (state=1) & (code=333)
{
Send, {Ctrl Down}#{Right Down}{Ctrl Up}#{Right Up}
}

;switch back between desktops>>>arrow left<<< =======================================================
if (state=1) & (code=331)
{
Send, {Ctrl Down}#{Left Down}{Ctrl Up}#{Left Up}
}

;show all desktops >>>arrow up<<< =======================================================
if (state=1) & (code=328)
{
Send, #{Tab Down}#{Tab up}
}

;use for new client in google sheet to mark the client green>>>big enter<<< =======================================================
if (state=1) & (code=28)
{
Send, New
send {tab 8}
Send, New
Send, {Shift Down}{tab 6}{Shift Up}
}

;copy client data from google sheet to past it in loan tracker>>>space<<< =======================================================
if (state=1) & (code=57)
{
WinActivate, ahk_class Chrome_WidgetWin_1
send, ^{c}
sleep 600
Send {tab}
send, ^{c}
sleep 600
Send {tab}
send, ^{c}
sleep 600
Send {tab 6}
send, ^{c}
sleep 600
Send {tab}
send, ^{c}
sleep 600
Send {tab}
send, ^{c}
sleep 600
Send {tab}
Sleep, 600
send, ^{c}
sleep, 600
Send, {Shift Down}{tab 11}{Shift Up}
WinActivate ahk_class SunAwtFrame
}

;print curent page in any pdf file>>>num 0<<< =======================================================
if (state=1) & (code=82)
{
WinActivate, ahk_class AcrobatSDIWindow
send, ^p
Send, {Tab 2}{u}{Tab 2}{Enter}
}

;open my loan tracker user>>>num +<<< =======================================================
if (state=1) & (code=78)
{
Run, C:\Users\Peter\Desktop\ABA_erp.bat
WinWaitActive, LoanTracker 3.1
Sleep, 1000
SendInput, {Numpad4}{Numpad5}
sleep, 200
send {tab}
sleep, 200
SendInput, megaboxine
sleep, 200
send {tab}
send {tab}
send {tab}
sleep, 200
send {Space}
}

;loan tracker tanfeez all>>>num -<<< =======================================================
if (state=1) & (code=74)
{
WinActivate, LoanTracker 3.1
send, {alt}
Sleep, 100
send, {Left}
Sleep, 100
send, {Down 4}
Sleep, 100
Send, {Left}
Sleep, 100
Send, {Left}
Sleep, 100
send, {Down 4}
Sleep, 100
MouseMove, 1006, 258
MouseClick,
WinWaitActive, LoanTracker 3.1
MouseMove, 1423, 657
Sleep 500
MouseClick,
Sleep 500
MouseMove, 1230, 658
Sleep 500
MouseClick,
Sleep 500
MouseMove, 637, 659
Sleep 1000
MouseClick,
Sleep 500
Send, {Space}
Sleep 500
Send, {Space}
Sleep, 500
MouseMove, 835, 111
Sleep, 200
MouseClick
Sleep,2000
;====tanfeez eradat====
WinActivate, LoanTracker 3.1
send, {alt}
Sleep, 300
send, {Left}
Sleep, 300
send, {Down 4}
Sleep, 300
send, {Left}
Sleep, 300
send, {Down}
Sleep, 300
Send, {Left}
Sleep, 300
send, {Down}
Sleep, 500
MouseMove, 980, 198
Sleep, 600
MouseClick,
Sleep, 600
MouseMove, 1406, 601
Sleep, 600
MouseClick,
Sleep, 600
MouseMove, 737, 594
Sleep, 600
MouseClick,
Send, {Space}
Sleep 500
Send, {Space}
MouseMove, 834, 107
Sleep, 600
MouseClick
}

;loan tracker client inform>>>num enter<<< =======================================================
if (state=1) & (code=284)
{
WinActivate ahk_exe jp2launcher.exe
InputBox, code, Enter code, Please enter your client code:
sleep, 200
MouseMove, 926, 110
sleep, 300
MouseClick
sleep, 300
MouseMove, 1109, 106
sleep, 300
MouseClick
sleep, 300
MouseMove, 630, 514
sleep, 300
MouseClick
sleep, 300
WinActivate ahk_exe jp2launcher.exe
SendInput {Numpad0}{Numpad0}{Numpad4}
Send {Enter}
Send,%code%
MouseMove, 1109, 106
sleep, 300
MouseClick
;sleep, 300
;MouseMove, 898, 110
;sleep, 300
;MouseClick
}

;scan id as photo>>>num 9<<< =======================================================
if (state=1) & (code=73)
{
WinActivate ahk_exe CNQMMAIN.EXE
MouseMove 570, 367
MouseClick
}

;search for iscour in 01 inform folder>>>num .<<< =======================================================
if (state=1) & (code=83)
{
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
}

;loan tracker due installments >>>num *<<< =======================================================
if (state=1) & (code=55)
{
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
MouseMove 796, 306
MouseClick
sleep 4000
WinActivate, LoanTracker 3.1
}

;open contract>>>Home<<< =======================================================
if (state=1) & (code=327)
{
	run, C:\Users\Peter\Desktop\Contract.BAT
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
if (state=1) & (code=)
{
}
;>>><<< =======================================================
}
return