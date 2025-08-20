#SingleInstance force
#Persistent
#include Lib\AutoHotInterception.ahk


AHI := new AutoHotInterception()

keyboardId := AHI.GetKeyboardId(0x03F0, 0x0024)
AHI.SubscribeKeyboard(keyboardId, true, Func("KeyEvent"))


return

KeyEvent(code, state){
	ToolTip % "Keyboard Key - Code: " code ", State: " state

;{===Change Language (Backspace)==========
	if (state=1) & (code=35)
	{
	Send, {Alt Down}{Shift Down}{Alt Up}{Shift Up}
	}
;}0x413C, 0x2105
;{===Copy[C]=================================
if (state=1) & (code=46)
	{
	Send, ^c
	}
	;}
;{===Past[V]==================================
if (state=1) & (code=47)
	{
	Send, ^v
	}
;}
;{===Past as Values[F1]=======================
if (state=1) & (code=59)
	{
	Send, {Ctrl  Down}{Alt Down}{v Down} {Ctrl  Up}{Alt Up}{v Up}
	Send {Down 2}{Space}
	Send {Enter}
	}
	;}
;{===Groove Music[F2]========================
if (state=1) & (code=60)
	{
	send, #{s}
	sleep, 200
	SendInput Groove
	Sleep, 700
	Send, {Enter}
	sleep, 700
	MouseMove 115, 687
	Sleep, 700
	WinActivate ahk_exe ApplicationFrameHost.exe
	Sleep 700
	WinWaitActive ahk_exe ApplicationFrameHost.exe
	Sleep 700
	MouseClick
	Sleep 700
	MouseMove 652, 259
	sleep 700
	MouseClick
	}
	;}
;{===Task Manger(F3)=========================
if (state=1) & (code=61)
{
Send, ^+{Esc}
}
;}
;{===Get Mouse Position[Num Lock]===========
if (state=1) & (code=325)
	{
	MouseGetPos, xpos, ypos
MsgBox, The cursor is at X%xpos% Y%ypos%.
	}
	;}
;{===Jump to Desktops (Left)(Right)(Up)=======
if (state=1) & (code=333)
{
Send, {Ctrl Down}#{Right Down}{Ctrl Up}#{Right Up}
}

if (state=1) & (code=331)
{
Send, {Ctrl Down}#{Left Down}{Ctrl Up}#{Left Up}
}

if (state=1) & (code=328)
{
Send, #{Tab Down}#{Tab up}
}
;}
;{===Explorer(Down arrow)===================
if (state=1) & (code=336)
{
Send, #{e Down}#{e up}
}
;}
;{===Close(Escape)===========================
if (state=1) & (code=1)
{
Send, {Alt Down}{F4 Down}{Alt Up}{F4 Up}
}
;}
;{===Jumb between windows [Menu Key]======
if (state=1) & (code=79) & (code=80)
{
Send, {Alt Down}{Tab Down}{Alt Up}{Tab Up}
}
;}

if (state=1) & (code=349)
{
Send, {Alt Down}{Tab Down}{Alt Up}{Tab Up}
}























}


return
;Esc::
;ExitApp