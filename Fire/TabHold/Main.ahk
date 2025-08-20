#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Include <TapHoldManager>
#SingleInstance Force

thm := new TapHoldManager()
thm.Add("ctrl", Func("Func1"))

return
Func1(ishold, tabs, state){
if (ishold=0) & (tabs=1) & (state){
	SendInput, خط سير
}
if (ishold=0) & (tabs=2) & (state){
SendInput, خصم ربع يوم تأخير بصمة الحضور
}
if (ishold=1) & (tabs=1) & (state){
SendInput, خصم نصف يوم اغفال بصمة الانصراف
}




}
