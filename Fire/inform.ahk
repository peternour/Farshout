NumpadAdd::
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

return