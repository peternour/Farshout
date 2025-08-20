(New-object -Comobject Shell.Application).Windows() | ForEach-object {$_.quit()}
Get-Process| Where-object{$_.MainWindowTitle -ne ""}|Stop-Process