@echo off
netsh advfirewall firewall set rule group="Network discovery" new enable=yes
netsh advfirewall firewall set rule group="File and Printer sharing" new enable=yes
netsh advfirewall firewall set rule group="Public folder sharing" new enable=yes
REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" /v "RequirePasswordProtectedSharing" /t REG_DWORD /d 0 /f


echo Changing SSDP Discovery Service startup type to Manual...
sc config SSDPSRV start= demand
echo SSDP Discovery Service startup type has been set to Manual.

echo Changing UPnP Device Host service startup type to Manual...
sc config upnphost start= demand
echo UPnP Device Host service startup type has been set to Manual.

echo Starting Function Discovery Provider Host service...
sc start fdPHost
echo Function Discovery Provider Host service has been started.

echo Starting Function Discovery Resource Publication service...
sc start FDResPub
echo Function Discovery Resource Publication service has been started.

echo Creating registry key RpcAuthnLevelPrivacyEnabled...
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print" /v RpcAuthnLevelPrivacyEnabled /t REG_DWORD /d 0 
echo Registry key RpcAuthnLevelPrivacyEnabled has been created.

echo Restarting Print Spooler service...
net stop spooler
net start spooler
echo Print Spooler service has been restarted.
pause