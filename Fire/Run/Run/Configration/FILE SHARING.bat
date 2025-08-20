@ECHO OFF
REM Check if file and printer sharing is enabled.
net share

REM If file and printer sharing is not enabled, enable it.
if errorlevel 1 (
  netsh firewall set service type "file and printer sharing" enabled
  echo File and printer sharing has been enabled.
)

REM Check if the printer spooler service is running.
sc query spooler

REM If the printer spooler service is not running, start it.
if errorlevel 1 (
  sc start spooler
  echo The printer spooler service has been started.
)

REM Check if the NetBIOS over TCP/IP service is running.
sc query netbios

REM If the NetBIOS over TCP/IP service is not running, start it.
if errorlevel 1 (
  sc start netbios
  echo The NetBIOS over TCP/IP service has been started.
)

REM Check if the Remote Procedure Call (RPC) service is running.
sc query rpcss

REM If the Remote Procedure Call (RPC) service is not running, start it.
if errorlevel 1 (
  sc start rpcss
  echo The Remote Procedure Call (RPC) service has been started.
)

REM Check if the Server service is running.
sc query server

REM If the Server service is not running, start it.
if errorlevel 1 (
  sc start server
  echo The Server service has been started.
)

REM Check if the Workstation service is running.
sc query workstation

REM If the Workstation service is not running, start it.
if errorlevel 1 (
  sc start workstation
  echo The Workstation service has been started.
)

REM Check if the Windows Firewall service is running.
sc query firewall

REM If the Windows Firewall service is not running, start it.
if errorlevel 1 (
  sc start firewall
  echo The Windows Firewall service has been started.
)

REM Check if the Windows Firewall is configured to allow file and printer sharing.
netsh advfirewall firewall show rule name="File and Printer Sharing (In)"

REM If the Windows Firewall is not configured to allow file and printer sharing, enable it.
if errorlevel 1 (
  netsh advfirewall firewall set rule name="File and Printer Sharing (In)" new enable=Yes
  echo The Windows Firewall has been configured to allow file and printer sharing.
)

REM Check if the Windows Firewall is configured to allow NetBIOS over TCP/IP.
netsh advfirewall firewall show rule name="NetBIOS over TCP/IP (In)"

REM If the Windows Firewall is not configured to allow NetBIOS over TCP/IP, enable it.
if errorlevel 1 (
  netsh advfirewall firewall set rule name="NetBIOS over TCP/IP (In)" new enable=Yes
  echo The Windows Firewall has been configured to allow NetBIOS over TCP/IP.
)

REM Check if the Windows Firewall is configured to allow Remote Procedure Call (RPC).
netsh advfirewall firewall show rule name="Remote Procedure Call (RPC) (In)"

REM If the Windows Firewall is not configured to allow Remote Procedure Call (RPC), enable it.
if errorlevel 1 (
  netsh advfirewall firewall set rule name="Remote Procedure Call (RPC) (In)" new enable=Yes
  echo The Windows Firewall has been configured to allow Remote Procedure Call (RPC).
)

REM Check if the Windows Firewall is configured to allow Server service.
netsh advfirewall firewall show rule name="Server (In)"

REM If the Windows Firewall is not configured to allow Server service, enable it.
if errorlevel 1 (
  netsh advfirewall firewall set rule name="Server (In)" new enable=Yes
  echo The Windows Firewall has been configured to allow Server service.
)

REM Check if the Windows Firewall is configured to allow Workstation service.
netsh advfirewall firewall show rule name="Workstation (In)"

REM If the Windows Firewall is not configured to allow Workstation service, enable it.
if errorlevel 1 (
  netsh advfirewall firewall set rule name="Workstation (In)" new enable=Yes
  echo The Windows Firewall has been configured to allow Workstation service.
)

REM All services have been started and configured.
echo File and printer sharing is now enabled.
