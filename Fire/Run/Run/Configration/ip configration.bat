:: This batch file checks for network connection problems
:: and saves the output to a .txt file.
ECHO OFF
:: View network connection details
ipconfig /all >> ip_configration.txt
:: Check if Google.com is reachable
ping google.com >>ip_configration.txt
:: Run a traceroute to check the route to Google.com
tracert google.com >> ip_configration.txt