@echo off
title �޸�Զ������˿�
set port=23388
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\Wds\rdpwd\Tds\tcp" /v PortNumber /t REG_SZ /d %port% /f 1>nul
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" /v PortNumber /t REG_SZ /d %port% /f 1>nul
netsh advfirewall firewall add rule name="Զ������" dir=in protocol=tcp localport=%port% action=allow
set /p=allend.