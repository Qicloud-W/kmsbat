::原文来源 https://v0v.bid
@echo off
::获取运行路径
cd /d %~dp0


::获取管理员权限
%1 start "" mshta vbscript:createobject("shell.application").shellexecute("""%~0""","::",,"runas",1)(window.close)&exit


::配置BAT对话框样式
title --凌梦云KMS便捷脚本--
MODE con: COLS=70 lines=14
color 0a


::菜单
:begin
cls
echo.
echo.
echo.     -- KMS windows Servre 便捷激活脚本 --
echo.     -- 此脚本由 凌梦云 提供支持 --
echo.
echo. --[1]--激活 windows 系统（Windows 10、2019、2016、2012、2008）
echo. --[2]--退出脚本
echo.
echo.
choice /c 123 /n /m "请选择【1-3】："

echo. %errorlevel%
if %errorlevel% == 1 goto set_1
if %errorlevel% == 2 goto end


::激活windows系统
:set_1
set windowsv=
call :setwindows
::配置KMS服务器地址
cls
echo.
echo.
echo. --【请输入KMS服务器地址】
echo.
echo. --默认激活服务器为：kmsok.vps.si
echo.
set/p kms1=--默认直接按回车：
if not defined kms1 set kms1=kmsok.vps.si
echo.
echo. --配置成功，按【任意键】继续。
pause>nul

::配置KMS激活密钥
cls
echo.
echo.
echo. --【请输入KMS激活密钥】
echo.
echo. --默认KMS激活密钥为：自动匹配
echo.
set/p winkey=--默认直接按回车：
if not defined winkey set winkey=none
echo.
echo. --配置成功，按【任意键】继续。
pause>nul

::激活
cls
echo.
echo.
call :checkkms1
echo.
echo. 您选择的操作系统为：%windowsv%
echo.
echo. --正在激活，正在匹配密钥，请耐心等待.....
		call :%windowsv%
		cscript //Nologo %windir%\system32\slmgr.vbs /ipk %winkey% >nul
		cscript //Nologo %windir%\system32\slmgr.vbs /skms %kms1% >nul
		cscript //Nologo %windir%\system32\slmgr.vbs /ato >nul
		cscript //Nologo %windir%\system32\slmgr.vbs /xpr | find /i "激活" >nul && ( echo. & echo. ***** 操作系统 激活成功 ***** & echo. ) || ( echo. & echo. ***** 操作系统 激活失败 ***** & echo. )
echo.
echo. --操作已完成。如激活失败请联系凌梦云 客服
pause>nul
goto begin


:setwindows
cls
echo.
echo.
echo.             -- 请选择 windows 版本 --
echo.     -- 此脚本由 凌梦云  提供支持 --
echo.
echo. --[1]-- 激活 windows 10
echo. --[2]-- 激活 Windows 2019
echo. --[3]-- 激活 Windows 2016
echo. --[4]-- 激活 Windows 2012
echo. --[5]-- 激活 Windows 2008
echo.
choice /c 123456789f /n /m "请选择【1-5】："

echo. %errorlevel%
if %errorlevel% == 1 set windowsv=win10
if %errorlevel% == 2 set windowsv=win2019
if %errorlevel% == 3 set windowsv=win2016
if %errorlevel% == 4 set windowsv=win2012
if %errorlevel% == 5 set windowsv=win2008

goto :EOF


:win10
cscript //Nologo %windir%\system32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk MH37W-N47XK-V7XM9-C7227-GCQG9 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk WNMTR-4C88C-JK8YV-HQ7T2-76DF9 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 2F77B-TNFGY-69QQF-B8YKP-D69TJ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk QFFDN-GRT3P-VKWWX-X7T3R-8B639 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 44RPN-FTY23-9VTTB-MP9BX-T84FV 1>nul 2>nul
goto :EOF

:win2019
cscript //Nologo %windir%\system32\slmgr.vbs /ipk WMDGN-G9PQG-XVVXX-R3X43-63DFG 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk N69G4-B89J2-4G8F4-WWYCC-J464C 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk WVDHN-86M7X-466P6-VHXV7-YY726 1>nul 2>nul
goto :EOF

:win2016
cscript //Nologo %windir%\system32\slmgr.vbs /ipk CB7KF-BWN84-R7R2Y-793K2-8XDDG 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk JCKRF-N37P4-C2D82-9YXRT-4M63B 1>nul 2>nul
goto :EOF

:win2012
cscript //Nologo %windir%\system32\slmgr.vbs /ipk W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk KNC87-3J2TX-XB4WP-VCPJV-M4FWM 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk BN3D2-R7TKB-3YPBD-8DRP2-27GG4 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 8N2M2-HWPGY-7PGT9-HGDD8-GVGGY 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 2WN2H-YGCQR-KFX6K-CD6TF-84YXQ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk D2N9P-3P6X9-2R39C-7RTCD-MDVJX 1>nul 2>nul
goto :EOF

:win2008
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 74YFP-3QFB3-KQT8W-PMXWJ-7M648 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 6TPJF-RBVHG-WBW2R-86QPH-6RTM4 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk TT8MH-CG224-D3D7Q-498W2-9QCTX 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk YC6KT-GKW9T-YTKYR-T4X34-R7VHC 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 489J6-VHDMP-X63PK-3K798-CPX3Y 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 74YFP-3QFB3-KQT8W-PMXWJ-7M648 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk GT63C-RJFQ3-4GMB6-BRFB9-CB83V 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk WYR28-R7TFJ-3X2YQ-YCY4H-M249D 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk TM24T-X9RMF-VWXK6-X8JC9-BFGM2 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk YQGMW-MPWTJ-34KDK-48M3W-X4Q6V 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 39BXF-X8Q23-P2WWT-38T2F-G3FPG 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk RCTX3-KWVHP-BR6TB-RB6DM-6X7HP 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 7M67G-PC374-GR742-YH8V4-TCBY3 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 22XQ2-VRXRG-P8D42-K34TD-G3QQC 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 4DWFP-JF3DJ-B7DTH-78FJB-PDRHK 1>nul 2>nul
goto :EOF

:2016
cscript //Nologo %windir%\system32\slmgr.vbs /ipk CB7KF-BWN84-R7R2Y-793K2-8XDDG 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk JCKRF-N37P4-C2D82-9YXRT-4M63B 1>nul 2>nul
goto :EOF

::配置KMS服务器地址
cls
echo.
echo.
echo. --【请输入KMS服务器地址】
echo.
echo. --默认激活服务器为：kmsok.vps.si
echo.
set/p kms2=--默认直接按回车：
if not defined kms2 set kms2=kmsok.vps.si
echo.
echo. --配置成功，按【任意键】继续。
pause>nul

::配置KMS激活密钥
cls
echo.
echo.
echo. --【请输入KMS激活密钥】
echo.
echo. --默认KMS激活密钥为：自动匹配
echo.
set/p officekey=--默认直接按回车：
if not defined officekey set officekey=none
echo.
echo. --配置成功，按【任意键】继续。
pause>nul


::激活
cls
echo.
echo.
call :checkkms2
echo.
echo. --正在激活，请稍后.....
echo. %url%\ospp.vbs
if exist "%url%\ospp.vbs" ( echo. & echo. 检测安装路径配置正确，耐心等待激活..... ) else ( goto pathn )
		cd /d %url%
		call :%officev%
		cscript //nologo ospp.vbs /inpkey:%officekey% >nul
		cscript //nologo ospp.vbs /sethst:%kms2% >nul
		cscript //nologo ospp.vbs /act | find /i "successful" >nul && ( echo. & echo. ***** 激活成功 ***** & echo. ) || ( echo. & echo. ***** 激活失败 ***** & echo. )
echo.
echo. --操作已完成。如激活失败请联系凌梦云客服
echo.
pause>nul
goto begin



::检测 KMS服务器
:checkkms1
cls
echo.
echo.
echo. 正在检查激活服务器：%kms1% 请稍后.....
ping %kms1% | find /i "来自" >nul && ( goto :EOF ) || ( goto fail )


::检测 KMS服务器
:checkkms2
cls
echo.
echo.
echo. 正在检查激活服务器：%kms2% 请稍后.....
ping %kms2% | find /i "来自" >nul && ( goto :EOF ) || ( goto fail )


::检查失败 
:fail
cls
echo.
echo.
echo.
echo.
echo.
echo. ***** 错误：KMS激活服务器无效 *****
echo.
echo.
echo. --操作已完成！请联系凌梦云 客服 获取 最新线路
pause>nul
goto begin

::关闭脚本
:end
exit