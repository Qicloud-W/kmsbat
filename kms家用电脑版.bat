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
echo.     -- KMS 激活 windows 和 office 便捷脚本 --
echo.     -- 此脚本由 凌梦云 提供支持 --
echo.
echo. --[1]--激活 windows 系统（Windows 7、8、10、2008、2012、2016）
echo. --[2]--激活 office 软件（office 2010、2013、2016、office365）
echo. --[3]--退出脚本
echo.
echo.
choice /c 123 /n /m "请选择【1-3】："

echo. %errorlevel%
if %errorlevel% == 1 goto set_1
if %errorlevel% == 2 goto set_2
if %errorlevel% == 3 goto end


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
echo. --[1]-- 激活 windows 10       --[6]-- 激活 windows 2016
echo. --[2]-- 激活 windows 8.1      --[7]-- 激活 windows 2012R2
echo. --[3]-- 激活 windows 8        --[8]-- 激活 windows 2012
echo. --[4]-- 激活 windows 7        --[9]-- 激活 windows 2008R2
echo. --[5]-- 激活 windows vista    --[f]-- 激活 windows 2008
echo.
choice /c 123456789f /n /m "请选择【1-f】："

echo. %errorlevel%
if %errorlevel% == 1 set windowsv=win10
if %errorlevel% == 2 set windowsv=win8.1
if %errorlevel% == 3 set windowsv=win8
if %errorlevel% == 4 set windowsv=win7
if %errorlevel% == 5 set windowsv=vista

if %errorlevel% == 6 set windowsv=2016
if %errorlevel% == 7 set windowsv=2012R2
if %errorlevel% == 8 set windowsv=2012
if %errorlevel% == 9 set windowsv=2008R2
if %errorlevel% == 10 set windowsv=2008
goto :EOF


:win10
cscript //Nologo %windir%\system32\slmgr.vbs /ipk MH37W-N47XK-V7XM9-C7227-GCQG9 1>nul 2>nul
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

:win8.1
cscript //Nologo %windir%\system32\slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk HMCNV-VVBFX-7HMBH-CTY9B-B4FXY 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk TT4HM-HN7YT-62K67-RGRQJ-JFFXW 1>nul 2>nul
goto :EOF

:win8
cscript //Nologo %windir%\system32\slmgr.vbs /ipk NG4HW-VH26C-733KW-K6F98-J8CK4 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk XCVCF-2NXM9-723PB-MHCB7-2RYQQ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 32JNW-9KQ84-P47T8-D8GGY-CWCK7 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk JMNMF-RHW7P-DMY6X-RF3DR-X2BQT 1>nul 2>nul
goto :EOF

:win7
cscript //Nologo %windir%\system32\slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk MRPKT-YTG23-K7D7T-X2JMM-QY7MG 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk W82YF-2Q76Y-63HXB-FGJG9-GF7QX 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk YDRBP-3D83W-TY26F-D46B2-XCKRJ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk C29WB-22CC8-VJ326-GHFJW-H9DH4 1>nul 2>nul
goto :EOF

:vista
cscript //Nologo %windir%\system32\slmgr.vbs /ipk YFKBB-PQJJV-G996G-VWGXY-2V3X8 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk HMBQG-8H2RH-C77VX-27R82-VMQBT 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk VKK3X-68KWM-X2YGT-QR4M6-4BWMV 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk VTC42-BM838-43QHV-84HX6-XJXKV 1>nul 2>nul
goto :EOF

:2016
cscript //Nologo %windir%\system32\slmgr.vbs /ipk CB7KF-BWN84-R7R2Y-793K2-8XDDG 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk JCKRF-N37P4-C2D82-9YXRT-4M63B 1>nul 2>nul
goto :EOF

:2012R2
cscript //Nologo %windir%\system32\slmgr.vbs /ipk D2N9P-3P6X9-2R39C-7RTCD-MDVJX 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk KNC87-3J2TX-XB4WP-VCPJV-M4FWM 1>nul 2>nul
goto :EOF

:2012
cscript //Nologo %windir%\system32\slmgr.vbs /ipk BN3D2-R7TKB-3YPBD-8DRP2-27GG4 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 8N2M2-HWPGY-7PGT9-HGDD8-GVGGY 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 2WN2H-YGCQR-KFX6K-CD6TF-84YXQ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 4K36P-JN4VD-GDC6V-KDT89-DYFKP 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk XC9B7-NBPP2-83J2H-RHMBY-92BT4 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk HM7DN-YVMH3-46JC3-XYTG7-CYQJJ 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 48HP8-DN98B-MYWDG-T2DCC-8W83P 1>nul 2>nul
goto :EOF

:2008R2
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 6TPJF-RBVHG-WBW2R-86QPH-6RTM4 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk TT8MH-CG224-D3D7Q-498W2-9QCTX 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk YC6KT-GKW9T-YTKYR-T4X34-R7VHC 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 489J6-VHDMP-X63PK-3K798-CPX3Y 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk 74YFP-3QFB3-KQT8W-PMXWJ-7M648 1>nul 2>nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk GT63C-RJFQ3-4GMB6-BRFB9-CB83V 1>nul 2>nul
goto :EOF

:2008
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


::激活office套件
:set_2
set pathv=
set officev=
set url=
call :setoffice
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

::配置office安装目录
cls
echo.
echo.
echo. --【请输入正确的office安装目录】
echo.
echo. --默认为：C:\%pathv%\Microsoft Office\%officev%
echo.
set/p url=--默认直接按回车：
if not defined url set url=C:\%pathv%\Microsoft Office\%officev%
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


:setoffice
cls
echo.
echo.
echo.             -- 请选择 Office 版本 --
echo.     -- 此脚本由 凌梦云  提供支持 --
echo.
echo. --[1]-- 激活 Office 2010 32位   --[2]-- 激活 Office 2010 62位
echo. --[3]-- 激活 Office 2013 32位   --[4]-- 激活 Office 2013 62位
echo. --[5]-- 激活 Office 2016 32位   --[6]-- 激活 Office 2016 62位
echo. --[7]-- 激活 Office 365  32位   --[8]-- 激活 Office 365  62位
echo.
echo.
choice /c 12345678 /n /m "请选择【1-8】："

echo. %errorlevel%
if %errorlevel% == 1 set officev=office14&set pathv=Program Files (x86)
if %errorlevel% == 2 set officev=office14&set pathv=Program Files

if %errorlevel% == 3 set officev=office15&set pathv=Program Files (x86)
if %errorlevel% == 4 set officev=office15&set pathv=Program Files

if %errorlevel% == 5 set officev=office16&set pathv=Program Files (x86)
if %errorlevel% == 6 set officev=office16&set pathv=Program Files

if %errorlevel% == 7 set officev=office16&set pathv=Program Files (x86)
if %errorlevel% == 8 set officev=office16&set pathv=Program Files
goto :EOF

:office14
for /f %%x in ('dir /b ..\root\Licenses14\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" 1>nul 2>nul
for /f %%x in ('dir /b ..\root\Licenses14\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" 1>nul 2>nul
cscript ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB 1>nul 2>nul
cscript ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM 1>nul 2>nul
cscript ospp.vbs /inpkey:D6QFG-VBYP2-XQHM7-J97RH-VVRCK 1>nul 2>nul
goto :EOF

:office15
for /f %%x in ('dir /b ..\root\Licenses15\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" 1>nul 2>nul
for /f %%x in ('dir /b ..\root\Licenses15\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" 1>nul 2>nul
cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT 1>nul 2>nul
cscript ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4 1>nul 2>nul
goto :EOF

:office16
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" 1>nul 2>nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" 1>nul 2>nul
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 1>nul 2>nul
cscript ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM 1>nul 2>nul
goto :EOF


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


::路径错误
:pathn
cls
echo.
echo.
echo.
echo.
echo.
echo. ***** 错误：Office安装路径错误 请重新输入 *****
echo.
echo.
echo. --操作已完成！请联系凌梦云 客服 获取 最新线路
pause>nul
goto begin


::关闭脚本
:end
exit