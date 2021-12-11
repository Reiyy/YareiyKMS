@echo off

::设置编码为GBK
chcp 936

::取运行路径
setlocal EnableDelayedExpansion
cd /d "%~dp0"

::外观
title Windows KMS激活工具 Ver:4.0 2021/10/20 BY:夜棂依Yareiy
MODE con: COLS=78 lines=21
color 3f

::检测是否为WindowsXP系统
ver | find "5." > nul && goto :xptooff

::获取管理员权限
%1 start "" mshta vbscript:createobject("shell.application").shellexecute("""%~0""","::",,"runas",1)(window.close)&exit

::KMS服务器
set kmsroot=kms.yareiy.com

::主菜单
:begin
cls
echo.
echo.
echo.
echo.	###### Windows/Office KMS激活工具 ######
echo.	           Ver:4.0 2021/10/20
echo.	            BY:夜棂依Yareiy
echo.	  https://github.com/yareiy722/YareiyKMS
echo.
echo. [1]-激活 Windows 支持 Vista/7/8/8.1/10/2008/2008R2/2012/2012R2/2016
echo. [2]-激活 Office 支持 Office 2010/2013/2016/Office365
echo. [3]-卸载 Windows KMS激活
echo. [4]-卸载 Office KMS激活
echo. [5]-检查激活状态
echo. [6]-Office系列Retail版转Vol版
echo.
echo.
choice /c 123456 /n /m "请按下要选择的选项对应的数字【1-6】："

echo. %errorlevel%
if %errorlevel% == 1 goto windows
if %errorlevel% == 2 goto office
if %errorlevel% == 3 goto unwindows
if %errorlevel% == 4 goto unoffice
if %errorlevel% == 5 goto checkwin
if %errorlevel% == 6 goto officeverchange

::检查KMS服务器
:checkkms
cls
echo.
echo.
echo.
echo. 正在检查激活服务器...
ping %kmsroot% | find /i "来自" >nul && ( goto :EOF ) || ( goto fail )


::激活
:windows
cls
echo.
echo.
echo.
call :checkkms
echo.
echo. 已成功连接激活服务器！
cscript //Nologo %windir%\system32\slmgr.vbs /xpr | find "已永久激活" > NUL && goto wintooff

::检查系统版本
ver | find "6.0." > NUL &&  goto winvista
ver | find "6.1." > NUL &&  goto win7
ver | find "6.2." > NUL &&  goto win8
ver | find "6.3." > NUL &&  goto win81
ver | find "10.0." > NUL &&  goto win10
echo.
echo. 该Windows版本不支持KMS激活！
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin


::安装密钥
:winvista
echo.
echo. 当前Windows版本为 Windows Vista/2008，正在激活...
echo.

::vista

set Business=YFKBB-PQJJV-G996G-VWGXY-2V3X8
set BusinessN=HMBQG-8H2RH-C77VX-27R82-VMQBT

set Enterprise=VKK3X-68KWM-X2YGT-QR4M6-4BWMV
set EnterpriseN=VTC42-BM838-43QHV-84HX6-XJXKV

::2008

set ServerWeb=WYR28-R7TFJ-3X2YQ-YCY4H-M249D
set ServerStandard=TM24T-X9RMF-VWXK6-X8JC9-BFGM2

set ServerStandardV=W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
set ServerEnterprise=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V

set ServerEnterpriseV=39BXF-X8Q23-P2WWT-38T2F-G3FPG
set ServerWeb=RCTX3-KWVHP-BR6TB-RB6DM-6X7HP

set ServerDatacenter=7M67G-PC374-GR742-YH8V4-TCBY3
set ServerDatacenterV=22XQ2-VRXRG-P8D42-K34TD-G3QQC

set ServerEnterpriseIA64=4DWFP-JF3DJ-B7DTH-78FJB-PDRHK

goto windowsstart


::安装密钥
:win7
echo.

::win7

echo. 当前Windows版本为 Windows 7/2008 R2，正在激活...
echo.
set Professional=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
set ProfessionalN=MRPKT-YTG23-K7D7T-X2JMM-QY7MG
set ProfessionalE=W82YF-2Q76Y-63HXB-FGJG9-GF7QX

set Enterprise=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
set EnterpriseN=YDRBP-3D83W-TY26F-D46B2-XCKRJ
set EnterpriseE=C29WB-22CC8-VJ326-GHFJW-H9DH4

::2008R2

set ServerWeb=6TPJF-RBVHG-WBW2R-86QPH-6RTM4
set ServerHPC=TT8MH-CG224-D3D7Q-498W2-9QCTX

set ServerStandard=YC6KT-GKW9T-YTKYR-T4X34-R7VHC
set ServerEnterprise=489J6-VHDMP-X63PK-3K798-CPX3Y

set ServerDatacenter=74YFP-3QFB3-KQT8W-PMXWJ-7M648
set ServerEnterpriseIA64=GT63C-RJFQ3-4GMB6-BRFB9-CB83V

goto windowsstart


::安装密钥
:win8
echo.
echo. 当前Windows版本为 Windows 8/2012，正在激活...
echo.

::win8

set Professional=NG4HW-VH26C-733KW-K6F98-J8CK4
set ProfessionalN=XCVCF-2NXM9-723PB-MHCB7-2RYQQ

set Enterprise=32JNW-9KQ84-P47T8-D8GGY-CWCK7
set EnterpriseN=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT

::2012

set Core=BN3D2-R7TKB-3YPBD-8DRP2-27GG4
set CoreN=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY

set CoreSingleLanguage=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
set CoreCountrySpecific=4K36P-JN4VD-GDC6V-KDT89-DYFKP

set ServerStandard=XC9B7-NBPP2-83J2H-RHMBY-92BT4
set ServerMultiPointStandard=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ

set ServerMultiPointPremium=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
set ServerDatacenter=48HP8-DN98B-MYWDG-T2DCC-8W83P

goto windowsstart


::安装密钥
:win81
echo.
echo. 当前Windows版本为 Windows 8.1/2012R2，正在激活...
echo.

::win8.1

set Professional=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
set ProfessionalN=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY

set Enterprise=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
set EnterpriseN=TT4HM-HN7YT-62K67-RGRQJ-JFFXW

::2012R2

set ServerDatacenter=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
set ServerStandard=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
set ServerSolution=KNC87-3J2TX-XB4WP-VCPJV-M4FWM

goto windowsstart


::安装密钥
:win10
echo.
echo. 当前Windows版本为 Windows 10/2016，正在激活...
echo.

::win10

set Professional=W269N-WFGWX-YVC9B-4J6C9-T83GX
set ProfessionalN=MH37W-N47XK-V7XM9-C7227-GCQG9

set Enterprise=NPPR9-FWDCX-D2C8J-H872K-2YT43
set EnterpriseN=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4

set Education=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
set EducationN=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ


set EnterpriseS=WNMTR-4C88C-JK8YV-HQ7T2-76DF9
set EnterpriseSN=2F77B-TNFGY-69QQF-B8YKP-D69TJ

set EnterpriseN=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
set EnterpriseSN=QFFDN-GRT3P-VKWWX-X7T3R-8B639

set EnterpriseG=YYVX9-NTFWV-6MDM3-9PT4T-4M68B
set EnterpriseGN=44RPN-FTY23-9VTTB-MP9BX-T84FV

::2016

set ServerDatacenter=CB7KF-BWN84-R7R2Y-793K2-8XDDG
set ServerStandard=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
set ServerEssentials=JCKRF-N37P4-C2D82-9YXRT-4M63B

goto windowsstart


::激活Windows
:windowsstart
cls
echo.
echo.
echo.
echo. 正在激活Windows...
for /f "tokens=3 delims= " %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "EditionID"') do set EditionID=%%i
if defined %EditionID% (
	cscript //Nologo %windir%\system32\slmgr.vbs /ipk !%EditionID%! >nul
	cscript //Nologo %windir%\system32\slmgr.vbs /skms %kmsroot% >nul
	cscript //Nologo %windir%\system32\slmgr.vbs /ato >nul
	cscript //Nologo %windir%\system32\slmgr.vbs /xpr | find /i "激活" >nul && (
	echo. & echo. ***** 激活成功 ***** & echo.) || (echo. & echo. ***** 激活失败 ***** & echo.)
) else (
	echo.
	echo. 激活失败！请检查你的Windows是否为批量VL版本。
)
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin


::激活个锤子
:wintooff
cls
echo.
echo.
echo.
echo. 当前Windows已永久激活！无需重复激活！
echo.
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin


::检查Office版本
:office
cls
echo.
echo.
echo.
call :checkkms
echo.
echo. 已连接激活服务器....
echo.
echo. 正在检查安装的 Office 版本...
call :GetOfficePath 14 Office2010
call :ActOffice 14 Office2010
call :GetOfficePath 15 Office2013
call :ActOffice 15 Office2013
echo.
echo. 正在检查 Office2016 的安装路径...
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles%\Microsoft Office\Office16
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles(x86)%\Microsoft Office\Office16
if DEFINED _Office16Path (cls & echo. & echo. & echo. & echo. 已发现 Office2016
	call :ActOffice 16 Office2016
  ) else (cls & echo. & echo. & echo. & echo. 未发现Office2016！可能未安装在默认安装路径！ & echo. & echo. &echo.)
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin


::检测office路径
:GetOfficePath
echo.
echo. 正在检测 %2 的安装路径...
set _Office%1Path=
set _Reg32=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot
set _Reg64=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot
reg query "%_Reg32%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg32%" /v "Path"') do SET "_OfficePath1=%%b"
reg query "%_Reg64%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg64%" /v "Path"') do SET "_OfficePath2=%%b"
if DEFINED _OfficePath1 (if exist "%_OfficePath1%ospp.vbs" set _Office%1Path=!_OfficePath1!)
if DEFINED _OfficePath2 (if exist "%_OfficePath2%ospp.vbs" set _Office%1Path=!_OfficePath2!)
set _OfficePath1=
set _OfficePath2=
if DEFINED _Office%1Path (cls & echo. & echo. & echo. & echo. 已发现 %2) else (cls & echo. & echo. & echo. & echo. 未发现 %2)
goto :EOF


::激活office
:ActOffice
if DEFINED _Office%1Path (
	cd /d "!_Office%1Path!"
	echo.
	echo. 正在检查 %2 的激活状态...
	cscript //nologo ospp.vbs /act | find /i "successful" > NUL && (
		echo. & echo. ***** %2 已激活！无需重复激活！ ***** & echo.) || (
		echo. & echo. %2 未激活，正在激活...
		if %1 EQU 14 call :Licens14
		if %1 EQU 15 call :Licens15
		if %1 EQU 16 call :Licens16
		cscript //nologo ospp.vbs /sethst:%kmsroot% >nul
		cscript //nologo ospp.vbs /act | find /i "successful" >nul && (
		echo. & echo. ***** %2 激活成功 ***** & echo.) || (echo. & echo. ***** %2 激活失败 ***** & echo.)
		)
cd /d "%~dp0"
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin
)
cd /d "%~dp0"
goto :EOF


::导入office激活密钥
:Licens14
for /f %%x in ('dir /b ..\root\Licenses14\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses14\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" >nul

cscript ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul
cscript ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul
cscript ospp.vbs /inpkey:D6QFG-VBYP2-XQHM7-J97RH-VVRCK >nul
goto :EOF


::导入office激活密钥
:Licens15
for /f %%x in ('dir /b ..\root\Licenses15\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses15\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" >nul

cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT >nul
cscript ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4 >nul
goto :EOF


::导入office激活密钥
:Licens16
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul
cscript ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM >nul
goto :EOF


::检测到xp
:xptooff
cls
echo.
echo.
echo.
echo. ***** 当前系统为 WindowsXP 或 Windows2003！不支持KMS激活。 *****
echo.
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin


::检查失败
:fail
cls
echo.
echo.
echo.
echo. ***** 无法连接到服务器，请检查网络连接，或前往github.com/yareiy/yareiy-kms获取最新版本。 *****
echo.
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin


::卸载Windows激活
:unwindows
cls
echo.
echo.
echo.
echo. 正在卸载 Windows KMS激活...
	cscript //Nologo %windir%\system32\slmgr.vbs /upk >nul
	cscript //Nologo %windir%\system32\slmgr.vbs /ckms >nul
	cscript //Nologo %windir%\system32\slmgr.vbs /xpr | find /i "错误" >nul && (
	echo. & echo. ***** 卸载成功 ***** & echo.) || (echo. & echo. ***** 卸载失败 ***** & echo.)	
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin


::卸载office激活
:unoffice
cls
echo.
echo.
echo.
echo. 正在检查安装的 Office 版本...
call :GetOfficePath 14 Office2010
call :unoffice 14 Office2010
call :GetOfficePath 15 Office2013
call :unoffice 15 Office2013
echo.
echo. 正在检测 Office2016 系列产品的安装路径...
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles%\Microsoft Office\Office16
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles(x86)%\Microsoft Office\Office16
if DEFINED _Office16Path (cls & echo. & echo. & echo. & echo. 已发现 Office2016
	call :unoffice 16 Office2016
  ) else (cls & echo. & echo. & echo. & echo. 未发现Office2016！可能未安装在默认安装路径！ & echo. & echo. & echo.)
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin


:checkwin
cls
echo.
echo.
echo.
echo. [1]-检查Windows激活详情
echo. [2]-检查Windows激活状态
echo. [3]-返回主菜单
echo.
echo.
choice /c 123 /n /m "请按下要选择的选项对应的数字【1-3】："

echo. %errorlevel%
if %errorlevel% == 1 goto slmgrdlv
if %errorlevel% == 2 goto slmgrxpr
if %errorlevel% == 3 goto begin


:officeverchange
::检测office安装路径
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set officev=Microsoft Office2016/Office365
cd /d "%ProgramFiles%\Microsoft Office\Office16"

if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set officev=Microsoft Office2016/Office365
cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" set officev=Microsoft Office2013
cd /d "%ProgramFiles%\Microsoft Office\Office15"

if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" set officev=Microsoft Office2013
cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15"

if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" set officev=Microsoft Office2010
cd /d "%ProgramFiles%\Microsoft Office\Office14"

if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" set officev=Microsoft Office2010
cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14"

if not defined officev set officev=未检测到Office！你可能改变了默认安装目录！

::菜单
cls
echo.
echo.
echo.
echo.    [0]-返回主菜单
echo.
echo.    [1]-Retail版 Office Pro Plus 2016/365 转Vol版
echo.    [2]-Retail版 Office Visio Pro 2016/365 转Vol版
echo.    [3]-Retail版 Office Project Pro 2016/365 转Vol版
echo.
echo.    [4]-Retail版 Office Pro Plus 2013 转Vol版
echo.    [5]-Retail版 Office Visio Pro 2013 转Vol版
echo.    [6]-Retail版 Office Project Pro 2013 转Vol版
echo.
echo.    [7]-Retail版 Office Pro Plus 2010 转Vol版
echo.    [8]-Retail版 Office Visio Pro 2010 转Vol版
echo.    [9]-Retail版 Office Project Pro 2010 转Vol版
echo.
echo.  已安装：%officev%
echo.
echo.  如果你的Office是安装在非默认目录请：
echo.  复制本批处理到 office16/office15/office14 目录后运行
echo.
choice /c 1234567890 /n /m "请选择【0-9】："

echo. %errorlevel%
if %errorlevel% == 1 goto ProPlus2016
if %errorlevel% == 2 goto VisioPro2016
if %errorlevel% == 3 goto ProjectPro2016
if %errorlevel% == 4 goto ProPlus2013
if %errorlevel% == 5 goto VisioPro2013
if %errorlevel% == 6 goto ProjectPro2013
if %errorlevel% == 7 goto ProPlus2010
if %errorlevel% == 8 goto VisioPro2010
if %errorlevel% == 9 goto ProjectPro2010
if %errorlevel% == 10 goto begin


::版本转换

:ProPlus2016

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99

goto :e

:VisioPro2016

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK

goto :e

:ProjectPro2016

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses16\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses16\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT

goto :e

:ProPlus2013

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses15\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses15\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" >nul

cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT

goto :e

:VisioPro2013

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses15\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses15\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" >nul

cscript ospp.vbs /inpkey:C2FG9-N6J68-H8BTJ-BW3QX-RM3B3

goto :e

:ProjectPro2013

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses15\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses15\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%%x" >nul

cscript ospp.vbs /inpkey:FN8TT-7WMH6-2D4X9-M337T-2342K

goto :e

:ProPlus2010

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses14\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses14\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" >nul

cscript ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB

goto :e

:VisioPro2010

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses14\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses14\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" >nul

cscript ospp.vbs /inpkey:7MCW8-VRQVK-G677T-PDJCM-Q8TCP

goto :e

:ProjectPro2010

cls

echo.
echo.
echo.

echo. 正在安装 KMS 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses14\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" >nul

echo. 正在安装 MAK 密钥...

echo.

for /f %%x in ('dir /b ..\root\Licenses14\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses14\%%x" >nul

cscript ospp.vbs /inpkey:YGX6F-PGV49-PGW3J-9BTGG-VHKC6

goto :officeverchangeend


::操作完成
:officeverchangeend

echo.

echo. Done！ -按任意键返回主菜单-

pause >nul

goto officeverchange


:slmgrdlv
wscript //Nologo %windir%\system32\slmgr.vbs /dlv >nul
goto checkwin


:slmgrxpr
wscript //Nologo %windir%\system32\slmgr.vbs /xpr >nul
goto checkwin


::卸载office激活
:unoffice
if DEFINED _Office%1Path (
	cd /d "!_Office%1Path!"
	echo.
	echo. 正在卸载 %2  KMS 激活...
		cscript //nologo ospp.vbs /unpkey:H3GVB >nul
		cscript //nologo ospp.vbs /unpkey:8R6BM >nul
		cscript //nologo ospp.vbs /unpkey:VVRCK >nul
		cscript //nologo ospp.vbs /unpkey:T7DDX >nul
		cscript //nologo ospp.vbs /unpkey:CW9BM >nul
		cscript //nologo ospp.vbs /unpkey:4T3J4 >nul
		cscript //nologo ospp.vbs /unpkey:BT37T >nul
		cscript //nologo ospp.vbs /unpkey:D3XHX >nul
		cscript //nologo ospp.vbs /unpkey:X3DWQ >nul
		cscript //nologo ospp.vbs /unpkey:P4VTT >nul
		cscript //nologo ospp.vbs /unpkey:VHKC6 >nul
		cscript //nologo ospp.vbs /unpkey:F9PGB >nul
		cscript //nologo ospp.vbs /unpkey:83YTP >nul
		cscript //nologo ospp.vbs /unpkey:CRY7T >nul
		cscript //nologo ospp.vbs /unpkey:WX8BJ >nul
		cscript //nologo ospp.vbs /unpkey:Q8TCP >nul
		cscript //nologo ospp.vbs /unpkey:KHFGJ >nul

		cscript //nologo ospp.vbs /unpkey:GVGXT >nul
		cscript //nologo ospp.vbs /unpkey:92CD4 >nul
		cscript //nologo ospp.vbs /unpkey:2342K >nul
		cscript //nologo ospp.vbs /unpkey:8QHTT >nul
		cscript //nologo ospp.vbs /unpkey:RM3B3 >nul
		cscript //nologo ospp.vbs /unpkey:PGWR7 >nul
		cscript //nologo ospp.vbs /unpkey:4JM2D >nul
		cscript //nologo ospp.vbs /unpkey:BG7GB >nul
		cscript //nologo ospp.vbs /unpkey:F8894 >nul
		cscript //nologo ospp.vbs /unpkey:TCK7R >nul
		cscript //nologo ospp.vbs /unpkey:P34VW >nul
		cscript //nologo ospp.vbs /unpkey:2PMBT >nul
		cscript //nologo ospp.vbs /unpkey:4RD4F >nul
		cscript //nologo ospp.vbs /unpkey:FCXK4 >nul
		cscript //nologo ospp.vbs /unpkey:4GBJ7 >nul

		cscript //nologo ospp.vbs /unpkey:WFG99 >nul
		cscript //nologo ospp.vbs /unpkey:DRTFM >nul
		cscript //nologo ospp.vbs /unpkey:G83KT >nul
		cscript //nologo ospp.vbs /unpkey:KQBVC >nul
		cscript //nologo ospp.vbs /unpkey:RJRJK >nul
		cscript //nologo ospp.vbs /unpkey:W8GF4 >nul
		cscript //nologo ospp.vbs /unpkey:QPFDW >nul
		cscript //nologo ospp.vbs /unpkey:7FTBF >nul
		cscript //nologo ospp.vbs /unpkey:XW3J6 >nul
		cscript //nologo ospp.vbs /unpkey:6MT9B >nul
		cscript //nologo ospp.vbs /unpkey:BY6C6 >nul
		cscript //nologo ospp.vbs /unpkey:8837K >nul
		cscript //nologo ospp.vbs /unpkey:DDBV6 >nul
		cscript //nologo ospp.vbs /unpkey:3PFJ6 >nul
		cscript //nologo ospp.vbs /remhst | find /i "successful" >nul && (
		echo. & echo. ***** 卸载成功 ***** & echo.) || (echo. & echo. ***** 卸载失败 ***** & echo.)
cd /d "%~dp0"
echo.
echo. Done！ -按任意键返回主菜单-
pause>nul
goto begin
)
cd /d "%~dp0"
goto :EOF
