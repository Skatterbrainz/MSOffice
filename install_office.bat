@echo off
rem *****************************************************************************
rem Filename...... install.bat
rem Author........ David Stein
rem Date.......... 2019.01.09
rem Purpose....... Upgrade to Office 365 latest build
rem Comment....... Modified for use at Contoso Corp
rem Folder structure: 
rem \office365\
rem \office365\offscrub
rem \office365\office
rem files: office365.xml, projectpro.xml, visiopro.xml, allproducts.xml
rem *****************************************************************************
title Office Removal Package
CLS
SET APPNAME=InstallO365ProPlus
SET EXITCODE=0
SET LOG=%TEMP%\Contoso_OfficeScrub.log
echo %DATE% %TIME% info: scriptversion... 2019.01.09 >%LOG%
echo %DATE% %TIME% info: computername.... %COMPUTERNAME% >>%LOG%
echo %DATE% %TIME% info: user context.... %USERNAME% >>%LOG%
echo %DATE% %TIME% info: windows path.... %WINDIR% >>%LOG%
echo %DATE% %TIME% info: program files... %PROGRAMFILES% >>%LOG%
goto DETECTPRODUCTS

:DETECTPRODUCTS
echo %DATE% %TIME% Detecting installed Microsoft Office products...
echo ------------------- detecting products ------------------- >>%LOG%
set VISIO=0
set PROJECT=0
if exist "C:\Program Files (x86)\Microsoft Office\Office15\VISIO.exe" (
  echo %DATE% %TIME% info: found Visio 2013 32bit >>%LOG%"
  set VISIO=1
)
if exist "C:\Program Files (x86)\Microsoft Office\Office14\VISIO.exe" (
  echo %DATE% %TIME% info: found Visio 2010 32bit >>%LOG%"
  set VISIO=1
)
if exist "C:\Program Files (x86)\Microsoft Office\OFFICE11\VISIO.exe" (
  echo %DATE% %TIME% info: found Visio 2007 32bit >>%LOG%"
  set VISIO=1
)
if exist "C:\Program Files (x86)\Microsoft Office\root\Office16\VISIO.exe" (
  echo %DATE% %TIME% info: found Visio 2016 32bit >>%LOG%"
  set VISIO=1
)
if exist "C:\Program Files\Microsoft Office\root\Office16\VISIO.exe" (
  echo %DATE% %TIME% info: found Visio 2016 64bit >>%LOG%"
  set VISIO=1
)
if exist "C:\Program Files (x86)\Microsoft Office\Office15\WINPROJ.exe" (
  echo %DATE% %TIME% info: found Project 2013 32bit >>%LOG%"
  set PROJECT=1
)
if exist "C:\Program Files (x86)\Microsoft Office\Office14\WINPROJ.exe" (
  echo %DATE% %TIME% info: found Project 2010 32bit >>%LOG%"
  set PROJECT=1
)
if exist "C:\Program Files (x86)\Microsoft Office\OFFICE11\WINPROJ.exe" (
  echo %DATE% %TIME% info: found Project 2007 32bit >>%LOG%"
  set PROJECT=1
)
if exist "C:\Program Files (x86)\Microsoft Office\root\Office16\WINPROJ.exe" (
  echo %DATE% %TIME% info: found Project 2016 32bit >>%LOG%"
  set PROJECT=1
)
if exist "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.exe" (
  echo %DATE% %TIME% info: found Project 2016 64bit >>%LOG%"
  set PROJECT=1
)
echo %DATE% %TIME% info: VISIO is %VISIO% >>%LOG%
echo %DATE% %TIME% info: PROJECT is %PROJECT% >>%LOG%
goto SCRUB

:SCRUB
echo Removing Microsoft Office products...
echo --------------------- Office 2003 ------------------------ >>%LOG%
echo Removing Office 2003...
echo %DATE% %TIME% info: command is cscript.exe /nologo "%~dp0offscrub\OffScrub03.vbs" ALL /Quiet /Log %TEMP% >>%LOG%
cscript.exe /nologo "%~dp0offscrub\OffScrub03.vbs" ALL /Quiet /Log %TEMP%
echo %DATE% %TIME% info: exit code is %ERRORLEVEL% >>%LOG%
echo --------------------- Office 2007 ------------------------ >>%LOG%
echo Removing Office 2007...
echo %DATE% %TIME% info: command is cscript.exe /nologo "%~dp0offscrub\OffScrub07.vbs" ALL /Quiet /Log %TEMP% >>%LOG%
cscript.exe /nologo "%~dp0offscrub\OffScrub07.vbs" ALL /Quiet /Log %TEMP%
echo %DATE% %TIME% info: exit code is %ERRORLEVEL% >>%LOG%
echo --------------------- Office 2010 ------------------------ >>%LOG%
echo Removing Office 2010...
echo %DATE% %TIME% info: command is cscript.exe /nologo "%~dp0offscrub\OffScrub10.vbs" ALL /Quiet /Log %TEMP% >>%LOG%
cscript.exe /nologo "%~dp0offscrub\OffScrub10.vbs" ALL /Quiet /Log %TEMP%
echo %DATE% %TIME% info: exit code is %ERRORLEVEL% >>%LOG%
echo --------------------- Office 2013 ------------------------ >>%LOG%
echo Removing Office 2013...
echo %DATE% %TIME% info: command is cscript.exe /nologo "%~dp0offscrub\OffScrub_O15msi.vbs" ALL /Quiet /Log %TEMP% >>%LOG%
cscript.exe /nologo "%~dp0offscrub\OffScrub_O15msi.vbs" ALL /Quiet /Log %TEMP%
echo %DATE% %TIME% info: exit code is %ERRORLEVEL% >>%LOG%
echo --------------------- Office 2016 ------------------------ >>%LOG%
echo Removing Office 2016...
echo %DATE% %TIME% info: command is cscript.exe /nologo "%~dp0offscrub\OffScrub_O16msi.vbs" ALL /Quiet /Log %TEMP% >>%LOG%
cscript.exe /nologo "%~dp0offscrub\OffScrub_O16msi.vbs" ALL /Quiet /Log %TEMP%
echo %DATE% %TIME% info: exit code is %ERRORLEVEL% >>%LOG%
echo --------------------- Office C2R ------------------------ >>%LOG%
echo Removing Office 2016 C2R...
echo %DATE% %TIME% info: command is cscript.exe /nologo "%~dp0offscrub\OffScrubC2R.vbs" /Quiet /ReturnErrorOrSuccess CLIENTALL /Log %TEMP% >>%LOG%
cscript.exe /nologo "%~dp0offscrub\OffScrubC2R.vbs" /Quiet /ReturnErrorOrSuccess /Log %TEMP%
echo %DATE% %TIME% info: exit code is %ERRORLEVEL% >>%LOG%
goto VISIOVIEWER

:VISIOVIEWER
echo --------------------- Visio Viewer 2007 ------------------------ >>%LOG%
echo Removing Visio Viewer 2007...
echo %DATE% %TIME% info: removing Visio Viewer 2007... >>%LOG%
echo %DATE% %TIME% info: command is msiexec.exe /x {95120000-0052-0409-0000-0000000FF1CE} /qn /norestart >NUL >>%LOG%
msiexec.exe /x {95120000-0052-0409-0000-0000000FF1CE} /qn /norestart >NUL
if %ERRORLEVEL%==1605 (
  echo %DATE% %TIME% info: product was not found on this computer >>%LOG%
) else (
  echo %DATE% %TIME% info: exit code is %ERRORLEVEL% >>%LOG%
)
goto ADDONS

:ADDONS
echo --------------------- Cisco ViewMail ------------------------ >>%LOG%
echo Removing Addons...
echo %DATE% %TIME% info: removing Cisco ViewMail for Outlook >>%LOG%
echo %DATE% %TIME% info: command is msiexec.exe /x {D43281E3-D934-46D3-8341-66B7B4BFC626} /qn /norestart >NUL >>%LOG%
msiexec.exe /x {D43281E3-D934-46D3-8341-66B7B4BFC626} /qn /norestart >NUL
if %ERRORLEVEL%==1605 (
  echo %DATE% %TIME% info: product was not found on this computer >>%LOG%
) else (
  echo %DATE% %TIME% info: exit code is %ERRORLEVEL% >>%LOG%
)
goto SHORTCUTS

:SHORTCUTS
echo -------------------- Shortcuts ------------------------- >>%LOG%
echo Removing shortcuts...
echo %DATE% %TIME% info: hunting down leftover Office 2013 shortcuts... >>%LOG%
if exist "%PUBLIC%\Desktop\Access 2013.lnk" del "%PUBLIC%\Desktop\Access 2013.lnk" /F /Q >NUL
if exist "%PUBLIC%\Desktop\Excel 2013.lnk" del "%PUBLIC%\Desktop\Excel 2013.lnk" /F /Q >NUL
if exist "%PUBLIC%\Desktop\OneNote 2013.lnk" del "%PUBLIC%\Desktop\OneNote 2013.lnk" /F /Q >NUL
if exist "%PUBLIC%\Desktop\Outlook 2013.lnk" del "%PUBLIC%\Desktop\Outlook 2013.lnk" /F /Q >NUL
if exist "%PUBLIC%\Desktop\PowerPoint 2013.lnk" del "%PUBLIC%\Desktop\PowerPoint 2013.lnk" /F /Q >NUL
if exist "%PUBLIC%\Desktop\Publisher 2013.lnk" del "%PUBLIC%\Desktop\Publisher 2013.lnk" /F /Q >NUL
if exist "%PUBLIC%\Desktop\Word 2013.lnk" del "%PUBLIC%\Desktop\Word 2013.lnk" /F /Q >NUL
if exist "%PUBLIC%\Desktop\Skype for Business 2016.lnk" del "%PUBLIC%\Desktop\Skype for Business 2016.lnk" /F /Q >NUL
if exist "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013" (
  echo %DATE% %TIME% info: removing Office 2013 apps folder from Start Menu..." >>%LOG%
  DEL "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013" /F /Q >NUL
  RD ""%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013" /S /Q >NUL
)
if exist "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools" (
  echo %DATE% %TIME% info: removing Office 2016 tools folder from start menu..." >>%LOG%
  DEL "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools\*.*" /F /Q >NUL
  RD "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools" /S /Q >NUL
)
goto INSTALL

:INSTALL
echo Installing Office 365 ProPlus...
echo ----------------------- Installation ------------------- >>%LOG%
echo %DATE% %TIME% info: installing Microsoft Office 365 Pro Plus >>%LOG%
if %PROJECT%==1 (
  if %VISIO%==1 (
    echo %DATE% %TIME% info: installing with Project Pro and Visio Pro... >>%LOG%
    echo %DATE% %TIME% info: command: %~dp0setup.exe /configure %~dp0allproducts.xml >>%LOG%
    %~dp0setup.exe /configure %~dp0allproducts.xml
  ) else (
    echo %DATE% %TIME% info: installing with Project Pro... >>%LOG%
    echo %DATE% %TIME% info: command: %~dp0setup.exe /configure %~dp0projectpro.xml >>%LOG%
    %~dp0setup.exe /configure %~dp0projectpro.xml
  )
) else (
  if %VISIO%==1 (
    echo %DATE% %TIME% info: installing with Visio Pro... >>%LOG%
    echo %DATE% %TIME% info: command: %~dp0setup.exe /configure %~dp0visiopro.xml >>%LOG%
    %~dp0setup.exe /configure %~dp0visiopro.xml
  ) else (
    echo %DATE% %TIME% info: installing Office 365 ProPlus only... >>%LOG%
    echo %DATE% %TIME% info: command: %~dp0setup.exe /configure %~dp0office365.xml >>%LOG%
    %~dp0setup.exe /configure %~dp0office365.xml
  )
)
goto END

:END
echo completed!
echo %DATE% %TIME% info: completed >>%LOG%
EXIT /B %EXITCODE%
