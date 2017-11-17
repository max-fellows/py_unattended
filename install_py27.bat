@echo off
setlocal EnableDelayedExpansion

if "%~1" == "" goto usage

if "%~2" == "" (set PYTHONHOME=c:\python27) else (set PYTHONHOME=%~2)
set PATH=%PYTHONHOME%

set platform=%~1
if "%platform%" == "x86" goto x86
if "%platform%" == "x64" goto x64
goto usage

:x86
set python=python-2.7.14.msi
set access_db_engine=accessdatabaseengine.exe
goto common

:x64
set python=python-2.7.14.amd64.msi
set access_db_engine=accessdatabaseengine_x64.exe
goto common

:common
set i=0
for %%a in (
    %platform%\pip\*.*
    common\pip\*.*
    dbfpy openpyxl python-pptx requests paramiko configobj xlrd xlwt psycopg2
) do (
    set /a i+=1
    set pip_packages[!i!]=%%a
)
    
set num_packages=%i%
goto install

:install
if not exist %PYTHONHOME% (set reinstall=) else (set reinstall=REINSTALL=DefaultFeature)
c:\windows\system32\msiexec /i "%platform%\%python%" TARGETDIR=%PYTHONHOME% %reinstall% /quiet /passive
%platform%\%access_db_engine% /passive /quiet
for /l %%i in (1, 1, %num_packages%) do %PYTHONHOME%\scripts\pip install !pip_packages[%%i]!
goto end

:usage
echo install_py27.bat (x86/x64) [optional: python install path]

:end
