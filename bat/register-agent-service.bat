@echo off

rem %1 - version number

rem %2 - first 2 port characters. for 1540,1541,1560:1591 it will be 15

rem %3 - cluster reg catalog

rem %4 - user password

rem register-service 8.3.11.3034 25 "C:\Program Files\1cv8\srvinfo2541" strongpwd

set SrvUserName=.\USR1CV8

set SrvUserPwd=%4

set RangePort=%260:%291

set BasePort=%241

set CtrlPort=%240

set SrvcName="1C:Enterprise 8.3 Server Agent %CtrlPort% %1"

set BinPath="\"C:\Program Files\1cv8\%1\bin\ragent.exe\" -srvc -agent -regport %BasePort% -port %CtrlPort% -range %RangePort% -d \"%~3\" -debug"

set Desctiption="1C:Enterprise 8.3 Server Agent. Parameters: %1, %CtrlPort%, %BasePort%, %RangePort%"

if not exist "%~3" mkdir "%~3"

sc stop %SrvcName%

sc delete %SrvcName%

sc create %SrvcName% binPath= %BinPath% start= auto obj= %SrvUserName% password= %SrvUserPwd% displayname= %Desctiption% depend= Tcpip/Dnscache/lanmanworkstation/lanmanserver/