::AUTHOR: Christopher Weller
::DATE: 02 March 2017
::PURPOSE: Sets the IP address of my laptop to the correct setting for various lines.
::
@ECHO OFF

SET /P CHOICE="Choose a line to change the IP address for (type fa=Front Axle, h=Heat Treat, g=Green Machining, m=Mating, p=Pinion, R=Ring, or d= GM Online): "

2>NUL CALL :CASE_%CHOICE% # jump to :CASE_fa, :CASE_h, etc.
IF ERRORLEVEL 1 CALL :DEFAULT_CASE # if label doesn't exist

ipconfig /all
pause
cmd
EXIT /B

:CASE_d
  netsh interface ipv4 set address name="Wired Network" dhcp
  GOTO END_CASE

:CASE_g
  netsh interface ipv4 set address name="Wired Network" static 120.165.246.253 255.255.255.0 120.165.246.1
  GOTO END_CASE

:CASE_fa
  netsh interface ipv4 set address name="Wired Network" static 120.165.245.45 255.255.254.0 120.165.244.1
  GOTO END_CASE

:CASE_h
  netsh interface ipv4 set address name="Wired Network" static 120.165.243.253 255.255.255.0 120.165.243.1
  GOTO END_CASE

:CASE_m
  netsh interface ipv4 set address name="Wired Network" static 120.165.247.253 255.255.255.0 120.165.247.1
  GOTO END_CASE

:CASE_p
  netsh interface ipv4 set address name="Wired Network" static 120.165.252.253 255.255.255.0 120.165.252.1
  GOTO END_CASE

:CASE_r
  netsh interface ipv4 set address name="Wired Network" static 120.165.251.253 255.255.255.0 120.165.251.1
  GOTO END_CASE

:DEFAULT_CASE
  ECHO Unknown choice "%CHOICE%"
  GOTO END_CASE

:END_CASE
  VER > NUL # reset ERRORLEVEL
  GOTO :EOF # return from CALL


::cd\

::cd Programs Files\Kepware\KEPServerEX 5

::net stop "KEPServerEXV5"

::REM: cd\ (LOCATION OF PROJECT TO PLACE INTO RUNTIME)\project

::copy C:\project.opf C:\ProgramData\Kepware\KEPServerEX\V5\*.*
::copy C:\Users\qzf681\Documents\GAxx\Grob\52969_GRO_GAxx_FRONT_ASSEMBLY_GROB_OP10-OP435_G3.1_ROCKWELL_Machine_Type(A)_08032016.opf C:\ProgramData\Kepware\KEPServerEX\V5\Project.opf
::net start "KEPServerEXV5"