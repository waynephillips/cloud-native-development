@echo off
cls
Echo.
Echo Copies setup files for SAPGUI Install and Upgrade
Echo.
Echo If asked to login, use your DOW\USERID and PASSWORD
Echo.
Echo Example:  DOW\N004563
Echo.
Echo You will not see your password being typed!
Echo.

rem net use u: \\downeavm061.dow.com\Installation_Server /savecred
net config server /autodisconnect:-1
net use u: /delete /y
net use \\downeavm061.dow.com /d /y
net use \\downeavm061.dow.com\Installation_Server /d /y
rem net use u: \\downeavm061.dow.com\Installation_Server /savecred /persistent:yes
net use u: \\downeavm061.dow.com\Installation_Server

rem this fails copy /Y u:\Scripts\install.2.vbs %public%\desktop\install.2.vbs

rem copy /Y u:\Scripts\install.1.bat %userprofile%\desktop\install.1.bat
copy /Y "\\downeavm061.dow.com\Installation_Server\Scripts\install.2.vbs" "%userprofile%\desktop\install.2.vbs"
copy /Y "\\downeavm061.dow.com\Installation_Server\Scripts\run.as.admin.reg" "%userprofile%\desktop\run.as.admin.reg"

copy /Y "u:\Scripts\install.2.vbs" "%userprofile%\desktop\install.2.vbs"
copy /Y "u:\Scripts\run.as.admin.reg" "%userprofile%\desktop\run.as.admin.reg"

rem copy /Y "u:\Scripts\install.2.vbs" "c:\users\public\desktop\install.2.vbs"
rem copy /Y "u:\Scripts\run.as.admin.reg" c:\users\public\desktop\run.as.admin.reg"

rem copy /Y "u:\Scripts\install.2.vbs" "%userprofile%\downloads\install.2.vbs"
rem copy /Y "u:\Scripts\run.as.admin.reg" "%userprofile%\downloads\run.as.admin.reg"

rem copy /Y "u:\Scripts\install.2.vbs" "c:\users\public\downloads\install.2.vbs"
rem copy /Y "u:\Scripts\run.as.admin.reg" c:\users\public\downloads\run.as.admin.reg"
rem it wont disconnect, files in use...
rem net use u: /delete

rem beleive this one works:
regedit -s %userprofile%\desktop\run.as.admin.reg

rem this doesnt work:
rem REG IMPORT %userprofile%\desktop\run.as.admin.reg

cls

Echo.
Echo --==** IMPORTANT ** ==--
Echo --==** IMPORTANT ** ==--
Echo --==** IMPORTANT ** ==--
Echo.
Echo Now do these next steps:
Echo.
Echo Right click and RUN AS ADMINISTRATOR for
Echo "install.2.vbs" which is on your Desktop.
Echo.
Echo (If you do not have the RUN AS ADMINISTRATOR option
Echo  then Run "run.as.admin.reg" from your Desktop.)
Echo. 
pause
rem net use u: /delete