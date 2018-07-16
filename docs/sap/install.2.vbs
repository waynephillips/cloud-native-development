'Script to update, upgrade, remove, or install the SAPGUI & Dow delivered files
'Author: Nick Suter
'Revisions: 2013,2014,2015,2016
'Master copy resides: \\downeavm061\Installation_Server\Scripts
'todo:
'Shared workstations:
'	HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Dow\Dow Workstation\CurrentVersion
'	EWSSHRD
'	put XML in common folder
'	force saplogon to use HKLM to find XMl?
'

Option Explicit 

'If WScript.Arguments.length = 0 Then
  'Set objShell = CreateObject("Shell.Application")
  'Pass a bogus argument with leading blank space, say [ uac]
  'objShell.ShellExecute "wscript.exe", Chr(34) & _
  'WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
'Else

Dim oShell, oSysInfo, FSO, sName, PrgFiles, ProductVersion, sFilePath ,sProgram, PrgFilesx86, service, process, i, f, y, cmd, msiexecfile, ignoreoffice
Dim PathToAppData, temppath, userName, windowsdir, userdomain, pccomputername, destdir, sourcedir, sourcedirexe, scourcedirscripts, EWS, OSBIT, Ostype
Dim sapguidir, SAPSetupDir, regkey, objShell, objFolder, objFolderItem, strServiceName, objWMIService, colListOfServices, objService, objFile, packageoveride
Dim numerr, abouterr, logfile, debug, colProcessList, objProcess, message, code, logging, popup, runningfrom, file, found, docopy, dodel, quieter
Dim numfiles, filetocheck, logfolder, folder, ABC, userprofile, userdnsdomain, computertype, scriptver, iter, done, svc, nwver, desktopver, packagefile
Dim abort, a, x, NEWProductVersion, timeout, SourceDateModified, DestDateModified, upgrade, startTime, endTime, foundsource, founddest, strLine, strLineFile
Dim objReadFile, version, PROCESSORARCHITECTURE, wshNetwork, newfile, filename, filefolder, newstring, keyval, key, reboot
Dim systemdrive, status, jversion, nojava, JavaVer, JavaVer1, Collection, item, javasplit, objFolderFonts, currentstage, window, increment, targetstage
Dim objComputer, colComputers, computer, force, server, dtmConvertedDate, osinfo, os, dtmInstallDate, osystem, usedmdate, returncode, local, sourcedirsapguiexe
Dim count, counthis, allusersprofile, inifiles, shortcutfiles, nwbcfiles, oShellLink, sclocation, scfilename, NWBCdir, publicdesktop, update, packagename
Dim allusersapplicationdata, hive, loglocal, arg, args, doprogress, sccm, silent, ignorejava, thisscript, masterscript, publicfolder, colFiles
Dim prgtorun, params, waitonreturn, nr, strLineInFile, strText, strAction, aLines, badrows, badline, strNewContents, demofile, b, text, aLinesInFile, servicesfile
Dim regentry, keytype, citrix, answer, inifilesparent, gsdservers, hostname, network, dotdowdotcom, nonwbc, publicdocuments, tmppath, quiet, foundvf
Dim skipjavacheck, fullfilename, intAnswer, RC, havedotnet, waserror, uninstall, repair, localappdata, thing, strings, skipdotnetcheck, objWMIService2
Dim getCounttemp, cItems, oItem, colItems, objItem, commfiles, InstalledVersion, tempver, killprocs, created, objRegistry, rootsapfolder
Dim userservicesfile, userservicesfiletext, newservicesfile, newservicesfiletext, tempthisscript, allargs, answer2, runningproc, answer3, allprocsdead
Dim proccount, logonprog, filescopied, foldername, tempprgtorun, fixtxsap, txsapuser, officeversions, officeversion, tempkey, excelpath
Dim installjava, netrelease, fixhistory, commfiles64, whack, disablepersonas, allowjava64, RRC, robocopysouce, robocopydest, robocopyfiles, options
Dim processname, sourcedirfiles, packageDestDateModified

'update this whenever a new 740.exe is created.
'it is for (poorly) providing a way for local installs to know what version of SAPGUI in the the local 740.exe
'not currently in use
Dim packageinfo(1, 0) '2 columns, 1 row
packageinfo(0, 0) = "90930" '740.exe package version
packageinfo(1, 0) = "740003168959" 'sapgui.exe version

Set args = Wscript.Arguments
'If Not WScript.Arguments.Named.Exists("elevate") Then
	'For Each arg In args
	'	allargs = allargs & arg & " "
	'Next
	'msgbox WScript.FullName			'c:\windows\...\wscript.exe
	'msgbox WScript.ScriptFullName 	'...install.2.vbs
	'CreateObject("Shell.Application").ShellExecute WScript.FullName, stringthis(WScript.ScriptFullName) & " /elevate " & allargs, "", "runas", 1
	'msgbox("quittin")
	'WScript.Quit
'End If

'constants
Const strComputer = "."
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const MAX_ITER = 10
Const FONTS = &H14&
Const ALL_USERS_APPLICATION_DATA = &H23& 
Const HKEY_CURRENT_USER = &H80000001
Const FOF_NOCONFIRMATION = &H10&
Const javafile = "jre-8u60-windows-i586.exe"
Const start = "start"

'Things
Set oShell = CreateObject("Wscript.shell")
Set oSysInfo = CreateObject("adSystemInfo") 
Set objShell = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.filesystemobject")
Set wshNetwork = WScript.CreateObject("WScript.Network")
Set objFolderFonts = objShell.Namespace(FONTS)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
Set args = Wscript.Arguments

On Error Resume Next
	set window = createobject("internetexplorer.application")
	If Err <> 0 Then 
		'bad
		doprogress = FALSE
	Else
		doprogress = TRUE
	End If
On Error Goto 0

'variables that need oShell
PROCESSORARCHITECTURE = oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
'PrgFiles = oShell.ExpandEnvironmentStrings("%ProgramFiles%")
'PrgFilesx86 = oShell.ExpandEnvironmentStrings("%ProgramFiles (x86)%")
PathToAppData = oShell.ExpandEnvironmentStrings("%appdata%")	'roaming
localappdata = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")	'local
temppath = oShell.ExpandEnvironmentStrings("%temp%") 'temp
tmppath = oShell.ExpandEnvironmentStrings("%tmp%") 'tmp
userName = ucase(oShell.ExpandEnvironmentStrings("%USERNAME%"))
allusersprofile = oShell.ExpandEnvironmentStrings("%allusersprofile%")
userprofile = oShell.ExpandEnvironmentStrings("%userprofile%")
windowsdir = oShell.ExpandEnvironmentStrings("%SystemRoot%")
pccomputername = oShell.ExpandEnvironmentStrings("%computername%")
userdnsdomain = oShell.ExpandEnvironmentStrings("%USERDNSDOMAIN%")
userdomain = oShell.ExpandEnvironmentStrings("%USERDOMAIN%")
systemdrive = oShell.ExpandEnvironmentStrings("%systemdrive%")
publicfolder = oShell.ExpandEnvironmentStrings("%public%")
cmd = oShell.ExpandEnvironmentStrings("%comspec%")
commfiles=oShell.expandenvironmentstrings("%CommonProgramFiles%")
commfiles64=oShell.expandenvironmentstrings("%CommonProgramFiles(x86)%")

'variables
startTime = Now
sProgram = "sapgui.exe"
computer = ""
servicesfile = windowsdir & "\System32\drivers\etc\services"
msiexecfile = windowsdir & "\System32\msiexec.exe"
dotdowdotcom = ".dow.com"
thisscript = Wscript.ScriptFullName
tempthisscript = replace(thisscript,"\","\\")

'probably constants, no idea why they are here
badrows = ARRAY("ï»¿#Entries for all hR&H & hDow systems","#Entries for all hR&H & hDow systems","SAPMSPR1 3672/TCP","SAPMSCUV 3687/TCP", "SAPMSAUV 3687/TCP",_
		"SAPMSUB1 3680/TCP","SAPMSRUV 3687/TCP","SAPMSTUV 3687/TCP","SAPMSDMA 3634/TCP","SAPMSD6A 3646/TCP","SAPMSPCM 364/TCP",_
		"SAPMSP3M 368/TCP","SAPMSPMM 366/TCP","SAPMSZS1 362/TCP","SAPMSAGA 36XX/TCP","SAPMSCGB 3660/TCP","SAPMSD01 3610/TCP",_
		"SAPMSP01 3610/TCP","SAPMSPB1 3681/TCP","SAPMSQ01 3610/TCP","SAPMSSX1 3610/TCP","SAPMSTB1 3645/TCP","SAPMSTS1 3611/TCP","SAPMSTSI 3695/TCP",_
		"SAPMSUB2 3680/TCP","SAPMSZB6 360/TCP","SAPMSZB4 3642/TCP","SAPMSDBC 3600/TCP","SAPMSABC 3600/TCP","SAPMSTCO 3606/TCP","SAPMSTRM 3644/TCP" ,_
		"SAPMSPP2 3609/TCP","SAPMSPQ2 3619/TCP", "SAPMSDUV 3687/TCP", "SAPMSPP2 360/TCP", "SAPMSDCA 3620/TCP", "SAPMSCR1 3610/TCP", "SAPMSTB1 3626/TCP",_
		"SAPMSCRD 3624/TCP", "SAPMSCGA 3686/TCP", "SAPMSCCA 3648/TCP", "SAPMSCSA 3626/TCP", "SAPMSC6A 3687/TCP", "SAPMSCMA 3610/TCP", "SAPMSCVD 3628/TCP",_
		"SAPMSCID 3644/TCP", "SAPMSSVD 3602/TCP", "SAPMSSBH 3600/TCP", "SAPMSDBH 3600/TCP", "SAPMSDXH 3602/TCP", "SAPMSABH 3600/TCP", "SAPMSPBH 3600/TCP",_
		"SAPMSPBH 361/TCP", "SAPMSRID 3676/TCP", "SAPMSZS9 3672/TCP", "SAPMSUS9 3668/TCP", "SAPMSPXH 3602/TCP", "SAPMSTVO 3654/TCP", "SAPMSPVO 3624/TCP",_
		"SAPMSCBA 3602/TCP", "SAPMSSBA 3606/TCP", "SAPMSNBA 3600/TCP", "SAPMSQBA 3610/TCP", "SAPMSPED 3650/TCP", "SAPMSA1A 3622/TCP", "SAPMSD1A 3620/TCP",_
		"SAPMSABA 3608/TCP", "SAPMSTBA 3614/TCP", "SAPMSPBA 3618/TCP", "SAPMSPXH 3601/TCP")
gsdservers = ARRAY("USMDLTDOWS126", "USMDLTDOWS127", "USMDLTDOWS128", "USMDLTDOWS129", "USMDLTDOWS134", "USMDLTDOWS143", "USMDLTDOWS144", "USMDLTDOWS145", "USMDLTDOWG618")
officeversions = ARRAY("9.0","10.0","11.0","12.0","14.0","15.0","16.0") 'ignore 7.0, 8.0. There is no 13.0
version = 740
packagefile = "740.exe"
packagename = "740"
debug = 0
logging = 1
dodel = 1
reboot = 0
currentstage = 0
force = 0
returncode = 0
silent = FALSE
doprogress = TRUE
local = FALSE
update = FALSE 'TRUE = ONLY update XML/INI files and do NOT upgrade the SAPGUI version
upgrade = FALSE 'Force upgrade no matter the version
sccm = FALSE
docopy = TRUE
ignorejava = FALSE
skipjavacheck = FALSE
skipdotnetcheck = FALSE
quieter = FALSE
ignoreoffice = FALSE
packageoveride = FALSE
filescopied = FALSE
waserror = FALSE
uninstall = FALSE
repair = FALSE
killprocs = TRUE 'TRUE = do not permit user to abort (and run on reboot)
fixtxsap = FALSE
installjava = FALSE
whack = FALSE
disablepersonas = TRUE
allowjava64 = TRUE

'exit code statuses
' -9 sapgui didn't install correctly
' -8 sapgui didn't install correctly, recommend rebooting
' -3 didnt install, will install on next reboot
' -2 user chose to quit
' -1 bad, something broke
'  0 no status (yet)
'  1 every-things good
'  2 uninstall ran
status = 0
Const good = 1
Const bad = -1

'perhaps one day, change the cmd calls to run minimized


'vars that need FSO
runningfrom = FSO.GetFile(Wscript.ScriptFullName).ParentFolder
desktopver = FSO.GetFile(Wscript.ScriptFullName).DateLastModified

For Each arg In args
	If UCase(arg) = "NOCOPY" Then
		docopy = FALSE
	ElseIf UCase(arg) = "LOCAL" Then
		local = TRUE
		loglocal = TRUE
	ElseIf UCase(arg) = "IGNOREOFFICE" Then
		ignoreoffice = TRUE
	ElseIf UCase(arg) = "UPDATE" Then
		update = TRUE
		upgrade = FALSE
	ElseIf UCase(arg) = "UPGRADE" Then
		upgrade = TRUE
	ElseIf UCase(arg) = "FIXHISTORY" Then
		fixhistory = TRUE
	ElseIf UCase(arg) = "WHACK" Then
		whack = TRUE
		uninstall = TRUE
		skipjavacheck = TRUE
		skipdotnetcheck = TRUE
	ElseIf UCase(arg) = "INSTALLJAVA" Then
		installjava = TRUE
	ElseIf UCase(arg) = "REPAIR" Then
		repair = TRUE
		skipjavacheck = TRUE
		skipdotnetcheck = TRUE
	ElseIf UCase(arg) = "QUIETER" Then
		quieter = TRUE
	ElseIf UCase(arg) = "SILENT" Then
		silent = TRUE
	ElseIf UCase(arg) = "SCCM" Then
		'used for installing the SAPGUI from SCCM
		'kill procs
		'files installed from CWD folder
		'does not copy to temp
		'logs to program files\logs
		'kills procs
		'doesnt switch to no_nwbc installer as only 7xx.exe is in sccm dir (so 7xx.exe might fail)
		'quieter
		sccm = TRUE
		local = TRUE
		docopy = FALSE
		'doprogress = FALSE
		loglocal = TRUE
		'silent = TRUE
		quieter = TRUE
		'ignorejava = TRUE
		ignoreoffice = TRUE
		sourcedir = runningfrom
		killprocs = TRUE
		'upgrade = TRUE
	ElseIf UCase(arg) = "LOGLOCAL" Then
		loglocal = TRUE
	ElseIf UCase(arg) = "NONWBC" Then
		nonwbc = TRUE
	ElseIf UCase(arg) = "NOPROGRESS" Then
		doprogress = FALSE
	ElseIf UCase(arg) = "IGNOREJAVA" Then
		ignorejava = TRUE
	ElseIf UCase(arg) = "SKIPJAVACHECK" Then
		skipjavacheck = TRUE
	ElseIf UCase(arg) = "SKIPDOTNETCHECK" Then
		skipdotnetcheck = TRUE
	ElseIf UCase(arg) = "NETWORK" Then
		network = TRUE
	ElseIf UCase(arg) = "LOGNETWORK" Then
		loglocal = FALSE
	ElseIf UCase(arg) = "NOKILLPROCS" Then
		killprocs = FALSE
	ElseIf UCase(arg) = "NODOWDOTCOM" Then
		dotdowdotcom = ""	
	ElseIf UCase(arg) = "UNINSTALL" Then
		'this won't install after the uninstall
		uninstall = TRUE
		skipjavacheck = TRUE
		skipdotnetcheck = TRUE
	ElseIf UCase(arg) = "740" Then
		version = 740
		packagefile = "740.exe"
		packagename = "740"
	ElseIf UCase(arg) = "730" Then
		version = 730
		packagefile = "730.exe"
		packagename = "730"
	ElseIf UCase(arg) = "720" Then
		version = 720
		packagefile = "720.exe"
		packagename = "ALL_WITH_SERVICE"
	ElseIf UCase(arg) = "CWD" Then
		'run files straight from CWD. do not copy them anywhere
		sourcedir = runningfrom
		local = TRUE
		docopy = FALSE
	ElseIf left(UCase(arg),7) = "SOURCE=" Then
		'run files straight from source. do not copy them anywhere
		sourcedir = right(UCase(arg),len(UCase(arg))-7)
		local = TRUE
		docopy = FALSE
	ElseIf left(UCase(arg),8) = "DESTDIR=" Then
		'copy files from source to dest
		destdir = right(UCase(arg),len(UCase(arg))-8)
		docopy = TRUE
	ElseIf left(UCase(arg),12) = "PACKAGEFILE=" Then
		packagefile = right(UCase(arg),len(UCase(arg))-12)
		packageoveride = TRUE
	ElseIf left(UCase(arg),8) = "PACKAGE=" Then
		packagename = right(UCase(arg),len(UCase(arg))-8)
		packageoveride = TRUE
	ElseIf UCase(arg) = "COPY" Then
		docopy = TRUE
	ElseIf UCase(arg) = "NOTHING" Then
		msgbox "Doing Nothing!"
		exitcode(good)
	End If
	'msgbox(arg)
	allargs = allargs & arg & " "
Next

'some workstations can not access fully qualified .dow.com
If NOT local AND NOT dotdowdotcom = "" Then
	'work out if they can access FQDN
	If FSO.FolderExists("\\downeavm061.dow.com\Installation_Server\CustomerFilesDOW") Then
		'can reach it with .dow.com
		dotdowdotcom = ".dow.com"
	Else
		'can't reach it with .dow.com
		dotdowdotcom = ""
	End If
End If

masterscript = "\\downeavm061" & dotdowdotcom & "\Installation_Server\Scripts\install.2.vbs"

If local AND sourcedir = "" Then
	If silent or quieter Then
		'no msgbox
	Else
		'msgbox("LOCAL specified." & VbCrLf & "Using current directory as source of files.")
	End If
	sourcedir = runningfrom
	docopy = FALSE
End If

If (packagefile = "") AND (packagename <> "") Then
	msgbox("You must specify both a package and packagefile")
End If
If (packagefile <> "") AND (packagename = "") Then
	msgbox("You must specify both a package and packagefile")
End If

'authenticate to downeavm061
'if user can not access with current downeavm061 map
'so it unmaps all and then redoes and asks for credentials
If local OR silent Then 
	'do nothing
Else
	If NOT FSO.FolderExists("\\downeavm061" & dotdowdotcom & "\Installation_Server\CustomerFilesDOW") Then
		'can't reach it with .dow.com
		If NOT FSO.FolderExists("\\downeavm061\Installation_Server\CustomerFilesDOW") Then
			'can't reach server at all. share messed up?
			oShell.run "cmd /c net use u: /d /y", 1, TRUE
			oShell.run "cmd /c net use \\downeavm061.dow.com\Installation_Server /d /y", 1, TRUE
			oShell.run "cmd /c net use \\downeavm061.dow.com /d /y", 1, TRUE
			oShell.run "cmd /c net use \\downeavm061\Installation_Server /d /y", 1, TRUE
			oShell.run "cmd /c net use \\downeavm061 /d /y", 1, TRUE
			'oShell.run "cmd /c net config server /autodisconnect:-1", 1, TRUE
			oShell.run "cmd /c net config server /autodisconnect:60", 1, TRUE
			'can del with cmdkey.exe
			msgbox "If asked to authenticate to server DOWNEAVM061" & VbCrLf & "Enter your DOW\userid and Password." & VbCrLf & "Example:   DOW\N004563"
			'oShell.run "cmd /c net use u: \\downeavm061" & dotdowdotcom & "\Installation_Server /savecred /persistent:yes", 1, TRUE
			oShell.run "cmd /c net use \\downeavm061" & dotdowdotcom & "\Installation_Server", 1, TRUE
		Else
			'can reach server, but not with .dow.com
			dotdowdotcom = ""
		End If
	Else
		'can reach downeavm061.dow.com
	End If
End If

If sourcedir = "" Then
	sourcedir = "\\downeavm061" & dotdowdotcom & "\Installation_Server\CustomerFilesDOW"
	sourcedirfiles = "\\downeavm061" & dotdowdotcom & "\Installation_Server\" & version
	sourcedirsapguiexe = "\\downeavm061" & dotdowdotcom & "\Installation_Server\" & version & "\sapgui"
	sourcedirexe = "\\downeavm061" & dotdowdotcom & "\Installation_Server\ServerFiles"
	scourcedirscripts = "\\downeavm061" & dotdowdotcom & "\Installation_Server\Scripts"                               
Else
	'sourcedir is on command line
	sourcedirsapguiexe = sourcedir
	sourcedirexe = sourcedir
	scourcedirscripts = sourcedir                              
End If

If destdir = "" Then
	'place to copy files too was not on the command line
	destdir = temppath & "\SAPGUI.install"
End If
createfolder(destdir)


'set up log folder
If NOT loglocal Then
	logfolder = "\\downeavm061" & dotdowdotcom & "\logs"
End If


' start logging!
If logging Then
	'can't use logging functions here
	'msgbox("checking for network folder")
	If NOT FSO.FolderExists(logfolder) Then
		loglocal = TRUE
		If FSO.FolderExists("C:\Program Files\Logs") Then
			'this IF will exist when the first condition is met
			logfolder = "C:\Program Files\Logs"
		ElseIf FSO.FolderExists("C:\Program Files (x86)\Logs") Then
			logfolder = "C:\Program Files (x86)\Logs"
		ElseIf FSO.FolderExists(userprofile & "\Documents") Then
			logfolder = userprofile & "\Documents"
		ElseIf FSO.FolderExists(publicfolder & "\Documents") Then
			logfolder = publicfolder & "\Documents"
		ElseIf FSO.FolderExists(temppath) Then
			logfolder = temppath
		ElseIf FSO.FolderExists(tmppath) Then
			logfolder = tmppath
		Else
			'can't find somewhere to log!
			logging = FALSE
		End If
	End If
	If logging Then
		'msgbox("logfolder: " & logfolder)
		on error resume next
		If update Then 
			set logfile = FSO.OpenTextFile(logfolder & "\SAPGUI." & version & ".update." & userName & ".log" , ForAppending, True)
		Else
			set logfile = FSO.OpenTextFile(logfolder & "\SAPGUI." & version & ".install." & userName & ".log" , ForAppending, True)
		End If
		If Err <> 0 Then
			'unable to open log file, disable logging
			logging = FALSE
		End If
		on error goto 0
	End If
End If

log("*****************************************************************************************************************************")
log("Starting Script from " & runningfrom)
log("Being run by: " & userName)
GetTimeZoneOffset()
log("THIS ver: " & desktopver)

'check if this is the latest version
If NOT local Then
	log("Checking for latest version of this script")
	If exists (masterscript,0) Then
		nwver = FSO.GetFile(masterscript).DateLastModified
	End If
	log("Network ver: " & nwver)
	If desktopver < nwver Then
		log("Newer version is available on the network")
		If NOT silent Then
			log("Waiting for user: newer version on network")
			msgbox("There is a newer version of this script available" & VbCrLf & "Please follow the install instructions and re-run install.1.bat from:" & VbCrLf & "\\downeavm061.dow.com\installation_server\Scripts")
			log("Waiting for user: abort?")
			abort = msgbox("Continue to run this script anyway?" & VbCrLf & "It may work, but you may also have issues" , 4)
			'yes = 6, no = 7
			if abort = 7 then 
				log("User choose to abort script")
				exitcode(-2)
			End If 
		End If
		log("User choose to run script anyway")
	Else
		log("Running current version of the script. Nice work")
	End If

	'Exit if running from the network, as UAC prevents script from doing most of what is needed
	If left(runningfrom,1) = "C" OR left(runningfrom,1) = "D" OR left(runningfrom,1) = "E" Then
		'cool, its either C: or D: or E:
		log("Running from a local PC drive")
	Else
		'not running from either C: or D: or E:
		log("Not running from a local PC drive.")
		If userName = "N003352" OR network Then
		' its ok to run from the network, or as me
			log("Either N003352, or network is TRUE")
		Else
			'running from network, or citrix share, and not as me
			If instr(runningfrom,"$") > 0 Then
				'likely a citrix server
				'ie \\usnt193\faas001$\Desktop
				'If NOT silent Then
					'msgbox("Please move install.2.vbs to a physical drive." & VbCrLf & "It MUST run from a drive starting with C: or D:" & VbCrLf & "Run (as Admin) install.2.vbs from the new location." & VbCrLf & "This script will now quit.")
				'End If
				log("$ is in runningfrom, running from a Citrix server? Not quitting.")
			Else
				'$ isn't in runningfrom
				log("$ isn't in runningfrom")
				If NOT silent Then
					log("Waiting for user: not running from PC")
					log("Telling user to run from a local drive and quitting.")
					msgbox("You are not running this script from your DESKTOP!" & VbCrLf & "Please check the instructions you were provided" & VbCrLf & "You MUST run install.2.vbs from your DESKTOP" & VbCrLf & "This script will now quit.")
				End If
				exitcode(-2)
			End If
		End If
	End If
End If

log(".dow.com:" & dotdowdotcom)
checkcomputer()
osver()

'Here is a phony progress bar that uses an IE Window
'in code, the colon acts as a line feed
If doprogress = TRUE then
	log("Creating IE progress window")
	On Error Resume Next
	window.navigate2 "about:blank" : window.width = 600 : window.height = 100 : window.toolbar = false : window.menubar = false : window.statusbar = false : window.visible = True
	If Err <> 0 Then
		doprogress = FALSE
	Else
		window.document.write "<font color=blue>"
	End If
	On Error Goto 0
End If

progress 1,"Initialization"

'Welcome!
progress 2,"Welcome Instructions"
If NOT silent Then
	log("Waiting for user: welcome instructions")
	MsgBox "SAPGUI Install & Update." & VbCrLf & "Script will run in the background." & VbCrLf & "Script may run for up to 60 minutes." & VbCrLf & "Wait for the 'complete' pop-up indicating the script is complete." & VbCrLf & "Contact the GSD if you have issues."
End If

If citrix Then
	log("CITRIX Server")
	If NOT silent Then
		log("Waiting for user: CITRIX Server")
		answer = MsgBox("CITRIX Server detected" & VbCrLf & "Did you start this script with change user /install?", 4)
		If answer <> 6 Then
			log("Didn't start script with change user /install. Quitting")
			exitcode(bad)
		Else
			log("Started script with change user /install. Continuing")
		End If
	End if
End If	

'command line args
progress 1,"Checking args"
If WScript.Arguments.Count = 0 Then
	log("No command line arguments")
Else
	y = 0
	For Each arg In args
		log("Argument : " & y & " : " & arg)
		y = y + 1
	Next
End If

'folders
log("sourcedir: " & sourcedir)
log("destdir: " & destdir)
log("sourcedirsapguiexe: " & sourcedirsapguiexe)
log("sourcedirexe: " & sourcedirexe)
log("scourcedirscripts: " & scourcedirscripts)
log("logfolder: " & logfolder)
log("masterscript: " & masterscript)
log("thisscript: " & thisscript)
log("docopy: " & docopy)


'Has Admin rights?
progress 2,"Admin Rights?"
If CSI_IsAdmin Then 
	log("Running with ADMIN rights")
Else
	log("NOT running with ADMIN rights, or not sure. <================= BIG PROBLEM")
	If silent or SCCM Then
		'fail quietly
		log("silent or SCCM. Quitting script")
		exitcode(bad)
	Else
		log("Waiting for user: No Admin rights")
		answer = msgbox("No Admin rights." & VbCrLf & "The script will likely fail." & VbCrLf & "Next time, right click this script and select:" & VbCrLf & "'Run as Administrator'" & VbCrLf & "Do you wish to continue anyway?" & VbCrLf & "(not recommended)", 4)
		If answer = 6 Then
			log("User wants to run without admin rights. Script might fail")
		Else
			log("User doesn't wants to run without admin rights. Quitting")
			log("Waiting for user: Click OK to exit script")
			msgbox("Click OK to quit script." & VbCrLf & "Next time, right click this script and select:" & VbCrLf & "'Run as Administrator'")
			exitcode(-2)
		End If
	End If
End If

IsRebootPending()

Function IsRebootPending()
	On Error Resume Next
	Dim objReg
	Dim strKeyPath, strKeyValueName, strValue, arrValues
	Dim arrSubKeys, SubKey
	Dim blnRebootPending, blnPendingFileRenameOperations, blnAutoUpdate

	blnRebootPending = False
	blnPendingFileRenameOperations = False
	blnAutoUpdate = False

	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

	' Check Pending File Rename for reboot
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager"
	strKeyValueName = "PendingFileRenameOperations"
	objReg.GetMultiStringValue HKEY_LOCAL_MACHINE, strKeyPath, strKeyValueName, arrValues
	' Reboot needed if any values in the PendingFileRenameOperations key
	If Not IsNull(arrValues) Then
		For Each strValue In arrValues
			blnRebootPending = True
			blnPendingFileRenameOperations = True
		Next
	End If

	' Check Auto Update for reboot
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
	objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	If Not IsNull(arrSubKeys) Then
		For Each Subkey in arrSubKeys
			If SubKey = "RebootRequired" then
			blnRebootPending = True
			blnAutoUpdate = True
		End If
	  Next
	End If
	Set objReg = Nothing

	log("Reboot Pending is: " & blnRebootPending)
	log("Pending File Rename Operations: " & blnPendingFileRenameOperations)
	log("Pending Auto Update: " & blnAutoUpdate)
	IsRebootPending = blnRebootPending
End Function

'determine domain
progress 2,"Checking DOMAIN"
If UserDomain = "" OR UserDomain = "%USERDOMAIN%" Then
	UserDomain = wshNetwork.UserDomain
	If UserDomain = "" OR UserDomain = "%USERDOMAIN%" Then
		UserDomain = "UNKNOWN"
	End If
End If


'check computer name and domain
progress 3,"Checking PC Name"
If (left(pccomputername,3)) = "EWS" then
	computertype = "EWS"
ElseIf (left(pccomputername,3)) = "EW5" then
	computertype = "EWS"	
ElseIf (left(pccomputername,3)) = "EW6" then
	computertype = "EWS"
ElseIf (left(pccomputername,3)) = "EX5" then
	computertype = "EWS"
ElseIf (left(pccomputername,3)) = "EXS" then
	computertype = "EWS"
ElseIf (left(pccomputername,3)) = "MW7" then
	computertype = "ACN"
ElseIf (userdnsdomain) = "DIR.SVC.ACCENTURE.COM" then
	computertype = "ACN"
ElseIf (UserDomain) = "DIR" then
	computertype = "ACN"
ElseIf (UserDomain) = "SDR" then
	computertype = "SDR"
Else 
	computertype = "OTHER"
End if

log("Domain: " & UserDomain)
log("DNS: " & userdnsdomain)
log("Computer type: " & computertype)
log("Computer name: " & pccomputername)

' once there was this guy who, didnt have an appdata environment var.
progress 2,"Checking APPDATA"
If PathToAppData = "%appdata%" OR PathToAppData = "" Then
	log("AppData missing")
	If userprofile = "%userprofile%" OR userprofile = "" Then
		log("Userprofile missing")
		userprofile = systemdrive & "\users\" & userName
		log("New Userprofile: " & userprofile)
	End If
	If osystem = "XP" OR osystem = "WIN2003" Then
		PathToAppData = userprofile & "\Application Data"
	Else	
		PathToAppData = userprofile & "\AppData\Roaming"
	End If
	log("New AppData: " & PathToAppData)
End IF
log("Userprofile: " & userprofile)
log("Roaming AppData (PathToAppData): " & PathToAppData)
log("Local AppData (localappdata): " & localappdata)


progress 2,"Setting file locations"
' determine where ini files and shortcuts should go 
' shortcuts point to: %USERPROFILE%\AppData\Roaming\SAP\Common\
'
' %appdata%
' XP:	C:\Documents and Settings\All Users\Application Data
' XP:   C:\Documents and Settings\jay.y.zhang\Application Data
' XP:   C:\Documents and Settings\jay.y.zhang\Application Data\SAP\
' XP:   C:\Documents and Settings\jay.y.zhang\Application Data\SAP\Common
' XP:   C:\Documents and Settings\jay.y.zhang\Application Data\SAP\NWBC
' Win7: C:\Users\manoj.gopanapalli\AppData\Roaming
' Win7: C:\Users\manoj.gopanapalli\AppData\Roaming\SAP
' Win7: C:\Users\manoj.gopanapalli\AppData\Roaming\SAP\Common
' Win7: C:\Users\manoj.gopanapalli\AppData\Roaming\SAP\NWBC
'
'NWBC:
'C:\Program Files (x86)\SAP\NWBC40 (on 64–bit OS)
'C:\Program Files\SAP\NWBC40 (on 32–bit OS)
'A user’s personal settings are located in %APPDATA%\SAP\NWBC for Windows 7, Windows Vista, and Windows XP.
'Administrator configuration, such as predefined system connections or user settings, is located in the following folder:
'%ALLUSERSPROFILE%\SAP\NWBC for Windows 7 and Windows Vista
'%ALLUSERSPROFILE%\Application Data\SAP\NWBC for Windows XP
If osystem = "XP" OR osystem = "VISTA" Then
	'not shared, old folders
	inifilesparent = PathToAppData & "\SAP"
	inifiles = PathToAppData & "\SAP\Common"
	nwbcfiles = PathToAppData & "\SAP\NWBC"
	shortcutfiles = "C:\Documents and Settings\All Users\Start Menu\Programs\SAP Front End"
	publicdesktop = "C:\Documents and Settings\All Users\Desktop"
	publicdocuments = "C:\Documents and Settings\All Users\Documents"
ElseIf osystem = "WIN2000" OR osystem = "WIN2003" Then
	'shared, old folders
	inifilesparent = "C:\Documents and Settings\All Users\Application Data\SAP"
	inifiles = "C:\Documents and Settings\All Users\Application Data\SAP\Common" 'common location
	nwbcfiles = "C:\Documents and Settings\All Users\Application Data\SAP\NWBC" 'common location
	shortcutfiles = "C:\Documents and Settings\All Users\Start Menu\Programs\SAP Front End"
	publicdesktop = "C:\Documents and Settings\All Users\Desktop"
	publicdocuments = "C:\Documents and Settings\All Users\Documents"
ElseIf osystem = "WIN2008" OR osystem = "WIN2012" Then
	'shared, new folders
	inifilesparent = "C:\ProgramData\SAP"
	inifiles = "C:\ProgramData\SAP\Common" 'common location
	nwbcfiles = "C:\ProgramData\SAP\NWBC" 'common location
	shortcutfiles = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End"
	'publicdesktop = publicfolder & "\Public Desktop"
	publicdesktop = publicfolder & "\Desktop"
	publicdocuments = publicfolder & "\Documents"
Else
	'fallback to WIN7, WIN8 defaults
	'not shared, new folders
	inifilesparent = PathToAppData & "\SAP"
	inifiles = PathToAppData & "\SAP\Common"
	nwbcfiles = PathToAppData & "\SAP\NWBC"
	shortcutfiles = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End"
	'publicdesktop = publicfolder & "\Public Desktop"
	publicdesktop = publicfolder & "\Desktop"
	publicdocuments = publicfolder & "\Documents"
End If
log("inifilesparent: " & inifilesparent)
log("inifiles: " & inifiles)
log("shortcutfiles: " & shortcutfiles)

	
'work out if x32 or x64
progress 2,"Checking x32 or x64"
If KeyExists("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE") Then
	OsType = ReadKey("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
End If
If OsType = "x86" then
	OSBIT = 32
Elseif OsType = "AMD64" then
	OSBIT = 64
End If
If OSBIT = "" Then
	If PROCESSORARCHITECTURE = "x86" then
		OSBIT = 32
	Elseif PROCESSORARCHITECTURE = "AMD64" then
		OSBIT = 64
	Else 
		OSBIT = 32
	End If
End IF	
log("OS: " & OSBIT)

'check .NET
progress 2,"Checking for .NET registry"
log("Checking for .NET in registry")
If NOT skipdotnetcheck Then
	If NOT dotnet Then
		'not installed
		log(".NET 2.0 to 3.5 is NOT installed")
		If osystem = "WIN8" OR osystem = "WIN81" Then
			'try to trigger feature on demand
			'msgbox("Click OK to install .NET 3.5" & VbCrLf & "(Required for the SAPGUI)")
			'RC = runthis(windowsdir & "\system32\dism.exe", "/Online /Enable-Feature /FeatureName:NetFx3 /All", TRUE)
			'If RC <> 0 Then
			'	msgbox(".NET 3.5 did not install successfully" & VbCrLf & "Please contact the GSD" & VbCrLf & "Script will now quit")
			'	exitcode(bad)
			'End If
			log("Waiting for user: NO .NET 3.5")
			msgbox("Enable '.NET Framework 3.5' from" & VbCrLf & "'Turn Windows features on or off'" & VbCrLf & "Click OK once it is enabled")
			If NOT dotnet Then
				log("Waiting for user: STILL NO .NET 3.5")
				msgbox(".NET 3.5 is not installed." & VbCrLf & "Re-run this script once it is")
				exitcode(bad)
			End If
		Else
			log("Waiting for user: NO .NET 3.5")
			msgbox("SAPGUI needs .NET 3.5 installed" & VbCrLf & "Install .NET 3.5 from:" & VbCrLf & "https://www.microsoft.com/en-us/download/" & VbCrLf & "Click OK once installed")
			If NOT dotnet Then
				log("Waiting for user: STILL NO .NET 3.5")
				msgbox(".NET 3.5 is not installed." & VbCrLf & "Re-run this script once it is")
				exitcode(bad)
			End If
		End If
	End If
	'.net 2.0.50727or higher is installed
	log(".NET 2.0 to 3.5 is installed")
Else
	log("skipdotnetcheck set. Not checking .NET")
End If

Function dotnet()
	'Full install
	'Framework Version  Registry Key
	'1.0                HKLM\Software\Microsoft\.NETFramework\Policy\v1.0\3705 
	'1.1                HKLM\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322\Install 
	'2.0                HKLM\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727\Install 
	'3.0                HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.0\Setup\InstallSuccess 
	'3.5                HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.5\Install 
	'4.0 Client Profile HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Client\Install - doesnt contain 2.0->3.5
	'4.0 Full Profile   HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Install - doesnt contain 2.0->3.5
	'4.5				HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Release - doesnt contain 2.0->3.5
	'win8				HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4.0\Client - doesnt contain 2.0->3.5
	'Support Packs
	'1.0                HKLM\Software\Microsoft\Active Setup\Installed Components\{78705f0d-e8db-4b2d-8193-982bdda15ecd}\Version 
	'1.0[1]             HKLM\Software\Microsoft\Active Setup\Installed Components\{FDC11A6F-17D1-48f9-9EA3-9051954BAA24}\Version 
	'1.1                HKLM\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322\SP 
	'2.0                HKLM\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727\SP 
	'3.0                HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.0\InstallSuccess 
	'3.0                HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.0\SP 
	'3.5                HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.5\SP 
	'4.0 Client Profile HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Client\Servicing - doesnt contain 2.0->3.5
	'4.0 Full Profile   HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Servicing - doesnt contain 2.0->3.5

	'"Install"=dword:00000001
	log("Checking .NET 2.0 to 3.5")
	If ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727\Install") > 0 OR _
		ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727\SP\Install") > 0 OR _
		ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.0\Setup\InstallSuccess") > 0 OR _
		ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.0\SP") > 0 OR _
		ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.0\InstallSuccess ") > 0 OR _
		ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.5\Install") > 0 OR _
		ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.5\SP") > 0 Then
		'ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Client\Install") = 1 OR _
		'ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Client\Servicing") = 1 OR _
		'ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Install") = 1 OR _
		'ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Servicing") = 1 OR _
		'ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Release") >= 378389 Then 
		dotnet = TRUE
	End If
	log("Checked for .NET 2.0 to 3.5")
	log("dotnet: " & dotnet)
	log("TESTING: Just checking .NET 4.0 and higher")
	'needed for NWBC5.0 
	'https://www.microsoft.com/en-us/download/details.aspx?id=40779
	ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Client\Install") 
	ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Client\Servicing")
	log("TESTING: Just checking .NET 4.5 and higher")
	ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Install") 
	ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Servicing")
	netrelease = ReadKey("HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Release")
	Select Case netrelease
	Case 378389
		log("Installed: .NET Framework 4.5")
	Case 378675
		log("Installed: .NET Framework 4.5.1 installed with Windows 8.1 or Windows Server 2012 R2")
	Case 378758
		log("Installed: .NET Framework 4.5.1 installed on Windows 8, Windows 7 SP1, or Windows Vista SP")
	Case 379893
		log("Installed: .NET Framework 4.5.2")
	Case 393295
		log("Installed: .NET Framework 4.6")
	Case 393297
		log("Installed: .NET Framework 4.6")
	Case 394254
		log("Installed: .NET Framework 4.6.1")
	Case 394271
		log("Installed: .NET Framework 4.6.1")
	End Select
	log("Checked for .NET 4.0 and higher")
End Function

If installjava Then
	'exe is hardcoded. use in exceptions
	'unstall all java. be careful: wmic product where "name like 'Java%'" call uninstall
	javainstall()
End If

'Not used
Function javainstall
	log("Installing Java...")
	If local Then
		RC = runthis(sourcedir & "\" & javafile, "SPONSORS=Disable AUTO_UPDATE=Disable WEB_ANALYTICS=Disable REBOOT=Disable", TRUE)
	Else
		RC = runthis(sourcedirexe & "\" & javafile, "SPONSORS=Disable AUTO_UPDATE=Disable WEB_ANALYTICS=Disable REBOOT=Disable", TRUE)
	End If
	If RC <> 0 Then
		log("Java didn't install correctly. Waiting for user")
		msgbox("Java didn't install correctly" & VbCrLf & "Install JAVA manually now" & VbCrLf & "Then click OK")
	End If
End Function

'If N003352, ask if check java
If userName = "N003352" and NOT uninstall and NOT skipjavacheck Then
	log("Waiting for Nick: Check JAVA?")
	answer = MsgBox("Check JAVA?", 4)
	If answer <> 6 Then
		skipjavacheck = TRUE
	End If
End If

'check java
log("ignorejava :" & ignorejava)
log("skipjavacheck :" & skipjavacheck)
progress 1,"Checking JAVA in registry"
If NOT skipjavacheck Then
	log("Checking for JAVA in registry")
	If OSBIT = 64 Then
		If KeyExists("HKLM\SOFTWARE\Wow6432Node\JavaSoft\Java Runtime Environment\CurrentVersion") Then
			jversion = ReadKey("HKLM\SOFTWARE\Wow6432Node\JavaSoft\Java Runtime Environment\CurrentVersion")
			log("Java Version in registry :" & jversion)
		Else
			log("Java not found in registry")
			nojava = TRUE
		End If
		log("TEST: Checking Java7FamilyVersion")
		KeyExists("HKLM\SOFTWARE\Wow6432Node\JavaSoft\Java Runtime Environment\Java7FamilyVersion")
		'Java7FamilyVersion = 1.7.0_51, even if 1.8 is installed!
	Else 
		If KeyExists("HKLM\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion") Then
			jversion = ReadKey("HKLM\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion")
			log("Java Version in registry :" & jversion)
		Else 
			log("Java not found in registry")
			nojava = TRUE
		End If
		log("TEST: Checking Java7FamilyVersion")
		KeyExists("HKLM\SOFTWARE\JavaSoft\Java Runtime Environment\Java7FamilyVersion")
	End If 
	If Err <> 0 Then 
		nojava = TRUE 
	End If
	If IsBlank(jversion,FALSE) Then 
		nojava = TRUE
	End If
	progress 0,"Checking JAVA in WMI"
	log("Running WMI Query for JAVA...")
	'Set Collection = objWMIService.ExecQuery("select Name,Version from Win32_Product where name like '%Java%' and version > 6")
	err.clear
	Set Collection = objWMIService.ExecQuery("select * from Win32_Product where name like '%Java%'")
	'Set Collection = objWMIService.ExecQuery("select * from Win32_Product where name like '%Java%' and NOT like '%updater%'") 'dont work
	returncode = err.number
	log("WMI Query completed with RC: " & returncode)
	If isnull(Collection) Then
		log("Collection is null")
	Else
		log("Collection is not null")
	End If
	'isblank(Collection) 'can't do this 
	log("Going to get count of collection")
	count = getCount(Collection)
	log("Got count of collection")
	If isnull(count) Then
	'If isblank(count) Then
		log("Count of Collection is null")
	Else
		log("Count of Collection is: " & count)
	End If
	If returncode = 0 AND count > 0 Then
		'WMI query was ok, found some entries
		log("Found " & Collection.count & " versions of JAVA in WMI")
		For Each Item In Collection
			log("WMI Found Java: " & item.name & " - " & item.version)
			'GOOD: WMI Found Java: Java 7 Update 25 - 7.0.250
			'BAD : WMI Found Java: Java 7 Update 25 (64-bit) - 7.0.250
			'2059424 - SAP GUI for Java: Requirements for Release 7.40 - min req 1.8.25, x86 or x64!
			If allowjava64 OR InStr(item.name,"64-bit") = 0 Then
				'64-bit not found. must be 32bit
				'log("JAVA is 32bit")
				'formatting
				'Javaver1 = replace(item.version,",","")
				'Javaver1 = replace(Javaver1,".","")
				'log("WMI Found Java Formatted: " & item.name & " - " & Javaver1)
				'Javaver1 = item.version
				If NOT IsBlank(item.version,FALSE) Then
					'replace , with .
					Javaver1 = replace(item.version,",",".")
					log("After , to . replace, Javaver1 becomes " & Javaver1)
					'Javaver1 = replace(Javaver1,".","")
					If Javaver1 > Javaver Then
						Log("Javaver1 is newer then Javaver. Updating Javaver")
						Javaver = Javaver1 
					End If
				End If
			Else
				log("Ignoring as JAVA is 64bit")
			End If
		Next
	Else
		log("returncode = 0 but count not greater than 0")
		log("meaning WMI ran, but returned 0 or -1 records")
	End If
	'returncode = err.number
	log("No other versions of Java found")
	progress 0,"Formatting JAVA version field"
	If returncode <> 0 Then 
		nojava = TRUE 
	End If
	log("Highest JAVA version found (Javaver): " & Javaver)
	If IsBlank(Javaver,FALSE) Then 
		'not found
		nojava = TRUE 
	Else
		'found java of at some version
		If instr(Javaver,",") > 0 Then 
			javasplit = split(Javaver,",")
		Else
			javasplit = split(Javaver,".")
		End If
		'log("javasplit: " & javasplit) 'cant do this
		'log("javasplit3: " & javasplit(UBound(javasplit)))
		If left(Javaver,1) < 6 then
			'Under 1.6_xxx BAD
			log("JAVA is 1.5 or lower - bad")
			nojava = TRUE
		ElseIf left(Javaver,1) = "6" then
			'1.6_xxx Only good if over 110
			If javasplit(UBound(javasplit)) < 110 then 
				log("JAVA is 1.6, but under 110 - bad")
				nojava = TRUE
			Else 
				log("JAVA is 1.6, over 110 - good")
				nojava = FALSE
			End If
		ElseIf left(Javaver,1) = "7" then
			'1.7_xxx Only good if over 40
			'doesnt work with 4 digit java's. ie 8.0.730.2
			If javasplit(UBound(javasplit)) < 40 then 
				log("1.7, but under 40 - bad")
				nojava = TRUE 
			Else 
				log("1.7, over 40 - good")
				nojava = FALSE
			End If
		ElseIf left(Javaver,1) > 7 then
			'1.8_xxx or higher GOOD
			log("1.8 or higher, good")
			nojava = FALSE
		Else
			'unknown version
			log("unknown version - bad")
			nojava = TRUE 
		End If
	End If
	progress 0,"Is JAVA above minimum?"
	log("jversion :" & jversion)
	log("Javaver :" & Javaver)
	If NOT nojava Then
		log("left1 Javaver :" & left(Javaver,1))
		log("last of javasplit :" & javasplit(UBound(javasplit)))
	End If	
	log("nojava :" & nojava)
	'2003 Doesn't find anything in WMI...
	'If os.Caption = "Microsoft(R) Windows(R) Server 2003, Enterprise Edition" Then
	'	nojava = FALSE
	'	log("Server is Microsoft(R) Windows(R) Server 2003, Enterprise Edition")
	'Else 
	'	log("IS NOT Microsoft(R) Windows(R) Server 2003, Enterprise Edition")
	'End If
	If osystem = "WIN2003" AND (jversion = "1.7" OR jversion = "1.8") Then
		log("OS is WIN2003, and JAVA version is at least 1.7, saying java is ok")
		nojava = FALSE
		ignorejava = TRUE
	End If
	If ignorejava Then
		log("ignorejava set, continuing script regardless of above java check result")
	Else
		If nojava Then
			log("Failed Java check")
			log("Telling user to download JRE and quittin")
			log("Waiting for user: NO JAVA 1")
			msgbox("SAPGUI needs 32BIT Java Runtime Environment" & VbCrLf & "Install 32BIT JRE 1.8 from: " & VbCrLf & "http://java.com/en/download/manual.jsp" & VbCrLf & "MUST be 32BIT - NOT 64BIT")
			'msgbox("SAPGUI needs 32BIT Java Runtime Environment" & VbCrLf & "Going to copy then run JAVA 1.7 installer...") 
			'copyfile sourcedirexe730, "jre-7u25-windows-i586.exe", destdir, ""
			'numerr = oShell.Run(sourcedir & "\ALL_with_service.exe " & "/Package=" & """ALL with service""" & " /noDlg", 1, True)
			log("Waiting for user: NO JAVA 2")
			msgbox("Script will now exit." & VbCrLf & "Re-run this script once JAVA has been installed")
			exitcode(bad)
		else
			log("Java checked and is OK!")
		End If
	End If
Else
	log("skipjavacheck set. Didn't even check Java")
End If

'remove run key
delkey "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run","SAPGUI 7.40 Install"

'look for SAPGUI
progress 2,"Finding SAPGUI"
SAPGUIlocation()

If uninstall Then
	log("Deleting old NWBC keys")
	RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\DelNwbcOptions.reg"), FALSE)
	stopupdater()
	RC = uninstallsapgui()
	delshortcuts()
	delvirtualstore()
	delinifiles()
	deljunk()
	delshortcuts2()
	delfile (inifiles & "\SapLogonTree.xml")
	removeaddremove()
	If whack Then
		delall()
	End If
	exitcode(2)
End If

Function delall
	log("Uninstalling the ALL of the SAPGUI")
	RC = runthis(SAPSetupDir & "\setup\NwSapSetup.exe", "/uninstall /all /noDlg", TRUE)
	delfolder("C:\Documents and Settings\All Users\Start Menu\Programs\SAP Front End")
	delfolder("C:\Documents and Settings\All Users\Start Menu\Programs\Business Explorer")
	delfolder("C:\Program Files (x86)\SAP")
	delfolder("C:\Program Files\SAP")
	delfolder("C:\Program Files (x86)\Common Files\Sap Shared")
	delfolder("C:\Program Files\Common Files\Sap Shared")
	delfolder(shortcutfiles)
	delfolder(pathtoappdata & "\SAP")
	delfolder(localappdata & "\SAP")
	delfolder(temppath & "\sapsetup")
	delfolder(tmppath & "\sapsetup")
	delkey "HKEY_CURRENT_USER\Software\SAP", ""
	delkey "HKEY_CURRENT_USER\Software\Wow6432Node\SAP", ""
	delkey "HKEY_LOCAL_MACHINE\SOFTWARE\SAP", ""
	delkey "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP", ""
	runthis cmd, "/c REG IMPORT " & stringthis(sourcedir & "\DelAllSAPKeys.reg"), TRUE
End Function

'If N003352, ask if we should copy
progress 2,"Do Copy?"
If userName = "N003352" and docopy = TRUE Then
	log("Waiting for Nick: copy files?")
	answer = MsgBox("docopy = " & docopy & VbCrLf & "Copy Files?", 4)
	If answer = 6 Then
		docopy = TRUE
	Else
		docopy = FALSE
	End If
End If

'always need to copy files
copythefiles

progress 2,"Running Processes?"
allprocsdead = FALSE
Do Until allprocsdead = TRUE
	If CheckProcesses() Then
		'found processes, ask user
		dokillprocs
	Else
		'didnt find any processes. exit loop and continue
		allprocsdead = TRUE
	End If
	If username = "N003352" Then
		allprocsdead = TRUE
	End If
Loop

Function CheckProcesses
	'find running SAPGUI & SAPLOGON & SAPLGPAD & NWBC
	log("Checking for running processes...")
	CheckProcesses = FALSE
	Set service = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	For Each Process in Service.InstancesOf ("Win32_Process")
		If ucase(Process.Name) = "SAPGUI.EXE" or ucase(Process.Name) = "SAPLGPAD.EXE" or ucase(Process.Name) = "SAPLOGON.EXE" or ucase(Process.Name) = "NWBC.EXE" then
			log("Found running process: " & Process.Name)
			CheckProcesses = TRUE
			If ucase(Process.Name) = "SAPGUI.EXE" then runningproc = "SAPGUI" End If
			If ucase(Process.Name) = "SAPLOGON.EXE" then runningproc = "SAPLOGON" End If
			If ucase(Process.Name) = "SAPLGPAD.EXE" then runningproc = "SAPLGPAD" End If
			If ucase(Process.Name) = "NWBC.EXE" then runningproc = "NWBC" End If
		End If
	Next
End Function

Function isrunning(processname)
	isrunning = FALSE
	If instr(processname, "\") <> 0 Then
		'contains a \, so remove path
		'truncate...
	End If
	Set service = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	For Each Process in Service.InstancesOf ("Win32_Process")
		If ucase(Process.Name) = ucase(processname) then
			'running!
			log(processname & " is already running")
			isrunning = TRUE
		Else
			log(processname & " is NOT already running")
		End If
	Next
	Set service = nothing
End Function

Function dokillprocs
	log("Waiting for user. Kill procs or not?")
	If killprocs Then
		'do not permit user to abort (and run on reboot)
		answer = msgbox(runningproc & " is running." & VbCrLf & "Cannot update when " & runningproc & " is running." _
			& VbCrLf & "RETRY (after you have closed " & runningproc & ")." _
			& VbCrLf & "CANCEL and processes will be killed.", vbRetryCancel +vbExclamation)
	Else
		'permit user to abort and run on reboot
		answer = msgbox(runningproc & " is running." & VbCrLf & "Cannot update when " & runningproc & " is running." _
			& VbCrLf & "ABORT and run the script on next reboot." _
			& VbCrLf & "RETRY (after you have closed " & runningproc & ")." _
			& VbCrLf & "IGNORE and processes will be killed.", vbAbortRetryIgnore+vbExclamation)
	End If
	Select Case answer
	Case vbAbort
		log("User choose ABORT - run script after reboot")
		'cant do this is you are running from a server, or a location without files
		If left(thisscript,2) = "\\" Then
			log("Waiting for user. Running from a server, can't do onreboot. Kill Procs & continuing script?")
			answer2 = msgbox("This script is running from a network share." & VbCrLf & "Cannot schedule run on reboot." & VbCrLf & "Want this script to kill SAP processes and continue?" & VbCrLf & "(or cancel this script)" ,VbOKCancel)
			Select Case answer2
			Case vbOK
				log("User choose to kill procs and continue")
				Killprocesses()
			Case vbCancel
				log("User choose to quit script")
				exitcode(-2)
			End Select
		Else
			'not running form a server - good
			If update Then
				'script is only in update mode, NOT upgrade - so sapgui installer will not be called.
				setonreboot() 'will copy files
			Else
				'script might try an upgrade, but we didnt copy sapgui installer.
			End If
		End If
	Case vbRetry
		log("User choose to retry process check")
	Case vbIgnore
		log("User choose to kill procs")
		KillProcesses
	Case vbCancel
		log("User choose to kill procs")
		KillProcesses
	End Select
End Function

Function KillProcesses
	log("Killing procs")
	Set colProcessList = service.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'sapgui.exe' OR Name = 'saplogon.exe' OR Name = 'saplgpad.exe' OR Name = 'nwbc.exe'")
	For Each objProcess in colProcessList 
		'If NOT userName = "N003352" Then
			On Error Resume Next
			objProcess.Terminate() 
			log("Terminated RC: " & Err.Number & " for " & objProcess.Name)
			On Error Goto 0
			WScript.Sleep 1000
		'End If
	Next  
End Function

Function setonreboot
	''Replace(string,find,replacewith[,start[,count[,compare]]])
	'[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run]
	'"My VBS Script"="wscript.exe C:\\myscript.vbs"
	If left(thisscript,2) = "\\" Then
		'running from a server, can't do this
		log("Running from a server, can't do on reboot. Killing procs and continuing script")
		Killprocesses
	Else
		log("Quitting script. Setting script to run from registry on next reboot")
		Writekey "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\SAPGUI 7.40 Install", "wscript.exe " & stringthis(thisscript) & " " & allargs & " source=" & stringthis(destdir), "REG_SZ"
		exitcode(-3)
	End if
End Function

Function robocopythis(robocopysource,robocopydest,robocopyfiles,options)
	RC = 99
	If exists(windowsdir & "\system32\robocopy.exe",0) Then
		RRC = runthis(windowsdir & "\system32\robocopy.exe", """" & robocopysource & """" & " " & """" & robocopydest & """" & " " & robocopyfiles & " " & " /R:3 /W:1 /ZB /MT:8 /ETA /TEE /LOG+:%temp%\robocopy.log " & options, TRUE)
		Select case RRC
			case 16 
				log("Robocopy: ***FATAL ERROR*** : " & RRC)
				RC = RRC
			case 15
				log("Robocopy: OKCOPY + FAIL + MISMATCHES + XTRA : " & RRC)
				RC = RRC
			case 14
				log("Robocopy: FAIL + MISMATCHES + XTRA : " & RRC)
				RC = RRC
			case 13
				log("Robocopy: OKCOPY + FAIL + MISMATCHES : " & RRC)
				RC = RRC
			case 12
				log("Robocopy: FAIL + MISMATCHES : " & RRC)
				RC = RRC
			case 11
				log("Robocopy: OKCOPY + FAIL + XTRA : " & RRC)
				RC = RRC
			case 10	
				log("Robocopy: FAIL + XTRA : " & RRC)
				RC = RRC
			case 9
				log("Robocopy: OKCOPY + FAIL : " & RRC)
				RC = RRC
			case 8
				log("Robocopy: FAIL : " & RRC)
				RC = RRC
			case 7
				log("Robocopy: OKCOPY + MISMATCHES + XTRA : " & RRC)
				RC = RRC
			case 6
				log("Robocopy: MISMATCHES + XTRA : " & RRC)
				RC = RRC
			case 5
				log("Robocopy: OKCOPY + MISMATCHES : " & RRC)
				RC = RRC
			case 4
				log("Robocopy: MISMATCHES : " & RRC)
				RC = RRC
			case 3
				log("Robocopy: OKCOPY + XTRA : " & RRC)
				RC = 0
			case 2 
				log("Robocopy: XTRA : " & RRC)
				RC = 0
			case 1
				log("Robocopy: OKCOPY : " & RRC)
				RC = 0
			case 0
				log("Robocopy: No Change : " & RRC)
				RC = 0
			case else
				log("Robocopy: UNKNOWN ERROR : " & RRC)
				RC = 99
		End Select
	Else
		'use xcopy instead
		RC = runthis(windowsdir & "\system32\xcopy.exe", """" & sourcedir & "\" & robocopyfiles & """" & " " & """" & destdir & "\" & """" & " /Y /R /H", TRUE)
	End If
	robocopythis = RC
End Function
					
'copy files from source folder to destination folder 
'if not local & not sccm: %temp%\SAPGUI.Install
progress 4,"Checking source and dest folders"

Function copythefiles
	filescopied = FALSE
	If username = "N003352" Then
		If left(sourcedir,2) = "\\" Then
			answer = MsgBox("sourcedir = " & sourcedir & VbCrLf & "destdir = " & destdir & VbCrLf & "Change source to:" & VbCrLf & "%appdata%\Temp\SAPGUI.install ?", 4)
			If answer = 6 Then
				sourcedir = "C:\Users\n003352\AppData\Local\Temp\SAPGUI.install"
				log("Nick changed changed sourcedir to " & sourcedir)
			End If
		End If
	End If 
	If docopy Then
		If createfolder(destdir) = FALSE Then
			'not able to create dest folder
			log("Waiting for user: cannot create destdir")
			msgbox "Cannot access necessary files" & VbCrLf & "Exiting Script"
			exitcode(bad)
		End If
	End If
	'always get a count of files
	If exists(sourcedir,0) Then 
		set folder = FSO.getfolder(sourcedir & "\")
		set ABC = folder.files
		numfiles = ABC.count
		log("Source DIR Files : " & numfiles & " in " & sourcedir)
	End If
	If docopy Then
		If exists(destdir,0) Then
			If exists(sourcedir,0) Then 
				set folder = FSO.getfolder(destdir & "\")
				set ABC = folder.files
				numfiles = ABC.count
				log("Dest DIR Files : " & numfiles & " in " & destdir)
				'set folder = FSO.getfolder(sourcedir & "\")
				'set ABC = folder.files
				'numfiles = ABC.count
				'log("Source DIR Files : " & numfiles & " in " & sourcedir)
				'log("copying individual files if newer from " & sourcedir)
				log("Copying files from: " & sourcedir)
				log("Copying files to  : " & destdir)
				progress 3,"Copying many files, please sit tight....."
				log("Copying all files")
				RC = robocopythis(sourcedir,destdir,"*.*", "/PURGE")
				If RC <> 0 Then 
					log("Copy of MANY files failed")
					log("Waiting for user: cannot copy")
					msgbox "Cannot copy necessary files" & VbCrLf & "Exiting Script"
					DisplayErrorInfo
					filescopied = FALSE
					exitcode(bad)
				Else
					filescopied = TRUE
				End If
				log("Copied all files")
			Else
				log("Cannot access sourcedir: " & sourcedir)
				log("Waiting for user: cannot access")
				msgbox "Cannot access necessary files" & VbCrLf & "Exiting Script"
				filescopied = FALSE
				exitcode(bad)
			End If
		Else
			log("Cannot access destdir: " & destdir)
			log("Waiting for user: cannot access")
			msgbox "Cannot access necessary files" & VbCrLf & "Exiting Script"
			filescopied = FALSE
			exitcode(bad)
		End If
	Else
		log("docopy set to FALSE, not copying files")
		progress 0,"NOT copying many files to temp folder"
		filescopied = TRUE
	End If
	If docopy Then
		sourcedir = destdir
		log("DOCOPY is TRUE & Copy has completed. sourcedir changed to " & sourcedir)
	Else
		destdir = sourcedir
		log("DOCOPY is FALSE. No Copy Performed. sourcedir NOT changed. Still : " & sourcedir)
		log("DOCOPY is FALSE. destdir changed to: " & sourcedir)
	End If
End Function


'Function to get modifed date
Function mdate(file)
	If exists(file,0) Then 
		Set f = FSO.GetFile(file)
		mdate = f.DateLastModified
		log(file & " was last modified: " & mdate)
	End If
End Function


'Function to get file version
Function fileversion(file)
	newstring = ""
	foundvf = FALSE
	ProductVersion = 0
	If exists(file,0) Then
		newfile = FSO.GetFile(file)
		filename = FSO.GetFileName(newfile)
		filefolder = FSO.GetParentFolderName(newfile)
		set objFolder = objShell.Namespace(filefolder)
		set objFolderItem = objFolder.ParseName(filename)
		Dim arrHeaders(300)
		For i = 0 To 300
			arrHeaders(i) = objFolder.GetDetailsOf(objFolder.Items, i)
			'WScript.Echo i &"- " & arrHeaders(i) & ": " & objFolder.GetDetailsOf(objFolderItem, i)
			If lcase(arrHeaders(i))= "file version" or lcase(arrHeaders(i))= "dateiversion" Then
				'need to add all languages
				ProductVersion = objFolder.GetDetailsOf(objFolderItem, i)
				log("TRY1: Full Version Info: " & file & " : " & ProductVersion)
				foundvf = TRUE
				Exit For
			End If
		Next
		If NOT foundvf Then
			log("TRY1: Full Version Info Check FAILED")
		End If
		'if still blank, lets try another way
		If IsBlank(ProductVersion,FALSE) Then
			ProductVersion = FSO.GetFileVersion(file)
			log("TRY2: Full Version Info: " & file & " : " & ProductVersion)
		End If
		If NOT IsBlank(ProductVersion,FALSE) Then
			a=split(ProductVersion,".")
			for each x in a
				if len(x) < 2 then x = "0" & x
				newstring = newstring & x
			next
			ProductVersion = newstring
			log("Formatted Version:" & file & " : " & ProductVersion)
		Else
			ProductVersion = 0
		End If
	Else
		log("Can't find file, setting fileversion to 0 and skipping")
		ProductVersion = 0
	End If	
	fileversion = ProductVersion
End Function

If repair Then
	'needs to call nwsapsetup from a package, or the server. Can't use 740.exe
	log("repair set. repairing SAPGUI")
	'RC = runthis(SAPSetupDir & "\setup\NwSapSetup.exe", "/repair", TRUE) 'can't run from c:\prog files
	RC = runthis(sourcedirfiles & "\Setup\NwSapSetup.exe", "/repair", TRUE) 'didn't like this either.
	log("SAPGUI repaired, RC: " & RC)
End If

'set installedversion for local sapgui.exe
Function SetInstalledVersion(tempver)
	If NOT IsBlank(tempver,FALSE) Then
		InstalledVersion = left(tempver,3)	
	Else
		InstalledVersion = 0
	End If
	log("InstalledVersion, Local SAPGUI.EXE version is " & InstalledVersion)
End Function

'version check of SAPGUI.EXE
progress 2,"Checking local SAPGUI.EXE version"
'update  = if set, only updates INI files, will never upgrade
'upgrade = if set, upgrades the sapgui to a new release
'always get local sapgui ver
log("Checking local SAPGUI.EXE version")
SourceDateModified = fileversion(sapguidir & "\SAPgui\" & sProgram)
SetInstalledVersion(SourceDateModified)
If SourceDateModified = "" OR SourceDateModified = 0 Then
		log("Can't get date or version on local sapgui.exe")
		log("SourceDateModified setting to -1")
		SourceDateModified = -1
End If
If uninstall OR local OR upgrade OR update Then
	If uninstall OR update or upgrade Then
		'dont bother checking network file version, set local version to 0
		log("Not checking network version. uninstall OR local or upgrade or update is set")
		log("DestDateModified setting to 0")
		DestDateModified = 0
	End If
	If local Then
		'need to set DestDateModified based on version of local 740.exe
		packageDestDateModified = fileversion(sourcedirexe & "\" & packagefile)
		log(sourcedirexe & "\" & packagefile & " version (packageDestDateModified) : " & packageDestDateModified)
		'For i = LBound(packageinfo) To UBound(packageinfo)
		'	If packageinfo(x)
		'Next
		log("DestDateModified setting to 0, due to package version check not (yet) implemented.")
		DestDateModified = 0
	End If
Else
	progress 0,"Checking network SAPGUI.EXE version"
	log("Checking network SAPGUI.EXE version")
	DestDateModified = fileversion(sourcedirsapguiexe & "\" & sProgram)
End If
If DestDateModified = "" Then
	log("DestDateModified is empty, setting to 0")
	DestDateModified = 0
End If
log("SourceDateModified = " & SourceDateModified)
log("DestDateModified   = " & DestDateModified)
If NOT upgrade Then
	'upgrade wasn't specified, so work out if to upgrade
	If DestDateModified > SourceDateModified Then
		log("Network is newer")
		upgrade = TRUE
	Else
		log("Local is the same version, or newer")
		upgrade = FALSE
	End if
End If
If local and NOT update Then
	log("local is TRUE & update is FALSE. Forcing install and upgrade")
	upgrade = TRUE
End if

'override
'If SourceDateModified = "7300.3.9.8950" Then
	'log("OVERRIDE: Skipping install as local is 7300.3.9.8950 (which is current on the network)")
	'upgrade = FALSE
'End If

stopupdater()

log("upgrade: " & upgrade)
log("update: " & update)

'upgrade SAPGUI, if newer version is avail
progress 2,"Upgrade or not?"
If NOT upgrade Then
	If NOT (silent OR quieter) Then
		log("Waiting for user: SAPGUI is current")
		MsgBox "SAPGUI is current. Not upgrading. Just applying updates"
	End If
	log("SAPGUI is current. Not upgrading. Just applying updates")
	progress 8,"SAPGUI is current"
End If

If NOT upgrade AND NOT update Then
	' fix for users running this script manually
	log("Setting to EN")
	Writekey "HKEY_CURRENT_USER\Software\SAP\General\Language", "EN", "REG_SZ"
	If version = 740 Then
		'For 740 - restore correct saplogon position, layout & options
		log("Importing new SAPLOGON registry settings")
		RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\saplogon.reg"), TRUE)
		'temp
	End If 
End If

If upgrade AND NOT update Then
	log("upgrade set, update not set")
	If NOT (silent OR quieter) Then
		log("Waiting for user: SAPGUI is not current")
		If SourceDateModified = -1 Then
			'no local SAPGUI
			msgbox("No SAPGUI detected" & VbCrLf & "Installing the current version")
		Else
			msgbox("Old SAPGUI detected" & VbCrLf & "Removing Old SAPGUI, then" & VbCrLf & "Installing the current version")
		End If
	End If
	'removed uninstall 10/12/16
	'BLG Put uninstall back in 12/5/2017
	uninstallsapgui()
	'Force EN
	log("Setting to EN")
	Writekey "HKEY_CURRENT_USER\Software\SAP\General\Language", "EN", "REG_SZ"
	InstallSAPGUI()
	'del NWBC keys
	log("Deleting old NWBC keys")
	RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\DelNwbcOptions.reg"), TRUE)
	If version = 740 Then
		'For 740 - restore correct saplogon position, layout & options
		log("Importing new SAPLOGON registry settings")
		RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\saplogon.reg"), TRUE)
		'temp
	End If
	'dont work...
	'If OSBIT = 32 Then
	'	DelKey "HKEY_LOCAL_MACHINE\SOFTWARE\SAP\NWBC",""
	'Else
	'	DelKey "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP\NWBC",""
	'End If
	'
	'HKEY_CURRENT_USER\Software\SAP\SAPLogon
	'overrides
	'HKEY_USERS\S-1-5-21-1060284298-861567501-682003330-876675\Software\SAP\SAPLogon
	'For SAP Logon (saplogon.exe) the setting under HKCU has higher priority.
	'For SAP Logon Pad (saplgpad.exe) the setting under HKLM has higher priority.
Else 
	log("Not upgrade, or update. NOT installing SAPGUI")
End If
	
Function uninstallsapgui()
	progress 0,"Uninstall SAPGUI"
	log("Uninstalling 740 package only")
	'RC = runthis(SAPSetupDir & "\setup\NwSapSetup.exe", "/uninstall /all /noDlg", TRUE)
	RC = runthis(SAPSetupDir & "\setup\NwSapSetup.exe", "/uninstall /Package=740 /noDlg", TRUE)
	'or indivdually (all part of 740 package)
	'/Product="SAPGUI" -- sapgui
	'/Product="NWBCGUI" -- sapgui desktop icons & shortcut
	'/Product="SAPWUS" -- updater service
	'/Product="SCRIPTED" -- legacy text editor
	'/Product="NWBC50" -- nwbc 5.0
	'/Product="SapBI" -- business explorer
	'/Package="JNET" -- JNET - no longer in 740
	'/Product="GUIISHMED" -- i.s.h.med planning grid -- possibly once in 740?
	'/Product="KW" -- KW add on -- possibly once in 740?
	'/Product="SRX" -- JAWS -- was never in 740
	log("Uninstalled 740 package, RC: " & RC)
	log("NOT Del old SAP Keys")
	'runthis cmd, "/c REG IMPORT " & stringthis(sourcedir & "\DelAllSAPKeys.reg"), TRUE
	delfile (inifiles & "\SapLogonTree.xml")
	uninstallsapgui = RC
End Function

Function InstallSAPGUI()
	If NOT (silent OR quieter) Then
		log("Waiting for user: Copying installer")	
		MsgBox "Copying installer will take an additional 10-20 minutes" & VbCrLf & "Copy will run in the background."
	Else
		log("Copying installer")	
	End If
	progress 2,"SAPGUI is not current"
	waserror = FALSE
	'If exists(destdir,0) Then
		If exists(sourcedirexe,0) Then 
			log("Installing SAPGUI")
			If NOT packageoveride Then
				log("Looking for MS OFFICE")
				If NOT checkoffice OR nonwbc Then
					'didnt find MSOFFICE, or told to install the no nwbc package
					log("Switching to NO-NWBC package")
					packagefile = version & "_no_nwbc.exe"
					packagename = version & "_no_nwbc"					
				End If
			Else
				log("Package specified, packageoveride: " & packageoveride)
				log("Package specified, packagefile: " & packagefile)
				log("Package specified, packagename: " & packagename)
			End If
			If docopy = TRUE Then
				progress 0,"Copying installer, will take 5 to 30 minutes..."
				RC = robocopythis(sourcedirexe, destdir, packagefile,"")
				'copyfile sourcedirexe, packagefile, destdir, ""
				If RC <> 0 Then
					log("Copy of " & packagefile & " failed. Expect bad things")
				End If
			End If
			progress 2,"Running SAPGUI installer"
			If silent Then
				'numerr = runthis(sourcedir & "\730.exe", "/Package=" & """730""" & " /silent", TRUE)
				RC = runthis(sourcedir & "\" & packagefile, "/Package=" & packagename & " /silent", TRUE)
			Else
				'numerr = runthis(sourcedir & "\730.exe", "/Package=" & """730""" & " /noDlg", TRUE)
				RC = runthis(sourcedir & "\" & packagefile, "/Package=" & packagename & " /noDlg", TRUE)
			End if
			log(packagefile & " completed with return code: " & RC)
			progress 2,"Checking Installer return code"
			If RC = -1 Then
				' process failed to start or aborted
				status = -9
				log("SAPGUI Installer died, or failed to start. Install failed. No reboot needed.")
				waserror = TRUE
			Elseif RC = -1073741818 Then 
				' In-page I/O Error (never seen this)
				status = -8
				log("SAPGUI Installer died - In-page I/O Error. Install failed. Recommend rebooting before trying again.")
				waserror = TRUE
			Elseif RC = -1073741819 Then 
				' Access Violation
				status = -8
				log("SAPGUI Installer died - Access Violation. Install failed. Recommend rebooting before trying again.")
				waserror = TRUE	
			Elseif RC = 0 Then
				' 0  - Process ended without any errors detected.
				log("SAPGUI Installed OK, no reboot needed, but recommending it.")
				addaddremove()
				reboot = 1
			Elseif RC = 129 Then
				' 129 - Reboot is recommended.
				log("SAPGUI Installer: 129 - Reboot is recommended.")
				addaddremove()
				reboot = 1
			Elseif RC = 130 Then 
				' 130 - Reboot was forced.
				log("SAPGUI Installer: 130 - Reboot was forced.")
				addaddremove()
				reboot = 2
			Else
				log("SAPGUI Installer Didn't return -1, ZERO or 129 or 130")
				'either critical or non-critical failure. reboot recommended.
				reboot = 2
				If RC < 144 Then
					'critical failure
					waserror = TRUE
					status = -9
					Select Case RC
						Case 1
							log("SAPGUI Installer error: 1 - ???")
						Case 3
							log("SAPGUI Installer error: 3 - Another instance of SAPSetup is running")
						Case 4
							log("SAPGUI Installer error: 4 - LSH failed")
						Case 16
							log("SAPGUI Installer error: 16 - SAPSetup started on WTS without administrator privileges")
						Case 26
							log("SAPGUI Installer error: 26 - WTS is not in install mode")
						Case 27
							log("SAPGUI Installer error: 27 - An error occurred in COM")
						Case 48
							log("SAPGUI Installer error: 48 - General error")
						Case 65
							log("SAPGUI Installer error: 65 - Process <SapSmartDel.exe> did not finish in 60000 milliseconds.")
						Case 67
							log("SAPGUI Installer error: 67 - Installation is canceled by the user.")
						Case 68
							log("SAPGUI Installer error: 68 - Invalid patch")
						Case 69
							log("SAPGUI Installer error: 69 - Installation engine registration failed")
						Case 70
							log("SAPGUI Installer error: 70 - Invalid XML files or Pre-req not met.")
					End Select
				Else
					'non-critical fail
					Select Case RC	
						Case 144
							log("SAPGUI Installer error: 144 - Error report has been created.")
						Case 145
							log("SAPGUI Installer error: 145 - Error report has been created and reboot is recommended.")
						Case 146
							log("SAPGUI Installer error: 146 - Error report has been created and reboot is forced.")
					End Select
				End If
			End If
			progress 2,"Looking for SAPGUI"
			SAPGUIlocation()
			log("Post install SAPGUI version....")
			'SetInstalledVersion(fileversion(sapguidir & "\SAPgui\" & sProgram))
			SourceDateModified = fileversion(sapguidir & "\SAPgui\" & sProgram)
			SetInstalledVersion(SourceDateModified)
			If waserror AND exists(SAPSetupDir & "\LOGs\NwSapSetup.log",0) Then
				log("Saving NwSapSetup.log to log dir")
				copyfile SAPSetupDir & "\LOGs", "NwSapSetup.log", logfolder, userName & "." & "NwSapSetup.log"
				Set objReadFile = FSO.OpenTextFile(SAPSetupDir & "\LOGs\NwSapSetup.log", 1, False, -1)
				Do Until objReadFile.AtEndOfStream
					strLine = objReadFile.ReadLine
					strLine = trim(strLine)
					if instr(strLine, "Successfully saved Error-Report XML") Then
						strLine = trim(strLine)
						a = split(strLine, "'")
						for each x in a
							if left(x,1) = "C" OR left(x,1) = "D" Then
								strLine = x
								strLine = trim(strLine)
							End If
						next
						a = split(strLine, "\")
						for each x in a
							if left(x,14) = "SapSetupErrors" Then
								strLineFile = x
								strLineFile = trim(strLineFile)
							End If
						next
						log("Saving Errors XML to log dir")
						If exists(strLine,0) Then
							'have to fix this one day
							On Error Resume Next
							FSO.Copyfile strLine, logfolder & "\" & userName & "." & strLineFile & ".log", TRUE
							On Error Goto 0
						End If
						Exit Do
					End If
				Loop
			End if
			If NOT waserror Then
				'remove NWBC4
				'NwSapSetup.exe /product:JNet /uninstall [/nodlg | /silent]
				'removed the uninstall 10/12/16
				'log("Installed 7.40, removing NWBC40")
				'RC = runthis(SAPSetupDir & "\setup\NwSapSetup.exe", "/product:NWBC40 /uninstall /noDlg", TRUE)
				'log("Uninstalled NWBC40, RC: " & RC)
			End If
		Else
			log("Can't install SAPGUI!")
			exitcode(bad)
		End If
	'Else
	'	log("Can't install SAPGUI!")
	'End If
End Function
log("waserror :" & waserror)

progress 2,"Stopping updater"
stopupdater()

Function stopupdater
	'stop updater service
	log("Looking for updater service")
	strServiceName = "NWSAPAutoWorkstationUpdateSvc"
	'Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name ='" & strServiceName & "'")
	For Each objService in colListOfServices
		log("Updater service found and attempting to stop...")
		progress 0,"Stopping updater service"
		objService.StopService()	
		iter = 0
		Do
			Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name ='" & strServiceName & "'")
			done = TRUE
			For Each svc In colListOfServices
				If svc.State <> "Stopped" Then
					done = False
					log("Updater Service not stopped yet. Waiting 1s. Attempt #" & iter)
					WScript.Sleep 1000
					'Exit For
				Else
					log("Updater Service was found stopped")
					done = TRUE
				End If
				Exit For
			Next
			iter = iter + 1
			If iter > MAX_ITER Then
				log("Timeout trying to stop updater service")
				timeout = TRUE
				done = TRUE
				Exit Do
			End If
		Loop Until done
	Next
End Function

progress 2,"Fix folder permissions"
'fixsapfolderperms()
fixsapfolderperms2()

'add user:full to SAP folder
Function fixsapfolderperms
	'On Error Resume Next
	log("Fix SAP Folder perms #1")
	If exists(sapguidir,0) Then
		rootsapfolder = FSO.GetParentFolderName(sapguidir)
		If NOT IsBlank(rootsapfolder,FALSE) Then
			RC = runthis(windowsdir & "\system32\icacls.exe", stringthis(rootsapfolder) & " /grant Users:F /inheritance:e /T /C", FALSE)
		Else
			log("Not able to find sap root folder")
		End If
		'On Error Goto 0
	Else
		log("sapguidir doesnt exist")
	End If
End Function

Function fixsapfolderperms2
	log("Fix SAP Folder perms #2")
	On Error Resume Next
	CreateObject("Shell.Application").ShellExecute "cmd.exe", "/C " & stringthis(sourcedir & "\fix.folder.perms.bat"), "", "runas", 0
	'example runthis: RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\extension_fax_new.reg"), FALSE)
	'RC = runthis(start, "/C " & stringthis(sourcedir & "\fix.folder.perms.bat runas"), FALSE)
	log("RC from fix.folder.perms.bat: " & Err.Number)
	On Error Goto 0
End Function

progress 2,"Deleting bad shortcuts"
delshortcuts()

Function delshortcuts
	'Always delete left over shortcuts
	log("Delete left over shortcuts")
	delfile(userprofile & "\Desktop\hROH Users Prod Systems SSO.lnk")
	delfile(publicfolder & "\Desktop\hROH Users Prod Systems SSO.lnk")
	delfile(publicfolder & "\Public Desktop\hROH Users Prod Systems SSO.lnk")
	delfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End\hROH Users All Systems SSO.lnk")
	delfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End\SAP Logon.lnk")
	delfile(userprofile & "\Desktop\SAPlogon-Pad.lnk")
	delfile(publicfolder & "\Desktop\SAPlogon-Pad.lnk")
	delfile(publicfolder & "\Public Desktop\SAPlogon-Pad.lnk")
	delfile(userprofile & "\Desktop\SAP Logon.lnk")
	delfile(publicfolder & "\Desktop\SAP Logon.lnk")
	delfile(publicfolder & "\Public Desktop\SAP Logon.lnk")
	delfile("C:\Documents and Settings\All Users\Desktop\SAP Logon.lnk")
	delfile(publicdesktop & "\Dow SAP Systems.lnk")
	delfile(shortcutfiles & "\Dow SAP Systems.lnk")
End Function

progress 2,"Deleting UAC Virtualstore files"
delvirtualstore()

Function delvirtualstore
	'Delete UAC virtualstore files
	'If this the person running this script does not have full permission to write to windows, then UAC will cache it in virtual store.
	log("Delete UAC virtualstore files")
	delfile(localappdata & "\VirtualStore\Windows\saplogon.ini")
	delfile(localappdata & "\VirtualStore\Windows\SapLogonTree.xml")
	delfile(localappdata & "\VirtualStore\Windows\SAPMSG.ini")
	delfile(localappdata & "\VirtualStore\Windows\saproute.ini")
	delfile(localappdata & "\VirtualStore\Windows\sapshortcut.ini")
	delfile(localappdata & "\VirtualStore\Windows\saplogon.NP.SSO.ini")
	delfile(localappdata & "\VirtualStore\Windows\saplogon.NP.NONSSO.ini")
	delfile(localappdata & "\VirtualStore\Windows\saplogon.P.SSO.ini")
	delfile(localappdata & "\VirtualStore\Windows\saplogon.P.NONSSO.ini")
	delfile(localappdata & "\VirtualStore\Windows\SAPUILandscape.xml")
	delfile(localappdata & "\VirtualStore\Windows\SAPUILandscapeGlobal.xml")
	delfolder(localappdata & "\VirtualStore\Program Files\SAP")
End Function

progress 2,"Deleting ini files"
'dont run this because if the copy fails, then the user wont have any files...
delinifiles()

Function delinifiles
	log("Delete ini files")
	delfile(windowsdir & "\saplogon.ini")
	delfile(windowsdir & "\SAPMSG.ini")
	delfile(windowsdir & "\saproute.ini")
	delfile(inifiles & "\saplogon.NP.SSO.ini")
	delfile(inifiles & "\saplogon.NP.NONSSO.ini")
	delfile(inifiles & "\saplogon.P.SSO.ini")
	delfile(inifiles & "\saplogon.P.NONSSO.ini")
	delfile(inifiles & "\SAPUILandscape.xml")
	delfile(inifiles & "\SAPUILandscapeGlobal.xml")
End Function

'del taskbar shortcuts?
'C:\Users\n003352\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar

progress 2,"Deleting old/junk files"
deljunk()

'MSXML to fix issue?
'MSXML 4.0 Service Pack 3
'https://www.microsoft.com/en-us/download/details.aspx?id=15697

Function deljunk()
	log("Deleting bad NWBC Options xml")
	delfile("c:\programdata\sap\nwbc\nwbcOptions.xml")
	delfile(PathToAppData & "\sap\NWBC\NwbcOptions.xml")
	delfile(nwbcfiles & "\NwbcOptions.xml")

	'del rules file
	log("Del rules file")
	progress 0,"Deleting rules file"
	delfile (inifiles & "\saprules.xml")

	'del Tree file
	'log("Del Tree file")
	'progress 1,"Deleting Tree file"
	'delfile (inifiles & "\SapLogonTree.xml") 

	'del bad saplogon.ini file
	log("Del saplogon.ini file in appdata")
	progress 0,"Deleting saplogon.ini file in appdata"
	delfile (inifiles & "\saplogon.ini")
	
	'del run.as.admin.reg
	log("Del run.as.admin.reg")
	progress 0,"Deleting run.as.admin.reg"
	delfile (userprofile & "\Desktop\run.as.admin.reg")
	
	'del install.1.bat from desktop. 
	log("Del install.1.bat")
	progress 0,"Deleting install.1.bat"
	delfile (userprofile & "\Desktop\install.1.bat")
	
	'del install.2.vbs from desktop. 
	log("Del install.2.vbs")
	progress 0,"Deleting install.2.vbs"
	delfile (userprofile & "\Desktop\install.2.vbs")
	
	'del bad librfc32.dll file from wwi folder
	log("Del bad librfc32.dll from WWI folder")
	progress 0,"Deleting bad librfc32.dll from WWI folder"
	delfile (sapguidir & "\SAPgui\wwi\" & "librfc32.dll")
	
	'del bad graphics files in wwi folder
	log("Del bad graphics from WWI folder")
	progress 0,"Deleting bad graphics from WWI folder"
	delfile (sapguidir & "\SAPgui\wwi\" & "ANGUS.bmp")
	delfile (sapguidir & "\SAPgui\wwi\" & "Univation.bmp")
End Function

'make dirs if they do not exist
progress 2,"Create missing folders"
log("Make dirs if they do not exist")
'createfolder(inifilesparent)
'createfolder(PathToAppData)
'createfolder(PathToAppData & "\SAP")
'createfolder(PathToAppData & "\SAP\Common")
'createfolder(PathToAppData & "\SAP\NWBC")
createfolder(shortcutfiles)
createfolder(inifiles) '\SAP\Common
createfolder(nwbcfiles) '\SAP\NWBC
createfolder("C:\temp") 'for C:\PRINTLBL.BAT
createfolder("C:\guixt\scripts") 'for guixt
createfolder("C:\GuiXT\Cache") 'for guixt
'createfolder(inifilesparent & "\SAP\Common") 'same as inifiles
'createfolder(inifilesparent & "\SAP\NWBC")
'createfolder(commfiles & "\microsoft shared\vba\vba6\")
'createfolder(commfiles & "\microsoft shared\vba\vba7\")

log("Copying PRINTLBL.BAT to C:\")
progress 1,"Copying PRINTLBL.BAT to C:\"
copyfile sourcedir, "PRINTLBL.BAT", "C:\", ""

'Copying saproute.ini & sapmsg.ini to windows folder
log("Copying saproute.ini & sapmsg.ini to windows folder")
progress 1,"Copying saproute.ini & sapmsg.ini to windows folder"
copyfile sourcedir, "SAPMSG.INI", windowsdir, ""
If InstalledVersion = 740 Then
	log("Version is 740, copying saproute.ini without /H/")
	copyfile sourcedir, "saproute.ini", windowsdir, "saproute.ini"
Else
	log("Version is not 740, copying saproute.ini with /H/")
	copyfile sourcedir, "saproute.730.ini", windowsdir, "saproute.ini"
End If

progress 2,"Copying saplogon.ini to windows folder"
If UserDomain = "DOW" or UserDomain = "BSNCONNECT" or UserDomain = "JVSERVICES" or UserDomain = "JVSERVICES-CORE" Then
	log("Copying SSO saplogon.ini to windows folder")
	copyfile sourcedir, "saplogon.NP.SSO.ini", windowsdir, "saplogon.ini"
Else
	log("Copying NON SSO saplogon.ini to windows folder")
	copyfile sourcedir, "saplogon.NP.NONSSO.ini", windowsdir, "saplogon.ini"
End If

'copy newest ini files
If NOT InstalledVersion = 740 Then
	log("Version is not 740, copying all 4 saplogon ini files")
	log("Copying saplogon INI files")
	progress 0,"Copying newer INI files"
	copyfile sourcedir, "saplogon.NP.SSO.ini", inifiles, ""
	copyfile sourcedir, "saplogon.P.SSO.ini", inifiles, ""
	copyfile sourcedir, "saplogon.NP.NONSSO.ini", inifiles, ""
	copyfile sourcedir, "saplogon.P.NONSSO.ini", inifiles, ""
	'copyfile sourcedir, "NwbcOptions.xml", nwbcfiles, ""
End If

'force a pc to use XML's or INI's
'[HKEY_LOCAL_MACHINE\SOFTWARE(\Wow6432Node)\SAP\SAPLogon="LandscapeFormatEnabled"=REG_DWORD:00000001

'copy newest saplogon xml
log("Copying saplogon XML files")
progress 2,"Copying newer XML files"
copyfile sourcedir, "SAPUILandscape.xml", inifiles, ""
copyfile sourcedir, "SAPUILandscapeGlobal.xml", inifiles, ""

'create shortcuts
log("Create shortcuts")
'WIN7 & WIN2008: C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End
'XP & WIN2003:   C:\Documents and Settings\All Users\Start Menu\Programs\SAP Front End
progress 2,"Creating Shortcuts"
'NWBCdir = "C:\Program Files\SAP\NWBC40"
'sapguidir = "C:\Program Files\SAP\FrontEnd"
'inifiles = PathToAppData & "\SAP\Common"
'nwbcfiles = PathToAppData & "\SAP\NWBC"
'shortcutfiles = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End"
'desktop = publicfolder & "\Desktop"
If exists(sapguidir & "\SAPgui\saplogon.exe",0) Then
	logonprog = "saplogon.exe"
ElseIf exists(sapguidir & "\SAPgui\saplgpad.exe",0) Then
	logonprog = "saplgpad.exe"
Else
	'neither exists. wow...
	logonprog = "saplogon.exe"
End If

If InstalledVersion < 740 Then
	log("Version is under 740, creating old 4 saplogon shortcuts")
	'only create shortcuts for 4 ini files if under 740
	'createshortcut publicdesktop, "Netweaver Business Client", NWBCdir, "NWBC.exe", "", "Netweaver Business Client"
	If UserDomain = "DIR" Then
		'ACN, no SSO
		createshortcut publicdesktop, "Prod Systems NON SSO", sapguidir & "\SAPgui", logonprog, "/ini_file=""" & inifiles & "\saplogon.P.NONSSO.ini""", "Production SAP Systems"
	Else
		createshortcut publicdesktop, "Prod Systems SSO", sapguidir & "\SAPgui", logonprog, "/ini_file=""" & inifiles & "\saplogon.P.SSO.ini""", "Production SAP Systems Single Sign On"
	End If
	createshortcut shortcutfiles, "Prod Systems SSO", sapguidir & "\SAPgui", logonprog, "/ini_file=""" & inifiles & "\saplogon.P.SSO.ini""", "Production SAP Systems Single Sign On"
	createshortcut shortcutfiles, "Prod Systems NON SSO", sapguidir & "\SAPgui", logonprog, "/ini_file=""" & inifiles & "\saplogon.P.NONSSO.ini""", "Production SAP Systems"
	createshortcut shortcutfiles, "All Systems SSO", sapguidir & "\SAPgui", logonprog, "/ini_file=""" & inifiles & "\saplogon.NP.SSO.ini""", "All SAP Systems Single Sign On"
	createshortcut shortcutfiles, "All Systems NON SSO", sapguidir & "\SAPgui", logonprog, "/ini_file=""" & inifiles & "\saplogon.NP.NONSSO.ini""", "All SAP Systems"
	For Each hostname IN gsdservers
		If pccomputername = hostname Then
			'gsd server
			log("Creating additional shortcuts for GSD server")
			createfolder("D:\Program Files\GSD_Links")
			createshortcut "D:\Program Files\GSD_Links", "All Systems SSO", sapguidir & "\SAPgui", logonprog, "/ini_file=""" & inifiles & "\saplogon.NP.SSO.ini""", "All SAP Systems Single Sign On"
			createshortcut "D:\Program Files\GSD_Links", "All Systems NON SSO", sapguidir & "\SAPgui", logonprog, "/ini_file=""" & inifiles & "\saplogon.NP.NONSSO.ini""", "All SAP Systems"
		End If
	Next
End If

If InstalledVersion >= 740 Then
	'>=740
	log("Version 740 or higher, creating 740 shortcuts")
	createshortcut publicdesktop, "Dow SAP Systems", sapguidir & "\SAPgui", logonprog, "", "Dow SAP Systems"
	createshortcut shortcutfiles, "Dow SAP Systems", sapguidir & "\SAPgui", logonprog, "", "Dow SAP Systems"
	log("Version 740 or higher, deleting old 720/730 shortcuts")
	delshortcuts2()
End If

Function delshortcuts2
	'old 720 & 730 shortcuts are no longer needed
	delfile (shortcutfiles & "\Prod Systems SSO.lnk")
	delfile (shortcutfiles & "\Prod Systems NON SSO.lnk")
	delfile (shortcutfiles & "\All Systems SSO.lnk")
	delfile (shortcutfiles & "\All Systems NON SSO.lnk")
	delfile (publicdesktop & "\Prod Systems SSO.lnk")
	delfile (publicdesktop & "\Prod Systems NON SSO.lnk")
End Function
	

'function to create shortcuts
Function createshortcut(sclocation, scfilename, scfolder, scexe, scargs, scdescription)
	If exists(scfolder & "\" & scexe, FALSE) Then
		'set WshShell = WScript.CreateObject("WScript.Shell")
		'strDesktop = oShell.SpecialFolders("Desktop")
		log("Creating shortcut   : " & sclocation & "\" & scfilename & ".lnk")
		log("Creating shortcut to: " & scfolder & "\" & scexe)
		log("Remove read-only if set")
		exists sclocation & "\" & scfilename & ".lnk", TRUE
		'If FSO.FileExists(sclocation & "\" & scfilename & ".lnk") Then
		'	FSO.DeleteFile(sclocation & "\" & scfilename & ".lnk")
		'End If
		log("Shortcut target  : " & scfolder & "\" & scexe & " " & scargs)
		On Error Resume Next
		set oShellLink = oShell.CreateShortcut(sclocation & "\" & scfilename & ".lnk")
		oShellLink.TargetPath = scfolder & "\" & scexe
		'oShellLink.Arguments = "/ini_file=""" & scdestination & "\" & sctarget & """"
		oShellLink.Arguments = scargs
		'oShellLink.WindowStyle = 1
		'oShellLink.Hotkey = "CTRL+SHIFT+F"
		'oShellLink.IconLocation = "notepad.exe, 0"
		oShellLink.Description = scdescription
		oShellLink.WorkingDirectory = scfolder
		oShellLink.Save
		If err <> 0 Then 
			log("Unable to create above shortcut")
			DisplayErrorInfo
		Else
			log("Shortcut created")
		End If
		On Error Goto 0
	Else
		log("Shortcut Not Created")
	End If
 End Function

 
'copy updated wwi files
'we don't modify wwi.dot or wwilayt.dot
'customizable:  wwi.ini, wwilabel.ini, wwi_user.lbl, and wwi_all.lbl 
log("Copy updated wwi files")
progress 2,"Copying new wwi files"
If OSBIT = 64 Then
	copyfile sourcedir, "wwi.64.ini", sapguidir & "\SAPgui\wwi", "wwi.ini"
Else
	copyfile sourcedir, "wwi.ini", sapguidir & "\SAPgui\wwi", ""
End If
copyfile sourcedir, "wwidispl.dot", sapguidir & "\SAPgui\wwi", ""
copyfile sourcedir, "userexit.dot", sapguidir & "\SAPgui\wwi", ""
copyfile sourcedir, "wwi_user.lbl", sapguidir & "\SAPgui\wwi", ""
'once wwi.dot was missing, lets check and fix, well for now, lets just check they exist.
If NOT exists (sapguidir & "\SAPgui\wwi\wwi.dot",0) Then
End If
If NOT exists (sapguidir & "\SAPgui\wwi\wwilayt.dot",0) Then
End If


'copy all EHS files
log("Copy all EHS files")
progress 2,"Copying new EHS files"
If FSO.FolderExists(sapguidir & "\SAPgui\wwi\graphics") Then
	log("Copying: " & sourcedir & "\*.graphics to " & sapguidir & "\SAPgui\wwi\graphics\")
	log("If this fails, just continue")
	'need to change to arrays... ugh
	'On Error Resume Next
	'FSO.Copyfile sourcedir & "\*.bmp", sapguidir & "\SAPgui\wwi\graphics\"
	'FSO.Copyfile sourcedir & "\*.tif", sapguidir & "\SAPgui\wwi\graphics\"
	'FSO.Copyfile sourcedir & "\*.gif", sapguidir & "\SAPgui\wwi\graphics\"
	'FSO.Copyfile sourcedir & "\*.jpg", sapguidir & "\SAPgui\wwi\graphics\"
	'FSO.Copyfile sourcedir & "\*.png", sapguidir & "\SAPgui\wwi\graphics\"
	'On Error GoTo 0
	'
	'or
	'
	'for each oFile in oFSO.GetFolder("C:\Program Files\Keys1").Files
    '  oFSO.GetFile(oFile).Copy "C:\Program Files\keys2\", true
	'next
	'
	'or faster:
	'use robocopy, if it doesnt exist then use xcopy
	RC = robocopythis(sourcedir, sapguidir & "\SAPgui\wwi\graphics", "*.bmp *.jpg *.gif *.png *.tif *.tiff", "")
	'RC = robocopythis(sourcedir, sapguidir & "\SAPgui\wwi\graphics", "*.jpg")
	'RC = robocopythis(sourcedir, sapguidir & "\SAPgui\wwi\graphics", "*.gif")
	'RC = robocopythis(sourcedir, sapguidir & "\SAPgui\wwi\graphics", "*.png")
	'RC = robocopythis(sourcedir, sapguidir & "\SAPgui\wwi\graphics", "*.tif")
	'RC = runthis(windowsdir & "\system32\xcopy.exe", """" & sourcedir & "\*.bmp""" & " " & """" & sapguidir & "\SAPgui\wwi\graphics\""" & " /Y /R /H", FALSE)
	'RC = runthis(windowsdir & "\system32\xcopy.exe", """" & sourcedir & "\*.jpg""" & " " & """" & sapguidir & "\SAPgui\wwi\graphics\""" & " /Y /R /H", FALSE)
	'RC = runthis(windowsdir & "\system32\xcopy.exe", """" & sourcedir & "\*.gif""" & " " & """" & sapguidir & "\SAPgui\wwi\graphics\""" & " /Y /R /H", FALSE)
	'RC = runthis(windowsdir & "\system32\xcopy.exe", """" & sourcedir & "\*.png""" & " " & """" & sapguidir & "\SAPgui\wwi\graphics\""" & " /Y /R /H", FALSE)
	'RC = runthis(windowsdir & "\system32\xcopy.exe", """" & sourcedir & "\*.tif""" & " " & """" & sapguidir & "\SAPgui\wwi\graphics\""" & " /Y /R /H", FALSE)
End If

progress 1,"Copy EH&S WWI SP41 HF files"
'not fixed in 740003088954
If SourceDateModified = 740003088954 Then
	log("Copy EH&S WWI SP41 HF files")
	'always copy
	copyfile sourcedir, "WwiRes32.dll", sapguidir & "\SAPgui\wwi", ""
	copyfile sourcedir, "WwiRes64.dll", sapguidir & "\SAPgui\wwi", ""
	'i'm not convinced we use the 64bit ver. we'll find out!
	copyfile sourcedir, "WwiLabel.dll", sapguidir & "\SAPgui\wwi", ""
	'If OSBIT = 64 Then
	'	copyfile sourcedir, "WwiLabel.64.dll", sapguidir & "\SAPgui\wwi", "WwiLabel.dll"
	'Else
	'	copyfile sourcedir, "WwiLabel.dll", sapguidir & "\SAPgui\wwi", ""
	'End If
Else
	log("NOT Copying EH&S WWI SP41 HF files")
End If

'copy SSO file gsskrb5.dll
'need to copy .old if 2003
log("Copy SSO file gsskrb5.dll")
progress 1,"Copying SSO file"
copyfile sourcedir, "gsskrb5.dll", sapguidir & "\SAPgui", ""
copyfile sourcedir, "gx64krb5.dll", sapguidir & "\SAPgui", ""

'for excel macros for EHS in ESPS
'Only works below 2013
'add VBA converter files registry settings
'log("Adding VBA converter files registry settings")
'progress 1,"Adding VBA converter files registry settings"
'Writekey "HKEY_CLASSES_ROOT\TypeLib\{000204F3-0000-0000-C000-000000000046}\1.0", "Visual Basic For Applications", "REG_SZ"

'copy VBA Converter files
log("Copy VBA Converter files")
progress 1,"Copying VBA Converter files"
'MS HotFix: 926430
'only works for 32bit versions of XLS
'dll's excel 2007
copyfile sourcedir, "vbacv10.dll", "C:\Program Files\microsoft shared\vba\vba6\", ""
copyfile sourcedir, "vbacv10d.dll", "C:\Program Files\microsoft shared\vba\vba6\", ""
'dll's excel 2010
copyfile sourcedir, "vbacv10.dll", "C:\Program Files\microsoft shared\vba\vba7\", ""
copyfile sourcedir, "vbacv10d.dll", "C:\Program Files\microsoft shared\vba\vba7\", ""
'dll's excel 2013+
copyfile sourcedir, "vbacv10.dll", "C:\Program Files\microsoft shared\vba\vba7.1\", ""
copyfile sourcedir, "vbacv10d.dll", "C:\Program Files\microsoft shared\vba\vba7.1\", ""
'\Program Files\Microsoft Office 15\root\vfs\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA7.1 (says someone...)
If OSBIT = 64 Then
	'64bit OS needs them here instead:
	'dll's excel 2007
	copyfile sourcedir, "vbacv10.dll", "C:\Program Files (x86)\microsoft shared\vba\vba6\", ""
	copyfile sourcedir, "vbacv10d.dll", "C:\Program Files (x86)\microsoft shared\vba\vba6\", ""
	'dll's excel 2010
	copyfile sourcedir, "vbacv10.dll", "C:\Program Files (x86)\microsoft shared\vba\vba7\", ""
	copyfile sourcedir, "vbacv10d.dll", "C:\Program Files (x86)\microsoft shared\vba\vba7\", ""
	'dll's excel 2013+
	copyfile sourcedir, "vbacv10.dll", "C:\Program Files (x86)\microsoft shared\vba\vba7.1\", ""
	copyfile sourcedir, "vbacv10d.dll", "C:\Program Files (x86)\microsoft shared\vba\vba7.1\", ""
End If
'olb's to excel folder
excelpath = ReadKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\Path")
If NOT IsBlank(excelpath,TRUE) Then
	copyfile sourcedir, "xl5en32.olb", excelpath, ""
	copyfile sourcedir, "gren50.olb", excelpath, ""
	'If exists("C:\Program Files\Microsoft Office 15",0) Then
		'copyfile sourcedir, "xl5en32.olb", "C:\Program Files\Microsoft Office 15\root\office15", ""
		'copyfile sourcedir, "gren50.olb", "C:\Program Files\Microsoft Office 15\root\office15", ""
	'End If
End If
'olb's to windows
If OSBIT = 64 Then
	copyfile sourcedir, "vbaen32.olb", windowsdir & "\SysWOW64\", ""
	copyfile sourcedir, "vbaend32.olb", windowsdir & "\SysWOW64\", ""
Else
	copyfile sourcedir, "vbaen32.olb", windowsdir & "\System32\", ""
	copyfile sourcedir, "vbaend32.olb", windowsdir & "\System32\", ""
End If
CreateObject("Shell.Application").ShellExecute "cmd.exe", "/C " & stringthis(sourcedir & "\vbaconv.bat"), "", "runas", 0
log("RC from vbaconv.bat: " & Err.Number)

'Fix GuiXT
'Input Designer needs a license key and installing 740 seems to enable it, so disable it
log("Fixing GuiXT")
Writekey "HKEY_CURRENT_USER\Software\SAP\SAPGUI Front\SAP Frontend Server\Customize\GuiXT.ComponentInputAssistant", 0, "REG_DWORD"

'Fix "Invalid GUI input data: ST_USER_MAX_WSIZE wrong data"
'Oss note 2229515 - Invalid GUI data in SAPGUI
'Only needed for SAPGUI 7.40.0 -> 7.40.5
'log("SourceDateModified :" & SourceDateModified)
log("Invalid GUI input data: ST_USER_MAX_WSIZE wrong data")
If SourceDateModified < 740002060000 Then
	'need both keys
	Writekey "HKEY_CURRENT_USER\Software\SAP\SAPLogon\Options\SapguiNTCmdOpts", "/SUPPORTBIT_OFF=MAX_WSIZE /SUPPORTBIT_OFF=NO_FOCUS_ON_LIST", "REG_SZ"
Else
	'only need 1 key
	'Delkey "HKEY_CURRENT_USER\Software\SAP\SAPLogon\Options\SapguiNTCmdOpts", ""
	Writekey "HKEY_CURRENT_USER\Software\SAP\SAPLogon\Options\SapguiNTCmdOpts", "/SUPPORTBIT_OFF=NO_FOCUS_ON_LIST", "REG_SZ"
End If

'Fix "Enhanced search broken"
'Oss note ???
'Only needed for SAPGUI < 7.40.8
progress 1,"Enhanced search broken, disabling if below 7.40.8"
log("Enhanced search broken, disabling if below 7.40.8")
If SourceDateModified < 740003080000 Then
	'disable
	Writekey "HKEY_CURRENT_USER\Software\SAP\SAPGUI Front\SAP Frontend Server\Customize\EnhancedSearchMode", "0", "REG_DWORD"
Else
	'enable
	Writekey "HKEY_CURRENT_USER\Software\SAP\SAPGUI Front\SAP Frontend Server\Customize\EnhancedSearchMode", "1", "REG_DWORD"
End If

'old keys
progress 1,"Deleting old 7.10 & 7.20 add/remove keys"
log("Deleting old 7.10 & 7.20 add/remove keys, if version above 7.40.0")
If SourceDateModified > 740000000000 Then
	'7.10
	Delkey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SAPGUI710", ""
	Delkey "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\SAPGUI710", ""
	'7.20
	DelKey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\SAP SAPGUI EN",""
	DelKey "HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\SAP SAPGUI EN",""
End If

If fixhistory Then
	DoFixHistory
End If

'Fix Broken History - Isn't used, but could be.
Function DoFixHistory
	log("Fixing history")
	delfolder(PathToAppData & "\SAP\SAP GUI\History")
End Function

'set SNC_LIB 
log("Setting SNC_LIB Environment var")
progress 2,"Setting SNC_LIB Environment Var"
'KeyExists("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB")
If OSBIT = 32 then
	'oShell.run "cmd /c reg add " & """HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment""" & " /v SNC_LIB /t REG_EXPAND_SZ /d " & """C:\Program Files\SAP\FrontEnd\SAPgui\gsskrb5.dll""" & " /F"
	'oShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB", sapguidir & "\SAPgui\gsskrb5.dll", "REG_EXPAND_SZ"
	Writekey "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB", sapguidir & "\SAPgui\gsskrb5.dll", "REG_EXPAND_SZ"
	Writekey "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB_2", sapguidir & "\SecureLogin\lib\sapcrypto.dll", "REG_EXPAND_SZ"
'	KeyExists("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB")
else
	'oShell.run "cmd /c reg add HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment /v SNC_LIB /t REG_EXPAND_SZ /d C:\Program Files (x86)\SAP\FrontEnd\SAPgui\gsskrb5.dll /F"
	'oShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB", sapguidir & "\SAPgui\gsskrb5.dll", "REG_EXPAND_SZ"
	'BLG Update per Bill Cheng to support 64bit version of other sap applicaitons 11/7/2016
	WriteKey "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB", sapguidir & "\SAPgui\gsskrb5.dll", "REG_EXPAND_SZ"
	Writekey "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB_64", sapguidir & "\SAPgui\gx64krb5.dll", "REG_EXPAND_SZ"
	Writekey "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB_2", sapguidir & "\SecureLogin\lib\sapcrypto.dll", "REG_EXPAND_SZ"
	Writekey "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB_64_2", "C:\Program Files\SAP\FrontEnd\SecureLogin\lib\sapcrypto.dll", "REG_EXPAND_SZ"
'	KeyExists("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB")
end if

'log
KeyExists("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB")
KeyExists("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SNC_LIB_2")

'copy fonts OLD
'log("Copy fonts")
'progress 2,"Copying Fonts"
'If NOT exists(windowsdir & "\Fonts\3OF9.TTF",0) Then
	'objFolderFonts.CopyHere(sourcedir & "\3OF9.TTF")
'End If
'If NOT exists(windowsdir & "\Fonts\c39qrtr.ttf",0) Then
	'objFolderFonts.CopyHere(sourcedir & "\c39qrtr.ttf")
'End If
'If NOT exists(windowsdir & "\Fonts\c128rhp3.ttf",0) Then
	'objFolderFonts.CopyHere(sourcedir & "\c128rhp3.ttf")
'End If
'If NOT exists(windowsdir & "\Fonts\TCCM____.TTF",0) Then
	'objFolderFonts.CopyHere(sourcedir & "\TCCM____.TTF")
'End If
'If NOT exists(windowsdir & "\Fonts\Code-128 Normal.ttf",0) Then
'	objFolderFonts.CopyHere(sourcedir & "\Code-128 Normal.ttf")
'End If


'copy fonts NEW
'for MSDS's
progress 1,"Copying Fonts"
on error resume next
log("Copy fonts from " & sourcedir)
Set folder = FSO.getfolder(sourcedir)
set ABC = folder.files
For Each file in ABC
	fullfilename = file.path
	filename = file.name
	If instr(ucase(filename),"TTF") > 0 AND NOT instr(ucase(filename),"TINY") > 0 Then
		'found a TTF
		If NOT Exists(windowsdir & "\Fonts\" & filename,0) Then
			'not in windows, so copy
			log("Copying " &  fullfilename & " to " & windowsdir & "\Fonts\")
			On error resume next
			'objFolderFonts.CopyHere fullfilename, FOF_NOCONFIRMATION
			objFolderFonts.CopyHere fullfilename
			On error goto 0
			log("Coppied " &  fullfilename & " to " & windowsdir & "\Fonts\")
		End If
	Else
		'not found
		'msgbox "NOT TTF : " & filename
	End If
Next
On error goto 0

' update services file
log("Removing bad rows from services file")
progress 1,"Removing bad rows from services file"
'exists windowsdir & "\system32\drivers\etc\services",1
If exists(servicesfile,1) Then
	log("Found services file. Removing bad rows")
	strNewContents = ""
	Set userservicesfile = FSO.OpenTextFile(servicesfile, ForReading)
	Do Until userservicesfile.AtEndOfStream
		'read in servicesfile, one line at a time
		strLine = userservicesfile.ReadLine
		count = 0
		For Each badline in badrows
			If InStr(strLine, badline) <> 0 Then 
				'found a bad row
				log("Found bad row in services, removing: " & badline)
				count = 1
			End If
		Next
		If count = 0 Then 'there wasn't a match
			'msgbox "NO MATCH: " & badline & " in " & strLine
			'strNewContents becomes their services file, excluding bad rows.
			strNewContents = strNewContents & strLine & vbCrLf
		End If
	Loop
	userservicesfile.Close
Else
	log("Unable to find services file. Not removing bad rows")
End If

'open services for updating
'this will fail if not admin. code around it...
log("Appending new rows to services file")
progress 1,"Appending new rows to services file"
On Error Resume Next
Set userservicesfile = FSO.GetFile(servicesfile)
log("Size of users service file: " & CINT(userservicesfile.Size / 1024) & "KB")
Set userservicesfile = Nothing
Set userservicesfile = FSO.OpenTextFile(servicesfile, ForWriting)
If Err <> 0 Then
	'bad, can't open services file
	DisplayErrorInfo
	On Error Goto 0
	log("Not able to open users services file")
	If NOT Silent Then
		'log("Waiting for user: services file can't be updated")
		'msgbox("Not able to update services file" & vbCrLf & "This can happen if you did not 'Run as Admin'" & vbCrLf & "I recommended you try again & 'Run as Admin'")	
	End If
Else
	On Error Goto 0
	log("Able to open Services file for updating")
	If NOT IsBlank(strNewContents,TRUE) Then
		'if strNewContents isnt empty, then it was built from their services file. 
		'So write that back out - creating a new services file excluding bad rows
		userservicesfile.Write strNewContents
	End If
	userservicesfile.Close
	'read new services.txt from install dir
	log("Reading in new services.txt")
	exists sourcedir & "\services.txt",0
	Set newservicesfile = FSO.OpenTextFile (sourcedir & "\services.txt", ForReading)
	newservicesfiletext = newservicesfile.ReadAll
	newservicesfile.Close
	'split every line in new services file in to an array
	aLines = Split(newservicesfiletext, vbCrLf) 
	'remove bad lines from services file & add new ones.
	ApendLinesToFile aLines, servicesfile 
End If


' Append lines given as an array to the specified text file
Function ApendLinesToFile(aLinesToAppend, servicesfile)
	log("Opening userservicesfile to fill userservicesfiletext")
	Set userservicesfile = FSO.OpenTextFile(servicesfile)
	'Set demofile = FSO.GetFile(servicesfile)
	'a = demofile.Path
	On Error Resume Next
	'next fails 'Input past end of line' if file is empty
	userservicesfiletext = userservicesfile.ReadAll
	If Err.Number = 0 Then 
		'read file OK
		log("Successfully read users services file into userservicesfiletext")
	ElseIf Err.Number = 62 Then 
		'file is most likely empty
		log("Error reading users services file, it's likely empty")
		userservicesfiletext = ""
	Else 
		'another error, best to at least report this 
		log("Error reading users services file")
		serservicesfiletext = ""
	End If 
	On Error Goto 0
	aLinesInFile=Split(userservicesfiletext,VBCr)
	userservicesfile.Close
	Set userservicesfile = Nothing
	log("Backing up user services file...")
	'copyfile a, "", a & ".bak", ""
	On Error Resume Next
	'FSO.CopyFile userservicesfile hmmm
	On Error Goto 0
	log("Opening userservicesfile to append new rows")
	Set userservicesfile = FSO.OpenTextFile(servicesfile, ForAppending, True)
	For Each strLine In aLinesToAppend
		if VerifyIfExistLine(strLine,aLinesInFile) = False then
			userservicesfile.WriteLine strLine	
			log("Appended New Line: " & strLine)
		end if
	Next
	userservicesfile.Close
	Set userservicesfile = Nothing
	log("Finished appending new rows")
	'Set demofile = Nothing
End Function
log("Updated services file")

' Verify if line exists in the specified text file given as an array of lines
Function VerifyIfExistLine(strLineToAppend, aLinesInFile)
	nr = 0
	'msgbox "in VerifyIfExistLine"
	For each strLineInFile in aLinesInFile
		If InStr(strLineInFile,strLineToAppend) then
			nr = 1
		end if
	next
	if nr = 0 then
		VerifyIfExistLine = FALSE
		'msgbox strLineInFile & " Doesn't exist"
	else
		VerifyIfExistLine = TRUE
		'msgbox strLineInFile & " exists"
	end if
End Function


' fix .FAX association
'This is for ROH
log("Fix .FAX association")
progress 2,"Fix .FAX association"
RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\extension_fax_new.reg"), FALSE)
'RC = runthis(start, stringthis(sourcedir & "\extension_fax_new.reg"), FALSE)
'Win7:
'WORKS: Windows Photo Viewer: PhotoViewer.dll (shimgvw.dll?)
'Windows Live Photo Gallery: WLXPhotoGallery.exe
'Microsoft Office Picture Manager (removed in Office 2013)
'Microsoft Office Document Imaging: MSPVIEW.EXE
'Win8:


' install SSO for hROH SAP Systems
log("Install SSO for hROH SAP Systems")
progress 2,"Installing SSO for hROH SAP Systems"
'oShell.run "msiexec /qn /i " & sourcedir & "\CSTBscw-4.2.4-34901.Windows.x86.msi"
'32bit version. Fails on 64bit
RC = runthis(msiexecfile, "/qn /i " & stringthis(sourcedir & "\CSTBscw-4.2.4-34901.Windows.x86.msi"), FALSE)


'install the BSNConnect root certificate 
'prevents a SSL error
log("Install BSNConnect root certificate")
progress 2,"Installing BSNConnect root certificate"
RC = runthis(windowsdir & "\system32\certutil.exe", "-addstore -f -enterprise -user root " & stringthis(sourcedir & "\BSNC_Root_Certificate.cer"), FALSE) '> NUL


' fix saplogon_ini_file env var
log("Fix many saplogonreg settings")
progress 2,"Setting SAPGUI registry settings"
'for under 7.40 and legacy apps that need it
log("Reading default xml/ini location")
ReadKey("HKEY_CURRENT_USER\Software\SAP\SAPLogon\Options\PathConfigFilesLocal")
If server Then
	'not needed for users as this file goes to %appdata%\SAP\common
	'not sure these are needed at all...
	'priority 1  /LSXML_FILE=
	'priority 2
	WriteKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SAPLOGON_LSXML_FILE", inifiles & "\SAPUILandscape.xml", "REG_SZ"
	'servers = C:\ProgramData\SAP\Common\SAPUILandscape.xml
	'priority 3
	'saplgpad <- important for servers
	WriteKey "HKEY_LOCAL_MACHINE\SOFTWARE\" & regkey & "SAP\SAPLogon\Options\PathConfigFilesLocal", inifiles, "REG_EXPAND_SZ"
	'saplogon
	WriteKey "HKEY_CURRENT_USER\Software\SAP\SAPLogon\Options\PathConfigFilesLocal", inifiles, "REG_EXPAND_SZ"
	DelKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SAPLOGON_INI_FILE", ""
Else
	'WriteKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SAPLOGON_INI_FILE", windowsdir & "\saplogon.ini", "REG_SZ"
	'fix for Trinseo - same place for 32/64?
	WriteKey "HKEY_CURRENT_USER\Software\SAP\SAPLogon\Options\PathConfigFilesLocal", inifiles, "REG_EXPAND_SZ"
End If
WriteKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SAPLOGON_INI_FILE", windowsdir & "\saplogon.ini", "REG_SZ"
WriteKey "HKEY_LOCAL_MACHINE\Software\" & regkey & "SAP\SAPGUI Front\SAP Frontend Server\Security\SecurityLevel", "0", "REG_DWORD"
WriteKey "HKEY_LOCAL_MACHINE\Software\" & regkey & "SAP\SAPGUI Front\SAP Frontend Server\Security\DefaultAction", "0", "REG_DWORD"
'2220930 - How to enable SAP UI Landscape format for SAP GUI for Windows 7.40
WriteKey "HKEY_LOCAL_MACHINE\SOFTWARE\" & regkey & "SAP\SAPLogon\LandscapeFormatEnabled", "1", "REG_DWORD"

'fix .sap not opening from the portal
'log("Fixing .sap association")
'progress 1,"Fixing .sap association"
If OSBIT = 64 Then
	'tx.sap.64.reg
	RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\tx.sap.64.reg"), FALSE)
	'temp
Else
	'tx.sap.reg
	RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\tx.sap.reg"), FALSE)
	'temp
End If


' Set trusted locations for MSWORD 
'MSDS
progress 6,"Making MSWORD Settings for EH&S"
log("Fixing MS word settings")
For Each officeversion in officeversions
	'Office 97   -  7.0 - ignored
	'Office 98   -  8.0 - ignored
	'Office 2000 -  9.0
	'Office XP   - 10.0
	'Office 2003 - 11.0
	'Office 2007 - 12.0
	'Office 2010 - 14.0 - first version also in 64bit
	'Office 2013 - 15.0
	'Office 2016 - 16.0
	If server Then
		'WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Word\Security\Trusted Locations\Location9\Path", sapguidir & "\SAPgui\wwi", "REG_SZ"
		'WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Word\Security\Trusted Locations\Location9\AllowSubfolders", "1", "REG_DWORD"
		' fix .WWI association
		'log("Fix WWI stuff")
		'the next file uses Office12 folder, which is silly
		RC = runthis(cmd, "/c REG IMPORT " & stringthis(sourcedir & "\wwi.fix.for.servers.reg"), FALSE)
	End If
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Word\Security\Trusted Locations\Location9\Path", sapguidir & "\SAPgui\wwi", "REG_SZ"
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Word\Security\Trusted Locations\Location9\AllowSubfolders", "1", "REG_DWORD"
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Word\Security\Trusted Locations\Location9\Description", "WWI", "REG_SZ"
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Excel\Security\Trusted Locations\Location9\Path", sapguidir, "REG_SZ"
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Excel\Security\Trusted Locations\Location9\AllowSubfolders", "1", "REG_DWORD"
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Excel\Security\Trusted Locations\Location9\Description", "SAPGUI", "REG_SZ"

	' permit access to VBOM
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Excel\Security\AccessVBOM", "1", "REG_DWORD"
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Word\Security\AccessVBOM", "1", "REG_DWORD"

	' fix a bug that prevents EH&S docs drawing in MSWORD
	WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Office\" & officeversion & "\Word\Options\ShowDrawings", "1", "REG_DWORD"
	'wwioptions has it too
	'hkey users\s-1xx
Next	

'disable personas
'2309640 - How to disable the Wide Layout Modus using SAP Screen Personas 
progress 3,"Disable Personas in registry"
If disablepersonas Then
	log("Disable Personas in registry")
	' [HKEY_CURRENT_USER\SOFTWARE\SAP\SAPGUI Front\SAP Frontend Server\Customize] “WideLayoutMode”=dword:00000000
	' ^^ higher priority
	' [HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP\SAPGUI Front\SAP Frontend Server\Customize] “WideLayoutMode”=dword:00000000
	WriteKey "HKEY_CURRENT_USER\SOFTWARE\SAP\SAPGUI Front\SAP Frontend Server\Customize\WideLayoutMode", "00000000", "REG_DWORD"
	WriteKey "HKEY_LOCAL_MACHINE\SOFTWARE\" & regkey & "SAP\SAPGUI Front\SAP Frontend Server\Customize\WideLayoutMode", "00000000", "REG_DWORD"
End If

' add *.sadarasvc.com to the local intranet sites list in IE and add sadarasvc.com to the DNS suffix list
'log("add *.sadarasvc.com to the local intranet sites list in IE and add sadarasvc.com to the DNS suffix list")
'removed 9/8/16 nsuter.
'progress 2,"Adding *.sadarasvc.com to IE"
'WriteKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sadarasvc.com\*", 1, "REG_DWORD"
'If KeyExists("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\SearchList") Then
	'instr 0 = case sensitive
	'keyval = ReadKey("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\SearchList")
	'If IsBlank(keyval,TRUE) Then
		'no searchlist. create it
		'right order: ad.sadara.com, sadarasvc.com
		'log("searchlist empty, adding ad.sadarasvc.com,sadarasvc.com")
		'keyval = keyval & "dow.com,intranet.dow.com,nam.dow.com,rohmhaas.net,rohmhaas.com,em.net,eur.dow.com,lam.dow.com,asa.dow.com,aus.dow.com,afr.dow.com,sct.ucarb.com,jvservices.com,jvservices-core.com,jvext.local,dowrand.com,bsnconnect.com,sadarasvc.com,ad.sadarasvc.com,sadarasvc.com"
		'WriteKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\SearchList", keyval, "REG_SZ"
	'End If		
	'If IsBlank(InStr(1, keyval, ",ad.sadarasvc.com", 1),TRUE) AND IsBlank(InStr(1, keyval, ",sadarasvc.com", 1),TRUE) Then
		'neither were found, add both to end
		'log("searchlist: Neither ,ad.sadarasvc.com or ,sadarasvc.com were found. Adding both to end")
		'keyval = keyval & ",ad.sadarasvc.com,sadarasvc.com"
		'WriteKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\SearchList", keyval, "REG_SZ"
	'End If
	'If NOT IsBlank(InStr(1, keyval, "sadarasvc.com", 1),TRUE) AND IsBlank(InStr(1, keyval, "ad.sadarasvc.com", 1),TRUE) Then
		'sadarasvc.com was found, but ad.sadarasvc.com was not found
		'log("searchlist: sadarasvc.com found. ad.sadarasvc.com not found. fixing key to be ,ad.sadarasvc.com,sadarasvc.com")
		'Replace(string,find,replacewith[,start[,count[,compare]]])
		'keyval = replace(keyval,"sadarasvc.com","ad.sadarasvc.com,sadarasvc.com")
		'WriteKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\SearchList", keyval, "REG_SZ"
	'End If
	'If IsBlank(InStr(1, keyval, ",sadarasvc.com", 1),TRUE) AND NOT IsBlank(InStr(1, keyval, "ad.sadarasvc.com", 1),TRUE) Then
		'ad.sadarasvc.com was found, but sadarasvc.com was not found
		'just append sadarasvc.com
		'log("searchlist: ad.sadarasvc.com found. ,sadarasvc.com not found. appending sadarasvc.com to end")
		'keyval = keyval & ",sadarasvc.com"
		'WriteKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\SearchList", keyval, "REG_SZ"
	'End If
	'If NOT IsBlank(InStr(1, keyval, ",sadarasvc.com", 1),TRUE) AND NOT IsBlank(InStr(1, keyval, "ad.sadara.com", 1),TRUE) Then
		'both exist,chekc order
		'right order: ad.sadara.com, sadarasvc.com
		'If InStr(1, keyval, "ad.sadarasvc.com", 1) < InStr(1, keyval, "sadara.com", 1) Then
			'right order
			'log("searchlist: ad.sadarasvc.com occurs before sadara.com. good. not doing anything")
		'Else
			'not the right order (it is sadarasvc.com,ad.sadara.com)
			'log("searchlist: ad.sadarasvc.com occurs after sadara.com. bad. fixing")
			'keyval = replace(keyval,"sadara.com","")
			'keyval = keyval & ",sadara.com"
			'WriteKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\SearchList", keyval, "REG_SZ"
		'End if
	'End If
'End If


' update WkstaUpdater.cfg
log("update WkstaUpdater.cfg")
progress 2,"Updating WkstaUpdater.cfg"
If NOT timeout Then
	copyfile sourcedir, "WkstaUpdater.cfg", SAPSetupDir & "\setup\Updater", ""
Else
	log("Stopping service timed out, not replacing WkstaUpdater.cfg")
End If


' Fix not registered mscomctl.ocx - not sure this is needed
'oShell.Run("%comspec% /C regsvr32 /s " & windowsdir & "\system32\mscomctl.ocx")
'oShell.Run("%comspec% /C regsvr32 /s " & windowsdir & "\system32\mscomct2.ocx")


'function to check if a file exists, and remove readonly/hidden/system if readonly is set
Function exists(file,remove)
	If FSO.FileExists(file) OR FSO.FolderExists(file) Then
		log("File or Folder Exists: " & file)
		found = TRUE
		If remove Then
			log("Setting ATTRIB 0 for " & file)
			On Error Resume Next
			Set f = FSO.GetFile(file)
			f.Attributes = 0
			If Err <> 0 Then 
				log("Unable to set ATTRIB to 0 for " & file)
				DisplayErrorInfo
				'found = FALSE
			Else 
				log("Successfully set ATTRIB to 0 for " & file)
			End If
			On Error GoTo 0
		End If
	Else
		log("File or Folder doesn't Exist: " & file)
		'DisplayErrorInfo
		found = FALSE
	End If
exists = found
End Function

'function to del a file, if it exists
Function delfile(file)
	found = FALSE
	If exists(file,0) Then
		On Error Resume Next
		FSO.Deletefile file, TRUE
		If Err <> 0 Then 
			log("Unable to delete: " & file)
			DisplayErrorInfo
			found = FALSE
		Else
			log("Deleted: " & file)
			found = TRUE
		End If
		On Error GoTo 0
	Else
		log("Doesn't exist, can't delete: " & file)
		found = FALSE
	End If
	delfile = found
End Function

'function to del a folder, if it exists
Function delfolder(foldername)
	found = FALSE
	On Error Resume Next
	If FSO.FolderExists(foldername) Then
		FSO.DeleteFolder foldername, TRUE
		If Err <> 0 Then 
			log("Unable to delete: " & foldername)
			DisplayErrorInfo
			found = FALSE
		Else
			log("Deleted: " & foldername)
			found = TRUE
		End If
		
	Else
		log("Doesn't exist, can't delete: " & foldername)
		found = FALSE
	End If
	delfolder = found
	On Error GoTo 0
End Function

'function to create a folder
Function createfolder(folder)
	created = FALSE
	If exists(folder,0) Then
		'already exists, do nothing
		created = TRUE
	Else
		'doesnt exist
		'recursive create folder
		createfolder(FSO.GetParentFolderName(folder))
		On Error Resume Next
		FSO.CreateFolder folder	
		If Err <> 0 Then 
			log("Unable to create: " & folder)
			DisplayErrorInfo
			created = FALSE
		Else
			log("Created: " & folder)
			created = TRUE
		End If
		On Error GoTo 0
	End If
createfolder = created
End Function


'function to copy a file
Function copyfile(fromdir, fromfile, todir, tofile)
	err.clear
	If tofile = "" Then tofile = fromfile End If
	If fromdir = "" Then fromdir = "." End If
	If todir = "" Then todir = fromdir End If
	copyfile = FALSE
	If NOT exists(fromdir & "\" & fromfile,0) Then
		'source file doesn't exist!
		log("Source file Doesn't exist: " & fromdir & "\" & fromfile)
		copyfile = FALSE
	Else
		'sourcefile exists. 
		createfolder(todir)
		'If InStr(tofile,"*") > 0 Then
			'tofile has a wild card, so just check dir
		'	exists todir,0
		'Else
		'	exists todir & "\" & tofile,1
		'End If
		log("Copying " & fromdir & "\" & fromfile & " ----> " & todir & "\" & tofile & " .....")
		On Error Resume Next
		'one day get this working
	 	'If fromfile = tofile Then
			'can use robocopy
		'	runthis windowsdir & "\system32\robocopy.exe", """" & fromdir & """" & " " & """" & todir & """" & " " & """" & tofile & """" & " /Z /MT /ETA", TRUE
		'Else
			FSO.CopyFile fromdir & "\" & fromfile, todir & "\" & tofile, TRUE
		'End If
		If Err <> 0 Then 
			'copy file failed
			log("Unable to copy file")
			DisplayErrorInfo
			copyfile = FALSE
		Else
			'copied fine
			log("File copied")
			copyfile = TRUE
		End If
		On Error GoTo 0
	End If	
End Function


'sub to show all the error details
Sub DisplayErrorInfo
	log("ERROR      : " & Err.Number)
	log("ERROR Hex  : " & Hex(Err.Number))
	log("ERROR Src  : " & Err.Source)
	log("ERROR Desc : " & Err.Description)
    Err.Clear
End Sub

'function to log messages to log file
Function log(message)
	If debug then
		msgbox(message)
	End If
	If logging then
		message = Now & " : " & userName & " : " & message
		On Error Resume Next
		logfile.WriteLine(message)
		On Error GoTo 0
	End If
End Function

'function to create uninstall routine in add/remove programs
Function addaddremove
	If instr(ucase(runningfrom),"DESKTOP") > 0 Then
		'running from desktop, don't create add/remove
	Else
		'create unless...
		If left(runningfrom,2) = "\\" Then
			'running from share, don't create add/remove
		Else
			'ok, create add/remove
			'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40
			'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40
			WriteKey "HKLM\Software\" & regkey &  "Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40\DisplayName","Dow SAPGUI 7.40","REG_SZ"
			WriteKey "HKLM\Software\" & regkey &  "Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40\UninstallString","wscript.exe " & stringthis(thisscript) & " uninstall local loglocal quieter","REG_SZ"
			WriteKey "HKLM\Software\" & regkey &  "Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40\Publisher","SAP AG","REG_SZ"
			WriteKey "HKLM\Software\" & regkey &  "Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40\EstimatedSize",810000,"REG_DWORD"
			WriteKey "HKLM\Software\" & regkey &  "Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40\DisplayVersion",version,"REG_SZ"
		End If
	End If
End Function

'function to remove uninstall routine in add/remove programs
Function removeaddremove
	DelKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40",""
	DelKey "HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40",""
End Function

progress 1,"Deleting dupe add/remove key"
'this doesnt work. fix it one day...
If KeyExists("KEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40\DisplayName\Dow SAPGUI 7.40") AND KeyExists("KEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40\DisplayName\Dow SAPGUI 7.40") Then
	'somehow 2 keys were created. only need one.
	DelKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Dow SAPGUI 7.40",""
End If
 
'function to log exit code, close file and quit script.
Function exitcode(code)
	log("Script completed.")
	log("code  : " & code)
	log("reboot: " & reboot)
	progress 3,"Script completed"

	'exit code statuses
	' -9 sapgui didn't install correctly
	' -8 sapgui didn't install correctly, recommend rebooting
	' -3 didnt install, will install on next reboot
	' -2 user chose to quit
	' -1 bad, something broke
	'  0 no status (yet)
	'  1 every-things good
	'  2 uninstall ran
	
	If NOT silent Then
		Select Case code
		Case 0
			log("Waiting for user: install complete")
			msgbox "SAPGUI Install and update is complete." & VbCrLf & "Use desktop shortcut 'Dow SAP Systems'" & VbCrLf & "or the shortcut under" & VbCrLf & "START -> All Programs -> SAP Front End."
		Case 1
			log("Waiting for user: install complete")
			msgbox "SAPGUI Install and update is complete." & VbCrLf & "Use desktop shortcut 'Dow SAP Systems'" & VbCrLf & "or the shortcut under" & VbCrLf & "START -> All Programs -> SAP Front End."	
		End Select
	End If
	
	Select Case code
	Case -3
		log("Waiting for user: install will run on reboot")
		msgbox "SAPGUI Install/update deferred until your next reboot"	
		reboot = 4
	Case -1 
		log("Waiting for user: install complete, but failed")
		msgbox "SAPGUI failed to install/update correctly." & VbCrLf & "Contact the GSD for support."' & VbCrLf & "Create a ticket for" & VbCrLf & "C-DOW-SAP-R3-TECH"
		reboot = 4
	Case -9 
		log("Waiting for user: install complete, but failed")
		msgbox "SAPGUI failed to install/update correctly." & VbCrLf & "Contact the GSD for support."' & VbCrLf & "Create a ticket for" & VbCrLf & "C-DOW-SAP-R3-TECH"
		reboot = 4
	Case -8
		log("Waiting for user: install failed, reboot")
		msgbox "SAPGUI failed to install/update correctly." & VbCrLf & "Contact the GSD for support."' & VbCrLf & "Recommend rebooting before trying again"
		reboot = 3
	Case -2 
		log("Waiting for user: user choose quit script")
		msgbox "User choose to abort install." & VbCrLf & "SAPGUI and updates NOT installed"
		reboot = 4
	Case 2
		log("Waiting for user: uninstall ran")
		msgbox "SAPGUI uninstall complete"
		reboot = 1
	End Select

			
	If reboot = 1 Then
		'If NOT silent Then
			log("Waiting for user: you should reboot before using the SAPGUI")
			msgbox("Recommend rebooting for the SAPGUI to work correctly")
		'End If
	ElseIf reboot = 2 Then
		'If NOT (silent) Then
			log("Waiting for user: you MUST reboot")
			msgbox("You MUST reboot for the SAPGUI to work correctly")
		'End If
	ElseIf reboot = 3 Then
		'If NOT (silent) Then
			log("Waiting for user: you should reboot before installing again.")
			msgbox("Recommend rebooting before attemping install again.")
		'End If
	Else
		log("Reboot not required")
	End If
	
	endTime = Now
	log("Script ran for: " & DateDiff("s", startTime, endTime) & " seconds")
	log("Exiting Script with " & code)
	If logging Then
		logfile.Close
	End If
	'oShell.run "cmd /c net use u: /delete"
	'runthis cmd, "/c net use u: /delete", FALSE
	'close the progress window
	If doprogress = TRUE Then
		On Error Resume Next
		window.quit
		set window = nothing
		On Error GoTo 0
	End If
	'RC = = successfull
	If code = 1 Then code = 0 End If
	wscript.quit code
End Function

'function to show a progress bar in a IE window
Function progress(increment,text)
	targetstage = currentstage + increment
	On Error Resume Next
	If doprogress = TRUE then
		window.document.title = currentstage & "% " &  text
		do until currentstage = targetstage
			window.document.write "|"
			wscript.sleep 10
			window.document.title = currentstage & "% " &  text
			currentstage = currentstage + 1
		loop
	End If
	On Error GoTo 0
End Function

'function to check for admin rights
Function CSI_IsAdmin()
	'Version 1.31
	'http://csi-windows.com/toolkit/csi-isadmin
	CSI_IsAdmin = FALSE
	On Error Resume Next
	'key = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
	RC = ReadKey("HKEY_USERS\S-1-5-19\Environment\TEMP")
	If RC <> "" Then CSI_IsAdmin = TRUE
	'If Err = 0 Then CSI_IsAdmin = True
	'key = ""
	On Error GoTo 0
End Function

'where does sapgui live
Function SAPGUIlocation()
	log("Looking for SAPGUI...")
	If KeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\SAP\SAP Shared\SAPDestDir") OR KeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP\SAP Shared\SAPDestDir") Then
		'x32: HKEY_LOCAL_MACHINE\SOFTWARE\SAP\SAP Shared\SAPDestDir = C:\Program Files\SAP\FrontEnd
		'x64: HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP\SAP Shared\ = C:\Program Files (x86)\SAP\FrontEnd
		'sapgui.exe lives in:  sapguidir\SAPgui\sapgui.exe
		'SAPSetupDir is C:\Program Files\SAP\SAPsetup  or  C:\Program Files (x64)\SAP\SAPsetup
		If OSBIT = 32 Then
			If KeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\SAP\SAP Shared\SAPDestDir") then sapguidir = ReadKey("HKEY_LOCAL_MACHINE\SOFTWARE\SAP\SAP Shared\SAPDestDir") End If
			If KeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\SAP\SAP Shared\SAPSetupDir") then SAPSetupDir = ReadKey("HKEY_LOCAL_MACHINE\SOFTWARE\SAP\SAP Shared\SAPSetupDir") End If
			regkey = ""
		Else
			If KeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP\SAP Shared\SAPDestDir") then sapguidir = ReadKey("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP\SAP Shared\SAPDestDir") End If
			If KeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP\SAP Shared\SAPSetupDir") then SAPSetupDir = ReadKey("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SAP\SAP Shared\SAPSetupDir") End If
			'^^this fails for acn. could be as it Doesn't exist.
			regkey = "Wow6432Node\"
		End if
		log("SAP Keys exist")
	Else
		log("SAP Keys do not exist")
			If FSO.FolderExists("C:\Program Files\SAP\FrontEnd") Then
				sapguidir = "C:\Program Files\SAP\FrontEnd"
				SAPSetupDir = "C:\Program Files\SAP\SapSetup"
				log("EXISTS: C:\Program Files\SAP\FrontEnd")
			ElseIf FSO.FolderExists("C:\Program Files (x86)\SAP\FrontEnd") Then
				sapguidir = "C:\Program Files (x86)\SAP\FrontEnd"
				SAPSetupDir = "C:\Program Files (x86)\SAP\SapSetup"
				log("EXISTS: C:\Program Files (x86)\SAP\FrontEnd")
			Else
				log("SAP folders do not exist")
			End IF
		End If
	If sapguidir = "" or SAPSetupDir = "" then
		'sapgui not found
		'set it based on os
		log("Setting SAP vars based on OSBIT")
		If OSBIT = 32 then
			sapguidir = "C:\Program Files\SAP\FrontEnd"
			SAPSetupDir = "C:\Program Files\SAP\SapSetup"
		Else 
			sapguidir = "C:\Program Files (x86)\SAP\FrontEnd"
			SAPSetupDir = "C:\Program Files (x86)\SAP\SapSetup"
		End If
	End If
	If FSO.FolderExists("C:\Program Files\SAP\NWBC40") Then
		NWBCdir = "C:\Program Files\SAP\NWBC40"
	Elseif FSO.FolderExists("C:\Program Files (x86)\SAP\NWBC40") Then
		NWBCdir = "C:\Program Files (x86)\SAP\NWBC40"
	End If
	If FSO.FolderExists("C:\Program Files\SAP\NWBC50") Then
		NWBCdir = "C:\Program Files\SAP\NWBC50"
	Elseif FSO.FolderExists("C:\Program Files (x86)\SAP\NWBC50") Then
		NWBCdir = "C:\Program Files (x86)\SAP\NWBC50"
	End If
	If NWBCdir = "" Then
		If OSBIT = 32 Then
			If version = 740 Then
				NWBCdir = "C:\Program Files\SAP\NWBC50"
			Else
				NWBCdir = "C:\Program Files\SAP\NWBC40"
			End If
		Else
			If version = 740 Then
				NWBCdir = "C:\Program Files (x86)\SAP\NWBC50"
			Else
				NWBCdir = "C:\Program Files (x86)\SAP\NWBC40"
			End If
		End If
	End If
	log("sapguidir : " & sapguidir)
	log("SAPSetupDir : " & SAPSetupDir)
	log("NWBCdir: " & NWBCdir)
End Function

'function to check if a reg key exists
Function KeyExists(key)
	log("Checking for key " & key)
    On Error Resume Next
	keyval = oShell.RegRead(key)
    If Err = 0 Then 
		KeyExists = True
		log("Key exists : " & key & " === " & keyval)
	Else
		'DisplayErrorInfo
		KeyExists = False
		keyval = ""
		log("Can't read Key: " & key)
	End If
	On Error GoTo 0
End Function

'function to write a reg key
Function WriteKey(key,keyval,keytype)
	'Create first??
	'objRegistry.CreateKey HKEY_CURRENT_USER,strKeyPath
	'objRegistry.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValue,strScriptStatus
	'If keyval = "" Then 
	'	key = key & "\"
	'	keytype = "REG_SZ"
	'End If
	log("Writing: " & key & " = " & keyval & " as type: " & keytype)
    On Error Resume Next
	'If keytype = "" Then
	'	oShell.RegWrite key, keyval
	'Else
		oShell.RegWrite key, keyval, keytype
	'End If
	'objShell.RegWrite strRegName, anyValue, [strType]
	'The Value is automatically converted to a string when keyval is REG_SZ or REG_EXPAND_SZ, and to an integer when keyval is REG_DWORD or REG_BINARY.
	'The optional keytype parameter specifies the data type for the value, valid options are REG_SZ, REG_EXPAND_SZ, REG_DWORD and REG_BINARY.
    If Err = 0 Then 
		log("Key Written OK")
	Else
		DisplayErrorInfo
		log("Failed to write Key")
	End If
	On Error GoTo 0
End Function

'function to read a reg key
Function ReadKey(key)
	'returns the key
	log("Reading: " & key)
    On Error Resume Next
	ReadKey = oShell.RegRead(key)
    If Err = 0 Then 
		log("Key Read OK, value is: " & ReadKey)
	Else
		'DisplayErrorInfo
		log("Failed to Read Key")
	End If
	On Error GoTo 0
End Function

'function to delete a reg key
Function DelKey(key,value)
	'returns true (success), or false (fail)
	'To delete a key instead of a value terminate key with a backslash character \
	'HKCU\Key\ will del the branch
	'key	left side tree
	'value	key name
	'data	data of key
	If value <> "" Then
		tempkey = key & "\" & value
	Else
		tempkey = key
	End If
	log("Deleting: " & tempkey)
    On Error Resume Next
	oShell.RegDelete(tempkey)
    If Err = 0 Then 
		log("Key Deleted OK")
		DelKey = TRUE
	Else
		DisplayErrorInfo
		log("Failed to Delete Key")
		DelKey = FALSE
	End If
	On Error GoTo 0
End Function

function getCount(counthis)
    on error resume next
	Err.clear
	getCounttemp = -1
	If IsArray(counthis) Then
		log("Is array")
		log("Counting " & counthis & " ...")
		getCounttemp = ubound(counthis) + 1
	Else
		log("Isn't a array")
		'this fails...
		'log("Counting " & counthis & " ...")
		getCounttemp = counthis.Count
	End If
	If Err <> 0 Then 
		'an error occurred
		'next line will fail if not an array
		'log("Error getting count of " & counthis)
		log("Error getting count")
		DisplayErrorInfo
		getCount = -1
	Else
		log("Counted : Found " & getCounttemp & " elements")
		getCount = getCounttemp
	End If
	'On error goto 0
End function

'function to get all details of the OS
Function osver()
	Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
'	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set osinfo = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	log("Operating System INFO:")
	For Each os in osinfo
		log("Boot Device: " & os.BootDevice)
		log("Build Number: " & os.BuildNumber)
		log("Build Type: " & os.BuildType)
		log("Caption: " & os.Caption)
		log("Code Set: " & os.CodeSet)
		log("Country Code: " & os.CountryCode)
		log("Debug: " & os.Debug)
		log("Encryption Level: " & os.EncryptionLevel)
		dtmConvertedDate.Value = os.InstallDate
		dtmInstallDate = dtmConvertedDate.GetVarDate
		log("Install Date: " & dtmInstallDate)
		log("Licensed Users: " & os.NumberOfLicensedUsers)
		log("Organization: " & os.Organization)
		log("OS Language: " & os.OSLanguage)
		log("OS Product Suite: " & os.OSProductSuite)
		log("OS Type: " & os.OSType)
		log("Primary: " & os.Primary)
		log("Product Type: " & os.ProductType)
		log("Registered User: " & os.RegisteredUser)
		log("Serial Number: " & os.SerialNumber)
		log("Version: " & os.Version)
		If os.ProductType = 3 AND left(os.Version,3) = "5.0" Then
			osystem = "WIN2000"
		Elseif os.ProductType = 1 AND (left(os.Version,3) = "5.1" OR left(os.Version,3) = "5.2") Then
			osystem = "XP"
		Elseif os.ProductType = 3 AND left(os.Version,3) = "5.2" Then
			osystem = "WIN2003"
		Elseif os.ProductType = 1 AND left(os.Version,3) = "6.0" Then
			osystem = "VISTA"
		Elseif os.ProductType = 1 AND left(os.Version,3) = "6.1" Then
			osystem = "WIN7"
		Elseif os.ProductType = 3 AND (left(os.Version,3) = "6.0" OR left(os.Version,3) = "6.1") Then
			osystem = "WIN2008"
		Elseif os.ProductType = 3 AND (left(os.Version,3) = "6.2" OR left(os.Version,3) = "6.3") Then
			osystem = "WIN2012"
		Elseif os.ProductType = 1 AND left(os.Version,3) = "6.2" Then
			osystem = "WIN8"
		Elseif os.ProductType = 1 AND left(os.Version,3) = "6.3" Then
			osystem = "WIN81"
		Elseif os.ProductType = 3 AND left(os.Version,3) = "6.3" Then
			osystem = "WIN2012R3"
		Elseif os.ProductType = 1 AND left(os.Version,4) = "10.0" Then
			osystem = "WIN10"
		Else 
			osystem = "UNKNOWN"
		End If
	Next
	log("Operating System (osystem): " & osystem)
End Function

'function to check for office
Function checkoffice()
	'HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall
	'2013 32bit:"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{91150000-0011-0000-0000-0000000FF1CE}"
	'2013 64bit:"HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{91150000-0011-0000-0000-0000000FF1CE}"
	'2010 32bit:"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{91140000-0011-0000-0000-0000000FF1CE}"
	'2010 64bit:"HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{91140000-0011-0000-0000-0000000FF1CE}"
	found = FALSE
	For Each regentry IN ARRAY("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe\Path","HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\Path") 
		If KeyExists(regentry) Then
			keyval = ReadKey(regentry)
			If exists(keyval,0) Then
				found = TRUE
			End If
		End If
	Next
	'check versions of office
	'HKLM\Software\Microsoft\Office\12.0\Word\InstallRoot::Path
	'HKLM\Software\Wow6432Node\Microsoft\Office\12.0\Word\InstallRoot::Path
	'check details.version & bit from .exe
	If found Then
		log("Office exists")
		checkoffice = TRUE
	Else
		log("Office Doesn't exist")
		checkoffice = FALSE
		If ignoreoffice Then
			checkoffice = TRUE
		End If
	End If
End Function

'function to work out type of computer
Function checkcomputer()
	log("Checking type of computer")
	Set objWMIService2 = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Set colComputers = objWMIService2.ExecQuery ("SELECT * FROM Win32_ComputerSystem")
	log("Checked Win32_ComputerSystem")
	GetCount(colComputers)
	For Each objComputer in colComputers
		log("Domain Role : " & objComputer.DomainRole)
		Select Case objComputer.DomainRole
		Case 0
			computertype = "Standalone Workstation"
			server = FALSE
		Case 1 
			computertype = "Member Workstation"
			server = FALSE
		Case 2
			computertype = "Standalone Server"
			server = TRUE
		Case 3
			computertype = "Member Server"
			server = TRUE
		Case 4
			computertype = "Backup Domain Controller"
			server = TRUE
		Case 5
			computertype = "Primary Domain Controller"
			server = TRUE
		Case Else
			computertype = "Can't tell."
		End Select	
	Next
	log("Type of computer: " & computertype)
	If KeyExists("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Citrix\ProductName") Then
		citrix = TRUE
		log("Is a Citrix server")
	Else
		log("Not a Citrix server")
	End If
End Function

'function to run programs
Function runthis(prgtorun,params,waitonreturn)
	'waitonreturn - whether the script should wait for the program to finish executing before continuing to the next statement in your script. If set to true, script execution halts until the program finishes, and Run returns any error code returned by the program. If set to false (the default), the Run method returns immediately after starting the program, automatically returning 0 (not to be interpreted as an error code)
	'cmd /C     Run Command and then terminate
	'cmd /K     Run Command and then return to the CMD prompt. This is useful for testing, to examine variables
	numerr = -1
	'tempprgtorun = stringthis(prgtorun) 'breaks exists. So do after?
	If exists(prgtorun,0) OR prgtorun = start Then
		'isrunning(prgtorun) 'check if already running. lol that would be nice.
		If waitonreturn Then
			log("Going to run: " & prgtorun & " " & params)
		Else
			log("Going to run (without waiting): " & prgtorun & " " & params)
		End If
		On Error Resume Next
		If prgtorun = start Then
			If waitonreturn Then
				numerr = oShell.run("start /WAIT /B " & stringthis("Starting...") & " " & params) 'numerr will be ?
			Else
				numerr = oShell.run("start /B " & stringthis("Starting...") & " " & params) 'numerr will be ?
			End If
		Else
			numerr = oShell.run(stringthis(prgtorun) & " " & params, 0, waitonreturn) 'numerr will be 0 if waitonreturn is false
		End If
		'-1				not found?
		'0				program ran, or program ran & it returned 0
		'<initial>		program not found
		'-2147024894 	program not found
		'other			RC from program
		'
		'or
		'oShell.ShellExecute(sFile [, vArguments] [, vDirectory] [, vOperation] [, vShow])
		'numerr = oShell.ShellExecute(sFile [, vArguments] [, vDirectory] [, vOperation] [, vShow])
		'sFile	Required. A String that contains the name of the file on which ShellExecute will perform the action specified by vOperation.
		'vArguments	Optional. A Variant that contains the parameter values for the operation.
		'vDirectory	Optional. A Variant that contains the fully qualified path of the directory that contains the file specified by sFile. If this parameter is not specified, the current working directory is used.
		'vOperation	Optional. A Variant that specifies the operation to be performed. It should be set to one of the verb strings that is supported by the file. For a discussion of verbs, see the Remarks section. If this parameter is not specified, the default operation is performed.
		'vShow	Optional. A Variant that recommends how the window that belongs to the application that performs the operation should be displayed initially. The application can ignore this recommendation. vShow can take one of the following values. If this parameter is not specified, the application uses its default value.
		'0	Open the application with a hidden window.
		'1	Open the application with a normal window. If the window is minimized or maximized, the system restores it to its original size and position.
		'2	Open the application with a minimized window.
		'3	Open the application with a maximized window.
		'4	Open the application with its window at its most recent size and position. The active window remains active.
		'5	Open the application with its window at its current size and position.
		'7	Open the application with a minimized window. The active window remains active.
		'10	Open the application with its window in the default state specified by the application.
		If err <> 0 Then 
			log("RC from program (numerr): " & numerr)
			log("ERROR running " & prgtorun & " " & params)
			DisplayErrorInfo
		Else
			log("RC from program (numerr): " & numerr)
		End If
		On Error Goto 0
	Else
		log("Didn't exist, didn't run: " & prgtorun)
	End If
	runthis = numerr
	numerr = 0
End Function

'oshell.exec example:
'Dim objShell,oExec
'Set objShell = wscript.createobject("wscript.shell")
'Set oExec = objShell.Exec("calc.exe")
'WScript.Echo oExec.Status
'WScript.Echo oExec.ProcessID
'WScript.Echo oExec.ExitCode 
'
'while(!Pipe.StdOut.AtEndOfStream)
'WScript.StdOut.WriteLine(Pipe.StdOut.ReadLine())
'you can read the programs stdout/stderr 
'
'Do While oExec.Status = 0
    'WScript.Sleep 100
	'sleep until program completes.
'Loop

'returns True if Empty or NULL or Zero
Function IsBlank(Value,quiet)
	If IsEmpty(Value) or IsNull(Value) Then
		IsBlank = True
	ElseIf VarType(Value) = vbString Then
		If Value = "" Then
			IsBlank = True
		End If
	ElseIf IsObject(Value) Then
		If Value Is Nothing Then
			IsBlank = True
		End If
	ElseIf IsNumeric(Value) Then
		If Value = 0 Then
			IsBlank = True
		End If
	Else
		IsBlank = False
	End If
	If IsBlank = TRUE AND quiet = FALSE Then
		log("Is BLANK: " & Value)
	ElseIf IsBlank = FALSE AND quiet = FALSE Then
		log("Is NOT BLANK: " & Value)
	End If
End Function

Function stringthis(strings)
	stringthis = ""
	If IsArray(strings) Then
		For each thing in strings
			stringthis = stringthis + """" + thing + """"
		Next
	Else
		stringthis = """" + strings + """"
	End If
End function

Function GetTimeZoneOffset()
    Set cItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each oItem In cItems
        GetTimeZoneOffset = oItem.CurrentTimeZone / 60
        Exit For
    Next
	log("Hours ahead of UTC : " & GetTimeZoneOffset)
	Set colItems = objWMIService.ExecQuery("Select * from Win32_TimeZone")
	For Each objItem in colItems
		'log "Bias: " & objItem.Bias
		'log "Caption: " & objItem.Caption
		'log "Daylight Bias: " & objItem.DaylightBias
		'log "Daylight Day: " & objItem.DaylightDay
		'log "Daylight Day Of Week: " & objItem.DaylightDayOfWeek
		'log "Daylight Hour: " & objItem.DaylightHour
		'log "Daylight Millisecond: " & objItem.DaylightMillisecond
		'log "Daylight Minute: " & objItem.DaylightMinute
		'log "Daylight Month: " & objItem.DaylightMonth
		log "Daylight Name: " & objItem.DaylightName
		'log "Daylight Second: " & objItem.DaylightSecond
		'log "Daylight Year: " & objItem.DaylightYear
		'log "Description: " & objItem.Description
		'log "Setting ID: " & objItem.SettingID
		'log "Standard Bias: " & objItem.StandardBias
		'log "Standard Day: " & objItem.StandardDay
		'log "Standard Day Of Week: " & objItem.StandardDayOfWeek
		'log "Standard Hour: " & objItem.StandardHour
		'log "Standard Millisecond: " & objItem.StandardMillisecond
		'log "Standard Minute: " & objItem.StandardMinute
		'log "Standard Month: " & objItem.StandardMonth
		'log "Standard Name: " & objItem.StandardName
		'log "Standard Second: " & objItem.StandardSecond
		'log "Standard Year: " & objItem.StandardYear
	Next
End Function

exitcode(status)

'End If