#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\carmine.santoro\Pictures\21.ico
#AutoIt3Wrapper_Outfile_x64=OfficeOnlyInstall_no_updates.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=Office Installer unattended
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_ProductName=Office Installer
#AutoIt3Wrapper_Res_ProductVersion=1.0.0.1
#AutoIt3Wrapper_Res_CompanyName=CincinnatiState
#AutoIt3Wrapper_Res_LegalCopyright=Carm0
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <TrayConstants.au3> ; Required for the $TRAY_CHECKED and $TRAY_ICONSTATE_SHOW constants.
If UBound(ProcessList(@ScriptName)) > 2 Then Exit
TraySetToolTip("Installer")
HotKeySet("^!m", "MyExit") ; ctrl+Alt+m kills program ( hotkey )
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.
TrayCreateItem("About")
TrayItemSetOnEvent(-1, "About")
TrayCreateItem("") ; Create a separator line.
TrayCreateItem("Exit")
TrayItemSetOnEvent(-1, "ExitScript")
TraySetOnEvent($TRAY_EVENT_PRIMARYDOUBLE, "About") ; Display the About MsgBox when the tray icon is double clicked on with the primary mouse button.
TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.

SplashTextOn("Working", "", 220, 60, -1, -1, 16, "Tahoma", 10)
ControlSetText("Working", "", "Static1", "Office 2016", 2)
FileCopy(@ScriptDir & '\Office2016\office.7z', 'c:\windows\temp\office.7z', 1)
Do
	Sleep(500)
Until FileExists('c:\windows\temp\office.7z')
Sleep(3000)
ControlSetText("Working", "", "Static1", "Extracting Office 2016 package", 2)
ShellExecuteWait('C:\Program Files\7-Zip\7z.exe', ' x office.7z', 'C:\Windows\temp\', "", @SW_HIDE)
Sleep(3000)
ControlSetText("Working", "", "Static1", "Installing Office 2016 package", 2)
ShellExecuteWait('setup.exe', " /adminfile Carmine1.msp", 'c:\windows\temp\office', "", @SW_HIDE)
Sleep(3000)
FileDelete('c:\windows\temp\office.7z')

ControlSetText("Working", "", "Static1", "Extracting Visio/Project", 2)
FileCopy(@ScriptDir & '\Office2016\V-P.7z', 'c:\windows\temp\V-P.7z', 1)
Do
	Sleep(3000)
Until FileExists('c:\windows\temp\V-P.7z')
ShellExecuteWait('C:\Program Files\7-Zip\7z.exe', ' x V-P.7z', 'C:\Windows\temp\', "", @SW_HIDE)
ControlSetText("Working", "", "Static1", "Project", 2)
ShellExecuteWait('setup.exe', " /adminfile Project2016.msp", 'c:\windows\temp\Project2016', "", @SW_HIDE)
Sleep(3000)
ControlSetText("Working", "", "Static1", "Visio", 2)
ShellExecuteWait('setup.exe', " /adminfile Visio2016.msp", 'c:\windows\temp\Visio2016', "", @SW_HIDE)
Sleep(3000)
DirRemove('C:\windows\temp',1)

;FileDelete('C:\Users\Public\Desktop\Outlook 2016.lnk')
;FileDelete('C:\Users\Public\Desktop\Outlook 2016.lnk')
;FileDelete('C:\Users\Public\Desktop\Project 2016.lnk')
;FileDelete('C:\Users\Public\Desktop\Visio 2016.lnk')
;FileDelete('C:\Users\Public\Desktop\PowerPoint 2016.lnk')

#cs
SplashTextOn("Working", "", 220, 60, -1, -1, 16, "Tahoma", 10)
ControlSetText("Working", "", "Static1", "Copying Office Updates to temp", 2)
FileCopy(@ScriptDir & '\Office2016\OfficeUpdates.rar', 'c:\windows\temp\OfficeUpdates.rar', 1)
Do
	Sleep(500)
Until FileExists('c:\windows\temp\OfficeUpdates.rar')
Sleep(3000)
ControlSetText("Working", "", "Static1", "Extracting Office 2016 updates", 2)
ShellExecuteWait('C:\Program Files\7-Zip\7z.exe', ' x OfficeUpdates.rar', 'C:\Windows\temp\', "", @SW_HIDE)
SplashOff()
Sleep(3000)
ShellExecuteWait('OfficeUpdates.exe', "", @ScriptDir, "", @SW_HIDE)
#ce

Func MyExit()
	Exit
EndFunc   ;==>MyExit

Func About()
	; Display a message box about the AutoIt version and installation path of the AutoIt executable.
	MsgBox(0, "", "Office 2016 installer" & @CRLF & @CRLF & _
			"Version: 1.0.0.1" & @CRLF & _
			"Installer" & @CRLF & "CTRL+ALT+m to kill", 5) ; Find the folder of a full path.
EndFunc   ;==>About

Func ExitScript()
	Exit
EndFunc   ;==>ExitScript
