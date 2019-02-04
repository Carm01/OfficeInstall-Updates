#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\carmine.santoro\Downloads\Google-Noto-Emoji-People-Family-Love-12148-purple-heart.ico
#AutoIt3Wrapper_Outfile_x64=OfficeUpdates.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=Office Updates
#AutoIt3Wrapper_Res_Fileversion=1.0.0.9
#AutoIt3Wrapper_Res_ProductName=Office Updates
#AutoIt3Wrapper_Res_ProductVersion=1.0.0.9
#AutoIt3Wrapper_Res_LegalCopyright=Carm0
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <File.au3>
#include <Array.au3>
#include <Debug.au3>
#include <WinAPI.au3>
;~ #include <WinAPISys.au3>
#include <WinAPIShellEx.au3>
#include <SendMessage.au3>

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

If Not FileExists('c:\windows\temp\Updates') Then
	SplashTextOn("Processing Office Updates", "", 275, 90, -1, -1, 16, "Tahoma", 10)
	ControlSetText("Processing Office Updates", "", "Static1", 'Copying Office Updates', 2)
	FileCopy(@ScriptDir & '\Office2016\OfficeUpdates.rar', 'c:\windows\temp\OfficeUpdates.rar', 8)
	Do
		Sleep(500)
	Until FileExists('c:\windows\temp\OfficeUpdates.rar')
	ControlSetText("Processing Office Updates", "", "Static1", 'Extracting Office Updates', 2)
	ShellExecuteWait('C:\Program Files\7-Zip\7z.exe', ' x OfficeUpdates.rar', 'C:\Windows\temp\', "", @SW_HIDE)
	Sleep(1000)
EndIf

$a = _FileListToArrayRec('c:\windows\temp\Updates', '*.msp', $FLTAR_FILES, $FLTAR_NORECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
Dim $y[UBound($a)][2]
For $i = 0 To UBound($a) - 1
	;MsgBox(0,"", $i)
	$z = FileGetTime($a[$i], $FT_MODIFIED, $FT_STRING)
	$y[$i][0] = $a[$i]
	$y[$i][1] = $z
Next
_ArraySort($y, 0, 0, 0, 1)
;_DebugArrayDisplay($y)


SplashTextOn("Processing Office Updates", "", 275, 90, -1, -1, 16, "Tahoma", 10)
For $i = 1 To UBound($y) - 1
	$b = StringSplit($y[$i][0], '\')
	$c = UBound($b) - 1
	$d = $b[$c]
	ControlSetText("Processing Office Updates", "", "Static1", $i & ' of ' & UBound($y) - 1 & @CRLF & $d, 2)
	ShellExecuteWait($y[$i][0], ' /passive /quiet /norestart', "c:\windows\temp\Updates", "", @SW_HIDE)
	If Not ProcessExists('explorer.exe') Then
		ConsoleWrite("! Explorer PID: " & _Restart_Explorer() & "  *  Error: " & @error & @CRLF)
	EndIf
Next
SplashOff()

;ConsoleWrite("! Explorer PID: " & _Restart_Explorer() & "  *  Error: " & @error & @CRLF)
Sleep(5000)
DirRemove('c:\windows\temp\Updates', 1)
FileDelete('c:\windows\temp\OfficeUpdates.rar')
Run("shutdown -r -f -t 0 -m \\" & @ComputerName, @SystemDir, @SW_HIDE)

Func _Restart_Explorer()
	Local $ifailure = 100, $zfailure = 100, $rPID = 0, $iExplorerPath = @WindowsDir & "\Explorer.exe"
	_WinAPI_ShellChangeNotify($shcne_AssocChanged, 0, 0, 0) ; Save icon positions
	Local $hSystray = _WinAPI_FindWindow("Shell_TrayWnd", "")
	_SendMessage($hSystray, 1460, 0, 0) ; Close the Explorer shell gracefully
	While ProcessExists("Explorer.exe") ; Try Close the Explorer
		Sleep(10)
		$ifailure -= ProcessClose("Explorer.exe") ? 0 : 1
		If $ifailure < 1 Then Return SetError(1, 0, 0)
	WEnd
	RegDelete("HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\TrayNotify", "IconStreams")
	RegDelete("HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\TrayNotify", "PastIconsStream")
;~  _WMI_StartExplorer()
	While (Not ProcessExists("Explorer.exe")) ; Start the Explorer
		If Not FileExists($iExplorerPath) Then Return SetError(-1, 0, 0)
		Sleep(500)
		$rPID = ShellExecute($iExplorerPath)
		$zfailure -= $rPID ? 0 : 1
		If $zfailure < 1 Then Return SetError(2, 0, 0)
	WEnd
	Return $rPID
EndFunc   ;==>_Restart_Explorer


Func MyExit()
	Exit
EndFunc   ;==>MyExit

Func About()
	; Display a message box about the AutoIt version and installation path of the AutoIt executable.
	MsgBox(0, "", "Office Updater" & @CRLF & @CRLF & _
			"Version: 1.0.0.9" & @CRLF & _
			"Installer" & @CRLF & "CTRL+ALT+m to kill", 5) ; Find the folder of a full path.
EndFunc   ;==>About

Func ExitScript()
	Exit
EndFunc   ;==>ExitScript

#cs
	;_ArrayDisplay($a)
	SplashTextOn("Working", "", 275, 90, -1, -1, 16, "Tahoma", 10)
	For $i = 1 To UBound($a) - 1
	$b = StringSplit($a[$i], '\')
	$c = UBound($b) - 1
	$d = $b[$c]
	ControlSetText("Working", "", "Static1", $i & ' of ' & UBound($a) - 1 & @CRLF & $d, 2)
	ShellExecuteWait($a[$i], ' /passive /quiet /norestart', "", "", @SW_HIDE)
	Next
#ce
; https://docs.microsoft.com/en-us/previous-versions/office/office-2013-resource-kit/cc178995(v=office.15)
; http://supportishere.com/how-to-download-microsoft-office-2013-updates-the-easy-way/
