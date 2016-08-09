'
' AutoPlayer 0.1.0
' AutoDJ script to play higher-rated tracks more often
' Created by eucal
'
'
' APInstaller.vbs: AutoPlayer install / uninstall script
'
Option Explicit

Const ScriptName = "AutoPlayer"


Sub WriteIni(ini, section, key, val)
	If ini Is Nothing Then
		SDB.MessageBox "Could not write to ini file.", mtError, Array(mbOK)
		Exit Sub
	End If
	
	Select Case vartype(val)
	Case vbInteger
		ini.IntValue(section, key) = val
	Case vbString
		ini.StringValue(section, key) = val
	Case vbBoolean
		ini.BoolValue(section, key) = val
	Case Else
		SDB.MessageBox "Could not write object type " & typename(val), mtError, Array(mbOK)
	End Select
End Sub

'
' Installation routine
'
Function BeginInstall
	Dim scriptsIni : Set scriptsIni = SDB.Tools.IniFileByPath(SDB.ScriptsPath & "Scripts.ini")

	WriteIni scriptsIni, ScriptName, "DisplayName", ScriptName
	WriteIni scriptsIni, ScriptName, "ScriptType",  4
	WriteIni scriptsIni, ScriptName, "FileName",  ScriptName & "\APMain.vbs"
	WriteIni scriptsIni, ScriptName, "Language",    "VBScript"
	
	' set default values; preserve settings
	Dim rootPath : rootPath  = SDB.ScriptsPath & ScriptName & "\"
	Dim iniPath  : iniPath   = rootPath & ScriptName & ".ini"
	Dim mmIni    : Set mmIni = SDB.IniFile
	WriteIni mmIni, ScriptName, "RootPath", rootPath
	
	If Not SDB.Tools.FileSystem.FileExists(iniPath) Then
		' If the ini does not exist, it means that the program is not installed
		' or was uninstalled without preserving settings. So we have to create the root folder, too
		SDB.Tools.FileSystem.CreateFolder rootPath
		SDB.Tools.FileSystem.CreateTextFile(iniPath)
		
		If Not SDB.Tools.FileSystem.FileExists(iniPath) Then
			SDB.MessageBox "Error: Could not create ini file!", mtError, Array(mbOK)
			BeginInstall = -1
			Exit Function
		End If
	End If
	
	SDB.RefreshScriptItems
End Function


Function FinishInstall
	SDB.Tools.FileSystem.DeleteFile SDB.ScriptsPath & "APInstaller.vbs"
End Function


Sub RegExpReplace(ByRef str, ByVal pattern, ByVal replacement)
	With (New RegExp)
		.Global = True
		.Pattern = pattern
		str = .Replace(str, replacement)
	End With
End Sub

	
Sub DeletePath(folderPath)
	' Delete trailing backslash if present
	RegExpReplace folderPath, "\\$", ""
	
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(folderPath) Then
		fso.DeleteFolder(folderPath)
	End If
End Sub


'
' Uninstallation routine
'
Function FinishUninstall
	Dim msgDeleteSettings : msgDeleteSettings = "Do you want to remove " &_
		ScriptName & " settings as well?" & vbNewLine &_
		"If you click No, script settings will be left in AutoPlayer.ini"
	
	Dim deleteSettings
	If (SDB.MessageBox(SDB.Localize(msgDeleteSettings), mtConfirmation, Array(mbYes, mbNo)) = mrYes) Then
		deleteSettings = True
	Else
		deleteSettings = False
	End If
	
	Dim mmIni : Set mmIni = SDB.IniFile
	If mmIni Is Nothing Then
		SDB.MessageBox "Could not load Mediamonkey.ini file! Uninstallation could not be completed!", mtError, Array(mbOK)
		BeginUninstall = -1
		Exit Function
	End If

	Dim rootPath : rootPath = mmIni.StringValue(ScriptName, "RootPath")
	If deleteSettings Then ' just delete everything
		DeletePath rootPath
	End If
	
	' delete settings from scripts.ini
	Dim scriptsIni : Set scriptsIni = SDB.Tools.IniFileByPath(SDB.ScriptsPath & "Scripts.ini")
	If Not scriptsIni Is Nothing Then scriptsIni.DeleteSection ScriptName

	' delete root path from Mediamonkey.ini
	' since it is hardcoded and changes when changing drive letters etc.
	mmIni.DeleteSection ScriptName 
	
	' Refresh script items to remove control panel and menu item
	SDB.RefreshScriptItems
End Function
