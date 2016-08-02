'
' AutoPlayer 0.1.0
' AutoDJ script to play higher-rated tracks more often
' Created by eucal
'
'
' APInstaller.vbs: AutoPlayer install / uninstall script
'
Option Explicit


'
' Constant definitions
'
Const DefaultMinSpacingUnr = 105 ' Minimum time (days) between repeats of the same song (unknwn rating)
Const DefaultMinSpacingNew = 10	 ' For unskipped songs (SkipCount = 0)
Const DefaultMinSpacing50  = 30  ' For 5-star tracks
Const DefaultMinSpacing45  = 45
Const DefaultMinSpacing40  = 60
Const DefaultMinSpacing35  = 75
Const DefaultMinSpacing30  = 90
Const DefaultMinSpacing25  = 105
Const DefaultMinSpacing20  = 150
Const DefaultMinSpacing15  = 200
Const DefaultMinSpacing10  = 250
Const DefaultMinSpacing05  = 325
Const DefaultMinSpacing00  = 365 ' Bomb rating

Const ScriptName = "AutoPlayer"


'
' Writes a value if it does not exist in the ini file already.
'
Sub WriteIniIfNotExists(ini, section, key, val)
	If ini Is Nothing Then
		SDB.MessageBox "Could not write to ini file.", mtError, Array(mbOK)
		Exit Sub
	End If

	If (ini.StringValue(section, key) <> "") Then Exit Sub ' Do not overwrite existing values
	
	WriteIni ini, section, key, val
End Sub

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
	Dim mmIni      : Set mmIni      = SDB.IniFile
	
	WriteIni scriptsIni, ScriptName, "DisplayName", ScriptName
	WriteIni scriptsIni, ScriptName, "ScriptType",  4
	WriteIni scriptsIni, ScriptName, "FileName",  ScriptName & "\APMain.vbs"
	WriteIni scriptsIni, ScriptName, "Language",    "VBScript"
	
	' set default values; preserve settings
	Dim rootPath : rootPath = SDB.ScriptsPath & ScriptName & "\"
	SDB.Tools.FileSystem.CreateFolder rootPath
	
	WriteIniIfNotExists mmIni, ScriptName, "RootPath", rootPath
	
	' Create AutoPlayer.ini
	Dim iniPath : iniPath = rootPath & ScriptName & ".ini"
	Dim ini : Set Ini = SDB.Tools.IniFileByPath(iniPath)
	
	If ini Is Nothing Then
		SDB.Tools.FileSystem.CreateTextFile(iniPath)
		ini = SDB.Tools.IniFileByPath(iniPath)
		
		If ini Is Nothing Then
			SDB.MessageBox "Error: Could not create ini file!", mtError, Array(mbOK)
		End If
	End If
	
	' Preserve settings when reinstalling
	WriteIniIfNotExists ini, "Spacing", "MinSpacingUnr", DefaultMinSpacingUnr
	WriteIniIfNotExists ini, "Spacing", "MinSpacingNew", DefaultMinSpacingNew
	WriteIniIfNotExists ini, "Spacing", "MinSpacing50",  DefaultMinSpacing50
	WriteIniIfNotExists ini, "Spacing", "MinSpacing45",  DefaultMinSpacing45
	WriteIniIfNotExists ini, "Spacing", "MinSpacing40",  DefaultMinSpacing40
	WriteIniIfNotExists ini, "Spacing", "MinSpacing35",  DefaultMinSpacing35
	WriteIniIfNotExists ini, "Spacing", "MinSpacing30",  DefaultMinSpacing30
	WriteIniIfNotExists ini, "Spacing", "MinSpacing25",  DefaultMinSpacing25
	WriteIniIfNotExists ini, "Spacing", "MinSpacing20",  DefaultMinSpacing20
	WriteIniIfNotExists ini, "Spacing", "MinSpacing15",  DefaultMinSpacing15
	WriteIniIfNotExists ini, "Spacing", "MinSpacing10",  DefaultMinSpacing10
	WriteIniIfNotExists ini, "Spacing", "MinSpacing05",  DefaultMinSpacing05
	WriteIniIfNotExists ini, "Spacing", "MinSpacing00",  DefaultMinSpacing00

	Set scriptsIni = Nothing
	Set mmIni = Nothing
	
	SDB.RefreshScriptItems
End Function


Function FinishInstall
	SDB.Tools.FileSystem.DeleteFile SDB.ScriptsPath & "APInstaller.vbs"
End Function


Sub DeletePath(folderPath)
	' Delete trailing backslash
	With (New RegExp)
		.Global = True
		.Pattern = "\\$"
		folderPath = .Replace(folderPath, "")
	End With
	
	msgbox "Deleting folder " & folderPath, vbOK
	
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
	' since it is hardcoded and changes when changin drive letters etc.
	mmIni.DeleteSection ScriptName 
	
	Set mmIni = Nothing
	Set scriptsIni = Nothing
		
	' Refresh script items to remove control panel and menu item
	SDB.RefreshScriptItems
End Function
