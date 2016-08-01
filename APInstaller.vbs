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
Sub WriteIfNotExists(ini, section, key, val)
	If (ini.StringValue(section, key) <> "") Then Exit Sub ' Do not overwrite existing values
	
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
	' Add entries to script.ini if you need to show up in the Scripts menu
	Dim inif : Set inif = SDB.Tools.IniFileByPath(SDB.CurrentAddonInstallRoot & "Scripts\Scripts.ini")
	If Not (inif Is Nothing) Then
		inif.StringValue(ScriptName, "DisplayName") = ScriptName
		inif.IntValue   (ScriptName, "ScriptType")  = 4
		inif.StringValue(ScriptName, "FileName")    = ScriptName & "\APMain.vbs"
		inif.StringValue(ScriptName, "Language")    = "VBScript"
	End If

	Dim Ini : Set Ini = SDB.IniFile
	
	' set default values; preserve settings
	WriteIfNotExists Ini, ScriptName, "RootPath", SDB.CurrentAddonInstallRoot & "Scripts\AutoPlayer\"
	WriteIfNotExists Ini, ScriptName, "MinSpacingUnr", DefaultMinSpacingUnr
	WriteIfNotExists Ini, ScriptName, "MinSpacingNew", DefaultMinSpacingNew
	WriteIfNotExists Ini, ScriptName, "MinSpacing50",  DefaultMinSpacing50
	WriteIfNotExists Ini, ScriptName, "MinSpacing45",  DefaultMinSpacing45
	WriteIfNotExists Ini, ScriptName, "MinSpacing40",  DefaultMinSpacing40
	WriteIfNotExists Ini, ScriptName, "MinSpacing35",  DefaultMinSpacing35
	WriteIfNotExists Ini, ScriptName, "MinSpacing30",  DefaultMinSpacing30
	WriteIfNotExists Ini, ScriptName, "MinSpacing25",  DefaultMinSpacing25
	WriteIfNotExists Ini, ScriptName, "MinSpacing20",  DefaultMinSpacing20
	WriteIfNotExists Ini, ScriptName, "MinSpacing15",  DefaultMinSpacing15
	WriteIfNotExists Ini, ScriptName, "MinSpacing10",  DefaultMinSpacing10
	WriteIfNotExists Ini, ScriptName, "MinSpacing05",  DefaultMinSpacing05
	WriteIfNotExists Ini, ScriptName, "MinSpacing00",  DefaultMinSpacing00
	
	SDB.Tools.FileSystem.CreateFolder(SDB.CurrentAddonInstallRoot & "Scripts\" & ScriptName)
	SDB.RefreshScriptItems
End Function


'
' Uninstallation routine
'
Function BeginUninstall
	Dim MsgDeleteSettings : MsgDeleteSettings = "Do you want to remove " &_
		ScriptName & " settings as well?" & vbNewLine &_
		"If you click No, script settings will be left in MediaMonkey.ini"
	
	Dim deleteSettings
	deleteSettings = (SDB.MessageBox(SDB.Localize(MsgDeleteSettings), mtConfirmation, Array(mbYes, mbNo)) = mbYes)
	
	
	Dim Ini : Set Ini = SDB.IniFile
	
	' Remove settings from ini file
	If (Not Ini Is Nothing) And deleteSettings Then ' delete settings
		Ini.DeleteSection ScriptName
	End If
 
	' Remove entries from scripts/scripts.ini
	SDB.MessageBox SDB.ScriptsPath, mtInformation, Array(mbOK)
	Dim scriptsIni : Set scriptsIni = SDB.Tools.IniFileByPath(SDB.ScriptsPath & "Scripts.ini")
	
	If Not scriptsIni Is Nothing Then
		scriptsIni.DeleteSection(ScriptName)
	End If
	
	' delete AutoPlayer folder
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(SDB.ScriptsPath & ScriptName & "\") Then
		fso.DeleteFolder(SDB.ScriptsPath & ScriptName)
	End If
	
	' Refresh script items to remove control panel and menu item
	SDB.RefreshScriptItems
End Function
