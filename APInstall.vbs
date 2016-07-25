'
' AutoPlayer install / uninstall script
'

Const DefaultMinSpacingNew = 10		' Minimum time (days) between repeats of the same song (not skipped yet)
Const DefaultMinSpacing50  = 30		' repeat for 5-star tracks
Const DefaultMinSpacing45  = 45
Const DefaultMinSpacing40  = 60
Const DefaultMinSpacing35  = 75
Const DefaultMinSpacing30  = 90
Const DefaultMinSpacing25  = 105
Const DefaultMinSpacing20  = 150
Const DefaultMinSpacing15  = 200
Const DefaultMinSpacing10  = 250
Const DefaultMinSpacing05  = 325

Const ScriptName = "AutoPlayer"

'
' Installation routine
'
Function BeginInstall
	' base folder to work around MMW bug regarding local vs global folders
	Dim fso : set fso = CreateObject("Scripting.FileSystemObject")
	Dim Path : Path = fso.GetParentFolderName(Script.ScriptPath)
	
	' Add entries to script.ini if you need to show up in the Scripts menu
	Dim inif : Set inif = SDB.Tools.IniFileByPath(Path & "\Scripts.ini")
	
	If Not (inif Is Nothing) Then
		inif.StringValue(ScriptName, "DisplayName") = ScriptName
		inif.IntValue   (ScriptName, "ScriptType")  = 4
		inif.StringValue(ScriptName, "FileName")    = ScriptName & "APMain.vbs"
		inif.StringValue(ScriptName, "Language")    = "VBScript"
	End If

	Dim Ini : Set Ini = SDB.IniFile

	' Set default values; overwrite them if they already exist
	' to allow fresh reinstall
	Ini.IntValue(ScriptName, "MinSpacingNew") = DefaultMinSpacingNew
	Ini.IntValue(ScriptName, "MinSpacing50")  = DefaultMinSpacing50
	Ini.IntValue(ScriptName, "MinSpacing45")  = DefaultMinSpacing45
	Ini.IntValue(ScriptName, "MinSpacing40")  = DefaultMinSpacing40
	Ini.IntValue(ScriptName, "MinSpacing35")  = DefaultMinSpacing35
	Ini.IntValue(ScriptName, "MinSpacing30")  = DefaultMinSpacing30
	Ini.IntValue(ScriptName, "MinSpacing25")  = DefaultMinSpacing25
	Ini.IntValue(ScriptName, "MinSpacing20")  = DefaultMinSpacing20
	Ini.IntValue(ScriptName, "MinSpacing15")  = DefaultMinSpacing15
	Ini.IntValue(ScriptName, "MinSpacing10")  = DefaultMinSpacing10
	Ini.IntValue(ScriptName, "MinSpacing05")  = DefaultMinSpacing05

	If Not fso.FolderExists(Path & "\" & ScriptName & "\") Then
		fso.CreateFolder Path & "\" & ScriptName & "\"
	End If
End Function
