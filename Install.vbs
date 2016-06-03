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
Const IniSection = "AutoPlayer"

'
' Installation routine
'

' Add entries to script.ini if you need to show up in the Scripts menu
Dim inip : inip = SDB.ScriptsPath & "Scripts.ini"
Dim inif : Set inif = SDB.Tools.IniFileByPath(inip)

If Not (inif Is Nothing) Then
	inif.StringValue(ScriptName, "DisplayName") = ScriptName
	inif.IntValue   (ScriptName, "ScriptType")  = 4
	inif.StringValue(ScriptName, "FileName")    = "AutoPlayer.vbs"
	inif.StringValue(ScriptName, "Language")    = "VBScript"
End If

Dim Ini : Set Ini = SDB.IniFile

' Set default values; overwrite them if they already exist
' to allow fresh reinstall
Ini.IntValue(IniSection, "MinSpacingNew") = DefaultMinSpacingNew
Ini.IntValue(IniSection, "MinSpacing50")  = DefaultMinSpacing50
Ini.IntValue(IniSection, "MinSpacing45")  = DefaultMinSpacing45
Ini.IntValue(IniSection, "MinSpacing40")  = DefaultMinSpacing40
Ini.IntValue(IniSection, "MinSpacing35")  = DefaultMinSpacing35
Ini.IntValue(IniSection, "MinSpacing30")  = DefaultMinSpacing30
Ini.IntValue(IniSection, "MinSpacing25")  = DefaultMinSpacing25
Ini.IntValue(IniSection, "MinSpacing20")  = DefaultMinSpacing20
Ini.IntValue(IniSection, "MinSpacing15")  = DefaultMinSpacing15
Ini.IntValue(IniSection, "MinSpacing10")  = DefaultMinSpacing10
Ini.IntValue(IniSection, "MinSpacing05")  = DefaultMinSpacing05

