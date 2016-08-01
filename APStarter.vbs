'
' AutoPlayer 0.1.0
' AutoDJ script to play higher-rated tracks more often
' Created by eucal
'
'
' APStarter.vbs: Helper file to start the main script on startup.
'
Option Explicit


Sub OnStartUp
	' Include main script and call it
	Dim ini : Set ini = SDB.IniFile
	
	If Not ini Is Nothing Then
		Script.Include ini.StringValue("AutoPlayer", "RootPath") & "APMain.vbs"
		Call OnStartupMain()
	End If
End Sub

