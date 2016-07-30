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
	' Include main script
	Dim fso : set fso = CreateObject("Scripting.FileSystemObject")
	Dim Path : Path = fso.GetParentFolderName(Script.ScriptPath)
	Path = fso.GetParentFolderName(Path) & "\AutoPlayer\APMain.vbs"
	
	Dim f : set f = fso.OpenTextFile(Path, 1)
	Dim code : code = f.ReadAll()
	ExecuteGlobal code
	
	Call OnStartupMain()
End Sub

