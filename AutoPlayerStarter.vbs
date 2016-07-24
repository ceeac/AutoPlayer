'
' AutoPlayer 0.1.0
' AutoDJ script to play higher-rated tracks more often
' Created by eucal
'
Option Explicit

'
' Helper function to call the main AutoDJ script on startup.
'
Sub OnStartUp
	' Include main script so we can assign the callback function when the play something button is pressed
	
	Dim fso : set fso = CreateObject("Scripting.FileSystemObject")
	Dim Path : Path = fso.GetParentFolderName(Script.ScriptPath)
	Path = fso.GetParentFolderName(Path) & "\\AutoPlayer.vbs"
	
	Dim f : set f = fso.OpenTextFile(Path, 1)
	Dim code : code = f.ReadAll()
	ExecuteGlobal code
	
	Call OnStartupMain()
End Sub

