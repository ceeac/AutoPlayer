
Option Explicit

Sub OnStartUp
	' Include main script so we can assign the callback function when the play something button is pressed
	Dim fso : set fso = CreateObject("Scripting.FileSystemObject")
	Dim Path : Path = fso.GetParentFolderName(Script.ScriptPath)
	Path = fso.GetParentFolderName(Path) & "\\AutoPlayer.vbs"
	
	Dim f : set f = fso.OpenTextFile(Path, 1)
	Dim s : s = f.ReadAll()
	ExecuteGlobal s
	Call OnStartupMain()
End Sub




