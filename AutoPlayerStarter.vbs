
Option Explicit

' Get current time for use in SQL strings
Const CurrTimeUTC = "(JulianDay('now', 'utc')-2415018.5)"


Dim OptsPanel
Dim MenuItem ' Menu item to show / hide panel when clicked


Sub OnStartUp
	' Create quick options panel
	Set OptsPanel = SDB.UI.NewDockablePersistentPanel("APOptsPanel")
	OptsPanel.Common.SetRect 10, 10, 200, 400
	OptsPanel.Common.Visible = True
	OptsPanel.Caption = "AutoPlayer Quick Options"
	OptsPanel.DockedTo = 1 ' Left sidebar

	Script.RegisterEvent OptsPanel, "OnClose", "OptsPanelClose"
	
	' And add the necessary controls
	Dim PlayButton : Set PlayButton = SDB.UI.NewButton(OptsPanel)
	PlayButton.Caption = "Play something!"
	PlayButton.Common.SetRect 10, 10, 125, 25
	PlayButton.Common.Visible = True

	' Include main script so we can assign the callback function when the play something button is pressed
	Dim fso : set fso = CreateObject("Scripting.FileSystemObject")
	Dim Path : Path = fso.GetParentFolderName(Script.ScriptPath)
	Path = fso.GetParentFolderName(Path) & "\\AutoPlayer.vbs"
	
	Dim f : set f = fso.OpenTextFile(Path, 1)
	Dim s : s = f.ReadAll()
	ExecuteGlobal s
	Call OnStartupMain()
	
	Call Script.RegisterEvent(PlayButton, "OnClick", "ClearAndRefillNowPlaying")
	
	Dim Sep : Set Sep = SDB.UI.AddMenuItemSep(SDB.UI.Menu_View, 0, 0)
	Set MenuItem = SDB.UI.AddMenuItem(SDB.UI.Menu_View, 0, 0)
	MenuItem.Caption = "AutoPlayer Quick Options"
	
	Call Script.RegisterEvent(MenuItem, "OnClick", "OptsPanelShow")
End Sub


Sub OptsPanelShow(Item)
	OptsPanel.Common.Visible = Not OptsPanel.Common.Visible
	MenuItem.Checked = OptsPanel.Common.Visible
End Sub


Sub OptsPanelClose(Item) 
	MenuItem.Checked = False
End Sub 


