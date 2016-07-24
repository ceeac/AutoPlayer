
Option Explicit

' Get current time for use in SQL strings
Const CurrTimeUTC = "(JulianDay('now', 'utc')-2415018.5)"


Dim OptsPanel
Dim MenuItem ' Menu item to show / hide panel when clicked


Function GetUTCOffset()
	Dim Iter : Set Iter = SDB.Database.OpenSQL("SELECT (JulianDay('now', 'unixepoch', 'localtime') - JulianDay('now', 'unixepoch', 'utc')) AS UTCOffset")
	
	' round approxOffset to the nerarest 1/2 hour
	' This assumes the SQL query is faster than 15 minutes, which is probably justified.
	Dim UTCOffset : UTCOffset = Round(24 * 2 * Iter.ValueByIndex(0)) / 24 / 2
	Set Iter = Nothing
	GetUTCOffset = UTCOffset
End Function


Sub AppendSkip(Song)
	SDB.Database.ExecSQL("INSERT INTO Skipped(IDSong, SkippedDate, UTCOffset) " &_
		"VALUES (" & Song.ID & ", " & CurrTimeUTC & ", " & GetUTCOffset() & ")")
End Sub


Sub OnStartUp
	' Create Skip Table if it doesn't exist
	SDB.Database.ExecSQL("CREATE TABLE IF NOT EXISTS " &_
		"Skipped(" &_
			"IDSkipped INTEGER PRIMARY KEY AUTOINCREMENT, " &_
			"IDSong INTEGER, " &_
			"SkippedDate REAL, " &_
			"UTCOffset REAL, " &_
			"FOREIGN KEY(IDSong) REFERENCES Songs(ID) ON DELETE CASCADE" &_
		")")
	
	Call Script.RegisterEvent(SDB, "OnTrackSkipped", "AppendSkip")
	
	' Create default options
	If Not SDB.IniFile.IntValue("AutoPlayer", "MinSpacingNew") Then
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacingNew") = 10
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing50")  = 30
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing45")  = 45
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing40")  = 60
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing35")  = 75
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing30")  = 90
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing25")  = 105
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing20")  = 150
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing15")  = 200
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing10")  = 250
		SDB.IniFile.IntValue("AutoPlayer", "MinSpacing45")  = 325
	End If
	
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


