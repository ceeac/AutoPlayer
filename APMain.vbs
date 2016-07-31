'
' AutoPlayer 0.1.0
' AutoDJ script to play higher-rated tracks more often
' Created by eucal
'
'
' APMain.vbs: Main script file.
'
Option Explicit

'
' Constant definitions
'
Const DebugMode = False
Const CurrTime = "(JulianDay('now','localtime')-2415018.5)" ' Get current time for use in SQL strings
Const MaxSpacingTime = 999 ' Maximum value of 'MinSpacing*' values below
Const ScriptName = "AutoPlayer"


'
' Global variable definitions
'
Dim MinSpacingUnr
Dim MinSpacingNew
Dim MinSpacing50
Dim MinSpacing45
Dim MinSpacing40
Dim MinSpacing35
Dim MinSpacing30
Dim MinSpacing25
Dim MinSpacing20
Dim MinSpacing15
Dim MinSpacing10
Dim MinSpacing05
Dim MinSpacing00

Dim APOptionsLoaded : APOptionsLoaded = False

' UI stuff
Dim OptsPanel
Dim MenuItem ' Menu item to show / hide panel when clicked


'
' Start of script here
'
Sub OnStartupMain
	LoadAPOptions
	
	' Create quick options panel
	Set OptsPanel = SDB.UI.NewDockablePersistentPanel("APOptsPanel")
	OptsPanel.Caption = ScriptName & " Quick Options"
	
	If OptsPanel.IsNew Then
		OptsPanel.Common.SetRect 10, 10, 200, 400
		OptsPanel.Common.Visible = True
		OptsPanel.DockedTo = 1 ' Left sidebar
	End If
	
	Script.RegisterEvent OptsPanel, "OnClose", "OptsPanelClose"
	
	' And add the necessary controls
	Dim PlayButton : Set PlayButton = SDB.UI.NewButton(OptsPanel)
	PlayButton.Caption = "Play something!"
	PlayButton.Common.SetRect 10, 10, 125, 25
	PlayButton.Common.Visible = True

	Call Script.RegisterEvent(PlayButton, "OnClick", "ClearAndRefillNowPlaying")
	
	Dim Sep : Set Sep = SDB.UI.AddMenuItemSep(SDB.UI.Menu_View, 0, 0)
	Set MenuItem = SDB.UI.AddMenuItem(SDB.UI.Menu_View, 0, 0)
	MenuItem.Caption = ScriptName & " Quick Options"
	MenuItem.Checked = OptsPanel.Common.Visible
	
	Call Script.RegisterEvent(MenuItem, "OnClick", "OptsPanelShow")
End Sub


Sub OptsPanelShow(Item)
	OptsPanel.Common.Visible = Not OptsPanel.Common.Visible
	MenuItem.Checked = OptsPanel.Common.Visible
End Sub


Sub OptsPanelClose(Item) 
	MenuItem.Checked = False
End Sub 


'
' This function adds a button to the Options Panel
' to edit detailed options regarding AutoPlayer
'
Sub InitConfigSheet(OptionsPanel)
	Dim BtnOptions : Set BtnOptions = SDB.UI.NewButton(OptionsPanel)
	BtnOptions.Common.SetRect 10, 10, 130, 21
	BtnOptions.Caption = "Change configuration"
	Script.RegisterEvent BtnOptions, "OnClick", "ShowDetailedOptions"
End Sub


'
' Saves the configuration when requested.
'
Sub CloseConfigSheet(Panel, SaveConfig)
	Dim OptionsForm : Set OptionsForm = SDB.Objects(ScriptName & "OptsForm")
	If (Not OptionsForm Is Nothing) And SaveConfig Then
		' save spin edit values
		With OptionsForm.Common
			MinSpacingUnr = .ChildControl("Unr").Value
			MinSpacingNew = .ChildControl("New").Value
			MinSpacing50  = .ChildControl("Five").Value
			MinSpacing45  = .ChildControl("FourH").Value
			MinSpacing40  = .ChildControl("Four").Value
			MinSpacing35  = .ChildControl("ThreeH").Value
			MinSpacing30  = .ChildControl("Three").Value
			MinSpacing25  = .ChildControl("TwoH").Value
			MinSpacing20  = .ChildControl("Two").Value
			MinSpacing15  = .ChildControl("OneH").Value
			MinSpacing10  = .ChildControl("One").Value
			MinSpacing05  = .ChildControl("ZeroH").Value
			MinSpacing00  = .ChildControl("Zero").Value
		End With
		
		SaveAPOptions
	End If
	
	Set SDB.Objects(ScriptName & "OptsForm") = Nothing
	Set OptionsForm = Nothing
End Sub


Sub LoadAPOptions()
	If Not APOptionsLoaded Then
		Dim Ini : Set Ini = SDB.IniFile
		
		' Now load ini file values
		MinSpacingUnr = Ini.IntValue(ScriptName, "MinSpacingUnr")
		MinSpacingNew = Ini.IntValue(ScriptName, "MinSpacingNew")
		MinSpacing50  = Ini.IntValue(ScriptName, "MinSpacing50")
		MinSpacing45  = Ini.IntValue(ScriptName, "MinSpacing45")
		MinSpacing40  = Ini.IntValue(ScriptName, "MinSpacing40")
		MinSpacing35  = Ini.IntValue(ScriptName, "MinSpacing35")
		MinSpacing30  = Ini.IntValue(ScriptName, "MinSpacing30")
		MinSpacing25  = Ini.IntValue(ScriptName, "MinSpacing25")
		MinSpacing20  = Ini.IntValue(ScriptName, "MinSpacing20")
		MinSpacing15  = Ini.IntValue(ScriptName, "MinSpacing15")
		MinSpacing10  = Ini.IntValue(ScriptName, "MinSpacing10")
		MinSpacing05  = Ini.IntValue(ScriptName, "MinSpacing05")
		MinSpacing00  = Ini.IntValue(ScriptName, "MinSpacing00")
		
		APOptionsLoaded = True
	End If
End Sub


Sub SaveAPOptions()
	Dim Ini : Set Ini = SDB.IniFile
	
	Ini.IntValue(ScriptName, "MinSpacingUnr") = MinSpacingUnr
	Ini.IntValue(ScriptName, "MinSpacingNew") = MinSpacingNew
	Ini.IntValue(ScriptName, "MinSpacing50")  = MinSpacing50
	Ini.IntValue(ScriptName, "MinSpacing45")  = MinSpacing45
	Ini.IntValue(ScriptName, "MinSpacing40")  = MinSpacing40
	Ini.IntValue(ScriptName, "MinSpacing35")  = MinSpacing35
	Ini.IntValue(ScriptName, "MinSpacing30")  = MinSpacing30
	Ini.IntValue(ScriptName, "MinSpacing25")  = MinSpacing25
	Ini.IntValue(ScriptName, "MinSpacing20")  = MinSpacing20
	Ini.IntValue(ScriptName, "MinSpacing15")  = MinSpacing15
	Ini.IntValue(ScriptName, "MinSpacing10")  = MinSpacing10
	Ini.IntValue(ScriptName, "MinSpacing05")  = MinSpacing05
	Ini.IntValue(ScriptName, "MinSpacing00")  = MinSpacing00
End Sub


'
' Creates a line in the Options Sheet for a specific rating
'
Function CreateSpacingTimeLine(Parent, ByVal xoff, ByVal yoff, LeftLabelText, SpinName)
	Const SpinWidth = 45
	Const LeftLabelWidth = 200
	
	Dim LeftLabel : Set LeftLabel = SDB.UI.NewLabel(Parent)
	LeftLabel.Common.SetRect xoff, yoff+4, LeftLabelWidth, 25
	LeftLabel.Caption = SDB.Localize(LeftLabelText)
	xoff = xoff + LeftLabelWidth + 5
	
	Dim SpacingTimeEdit : Set SpacingTimeEdit = SDB.UI.NewSpinEdit(Parent)
	SpacingTimeEdit.Common.SetRect xoff, yoff, SpinWidth, 25
	SpacingTimeEdit.Common.ControlName = SpinName
	SpacingTimeEdit.MinValue = 0
	SpacingTimeEdit.MaxValue = MaxSpacingTime
	xoff = xoff + SpinWidth + 5
	
	Dim RightLabel : Set RightLabel = SDB.UI.NewLabel(Parent)
	RightLabel.Common.SetRect xoff, yoff+4, 50, 25
	RightLabel.Caption = SDB.Localize("days")
	
	Set CreateSpacingTimeLine = SpacingTimeEdit
End Function


'
' This function initializes the Options Widow for AutoPlayer.
'
Sub ShowDetailedOptions()
	Dim OptionsForm : Set OptionsForm = SDB.Objects(ScriptName & "OptsForm")
	If OptionsForm Is Nothing Then
		' Panel was not already visible before, create it
		LoadAPOptions
		
		' Show config panel
		Set OptionsForm = SDB.UI.NewForm
		OptionsForm.Common.SetRect 100, 100, 460, 375
		OptionsForm.BorderStyle  = 3
		OptionsForm.FormPosition = 4
		OptionsForm.Caption = ScriptName & " Settings"

		Const DeltaX = 0
		Const DeltaY = 25
		
		Dim X : X = 10
		Dim Y : Y = 10

		Dim MinSpacingUnrEdit : Set MinSpacingUnrEdit = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for unrated tracks:", "Unr")     : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacingNewEdit : Set MinSpacingNewEdit = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for unskipped tracks:", "New")   : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing50Edit  : Set MinSpacing50Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 5.0-star tracks:", "Five")   : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing45Edit  : Set MinSpacing45Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 4.5-star tracks:", "FourH")  : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing40Edit  : Set MinSpacing40Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 4.0-star tracks:", "Four")   : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing35Edit  : Set MinSpacing35Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 3.5-star tracks:", "ThreeH") : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing30Edit  : Set MinSpacing30Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 3.0-star tracks:", "Three")  : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing25Edit  : Set MinSpacing25Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 2.5-star tracks:", "TwoH")   : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing20Edit  : Set MinSpacing20Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 2.0-star tracks:", "Two")    : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing15Edit  : Set MinSpacing15Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 1.5-star tracks:", "OneH")   : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing10Edit  : Set MinSpacing10Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 1.0-star tracks:", "One")    : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing05Edit  : Set MinSpacing05Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for 0.5-star tracks:", "ZeroH")  : X = X + DeltaX : Y = Y + DeltaY
		Dim MinSpacing00Edit  : Set MinSpacing00Edit  = CreateSpacingTimeLine(OptionsForm, X, Y, "Min spacing for bomb tracks:", "Zero")       : X = X + DeltaX : Y = Y + DeltaY
		
		MinSpacingUnrEdit.Value = MinSpacingUnr
		MinSpacingNewEdit.Value = MinSpacingNew
		MinSpacing50Edit.Value  = MinSpacing50
		MinSpacing45Edit.Value  = MinSpacing45
		MinSpacing40Edit.Value  = MinSpacing40
		MinSpacing35Edit.Value  = MinSpacing35
		MinSpacing30Edit.Value  = MinSpacing30
		MinSpacing25Edit.Value  = MinSpacing25
		MinSpacing20Edit.Value  = MinSpacing20
		MinSpacing15Edit.Value  = MinSpacing15
		MinSpacing10Edit.Value  = MinSpacing10
		MinSpacing05Edit.Value  = MinSpacing05
		MinSpacing00Edit.Value  = MinSpacing00
		
		' Add OK button
		Dim OKButton : Set OKButton = SDB.UI.NewButton(OptionsForm)
		OKButton.Common.SetRect 300, 300, 130, 21
		OKButton.Caption = "&OK"
		OKButton.ModalResult = 1

		' Finally show the configuration dialogue
		Set SDB.Objects(ScriptName & "OptsForm") = OptionsForm
	End If
	
	OptionsForm.ShowModal
End Sub


' Writes a debug message if debug mode is enabled.
' Does nothing otherwise.
Sub DbgMsg(str)
	If DebugMode Then
		' Force debug output
		SDB.Tools.OutputDebugString("AP: " & str)
	Else
		' output only when using the debug version of MM
		SDB.Tools.OutputDebugStringMM("AP: " & str)
	End If
End Sub


Function FixSearchString(Str)
	Str = Replace(Str, "'", "''") '<--- Single quotes are escaped with another single quote
	FixSearchString = Str
End Function

	
'
' Checks if a track can be queued for Now Playing
' This is true if all of the following conditions hold:
'  - The track is not yet in the Now Playing list
'  - The track is in the database
'
Function IsTrackOK(Song)
	IsTrackOK = False
	
	' Sanity Check
	If Song.IsntInDB Then
		DbgMsg("Rejecting " & Song.ArtistName & " - " & Song.Title & ": Track is not in library.")
		Exit Function
	End If

	' Do not play a track from an album/artist if it's in the now playing list
	Dim i
	Dim NowPlayingSong

	For i = 0 To SDB.Player.CurrentSongList.Count-1
		Set NowPlayingSong = SDB.Player.CurrentSongList.Item(i)
		If NowPlayingSong.AlbumName = Song.AlbumName Or NowPlayingSong.AlbumArtistName = Song.AlbumArtistName Or NowPlayingSong.Title = Song.Title Then
			DbgMsg("Rejecting " & Song.ArtistName & " - " & Song.Title & ": Track is already in NowPlaying list")
			Exit Function
		End If
	Next
	
	' Check if file exists
	If Not SDB.Tools.FileSystem.FileExists(Song.Path) Then
		DbgMsg("Rejecting " & Iter.Item.ArtistName & " - " & Iter.Item.Title & ": File does not exist")
		Exit Function
	End If
	
	IsTrackOK = True
End Function


Function GetSpacingQuery(ByVal MinSpacingFactor)
	GetSpacingQuery = "(" &_
		"(SkipCount = 0 AND " & CurrTime & "-LastTimePlayed > " & MinSpacingNew & ") OR " &_
		"(SkipCount > 0 AND "            & "Rating  = -1 AND " & CurrTime & "-LastTimePlayed > " & MinSpacingUnr * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating >= 0 AND Rating <=  5 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing00  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating >  5 AND Rating <= 15 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing05  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 15 AND Rating <= 25 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing10  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 25 AND Rating <= 35 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing15  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 35 AND Rating <= 45 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing20  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 45 AND Rating <= 55 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing25  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 55 AND Rating <= 65 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing30  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 65 AND Rating <= 75 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing35  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 75 AND Rating <= 85 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing40  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 85 AND Rating <= 95 AND " & CurrTime & "-LastTimePlayed > " & MinSpacing45  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 95 AND "                  & CurrTime & "-LastTimePlayed > " & MinSpacing50  * MinSpacingFactor & ") )"
End Function


' Generates a new track to be queued for Now Playing
Function GenerateNewTrack

	LoadAPOptions
	
	' Select only tracks that have not been played for some time
	Dim QueryString : QueryString = "Custom3 NOT LIKE '%Archive%' AND PlayCounter > 0 AND " &_
		GetSpacingQuery(1) & " ORDER BY RANDOM(*)"
		
	' Clear message queue first
	SDB.ProcessMessages
	
	' Now query the SQL DB
	Dim Iter : Set Iter = SDB.Database.QuerySongs(QueryString)
	
	Do Until Iter.EOF
		' Check tracks if they can be inserted into the Now Playing list
		DbgMsg("Considering '" & Iter.Item.ArtistName & " - " & Iter.Item.Title & "'")
		
		If IsTrackOK(Iter.Item) Then
			DbgMsg("NowPlayingAdd '" & Iter.Item.ArtistName & " - " & Iter.Item.Title & "'")
			
			Set GenerateNewTrack = Iter.Item
			Set Iter = Nothing
			Exit Function
		End If
		
		Iter.Next
	Loop
	
	' Clean up
	Set Iter = Nothing
	SDB.ProcessMessages
	
	
	DbgMsg("Panic: Selecting random track")
	Set Iter = SDB.Database.QuerySongs("ORDER BY RANDOM(*) LIMIT 1")
	If Iter.EOF Then
		' There is nothing we can do about it; there are probably no tracks in the library
		DbgMsg("Giving up: No suitable track has been found")
		
		Set Iter = Nothing
		Set GenerateNewTrack = Nothing
		Exit Function
	End If
	
	' All OK -> Tell about now playing song
	DbgMsg("NowPlayingAdd " & Iter.Item.ArtistName & " - " & Iter.Item.Title)
	
	Set Iter = Nothing
End Function


'
' Clear now playing list and refill.
' Called when the 'Play something!' button is clicked.
'
' Does the following:
' 1. Stop playback
' 2. Clear Now Playing list
' 3. Query new track via GenerateNewTrack
' 4. Enable AutoDJ, disable shuffle
' 5. Play track
'
Sub ClearAndRefillNowPlaying
	SDB.Player.Stop
	SDB.Player.PlaylistClear
	
	' Get first track
	Dim NewSong : Set NewSong = GenerateNewTrack()
	SDB.Player.PlaylistAddTrack NewSong
	Set NewSong = Nothing
	
	SDB.Player.IsAutoDJ  = True
	SDB.Player.isShuffle = False
	
	' Clear message queue before starting playback (just to be sure)
	SDB.ProcessMessages
	SDB.Player.Play
End Sub

