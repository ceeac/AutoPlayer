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
' Class dealing with loading, storing and saving settings.
'
Class APSettings
	' settings version. Used for backwards compatibility.
	Private m_settingsVersion
	
	' Values for min spacing of tracks.
	' Index 0: Unrated
	' Index 1: Unskipped
	' Index 2: 5-star tracks
	' Index 3: 4.5-star tracks
	' etc.
	Private m_minSpacing(12)
	
	' Dictionary of known mood tags and if they are allowed.
	Private m_allowedMoods
	
	'
	' Properties
	'
	
	' Get ini file for saving and loading.
	Public Property Get IniFile
		Set IniFile = SDB.Tools.IniFileByPath(SDB.IniFile.StringValue(ScriptName, "RootPath") & ScriptName & ".ini")
	End Property
	
	Public Property Get MinSpacing(ByVal rating)
		MinSpacing = m_minSpacing(ratingToIndex(rating))
	End Property
	
	Public Property Let MinSpacing(ByVal rating, ByVal spacing)
		m_minSpacing(ratingToIndex(rating)) = spacing
	End Property
	
	Public Property Get AllowedMoods
		Set AllowedMoods = m_allowedMoods
	End Property
	
	Public Property Get MoodAllowed(ByVal mood)
		If Not m_AllowedMoods.Exists(mood) Then
			MoodAllowed = False
		Else
			MoodAllowed = m_allowedMoods(mood)
		End If
	End Property
	

	' "Constructor"
	Private Sub Class_Initialize
		m_SettingsVersion = 1
				
		' set default spacing values
		m_MinSpacing(ratingToIndex(100)) = 30  ' For 5-star tracks
		m_MinSpacing(ratingToIndex(90))  = 45
		m_MinSpacing(ratingToIndex(80))  = 60
		m_MinSpacing(ratingToIndex(70))  = 75
		m_MinSpacing(ratingToIndex(60))  = 90
		m_MinSpacing(ratingToIndex(50))  = 105
		m_MinSpacing(ratingToIndex(40))  = 150
		m_MinSpacing(ratingToIndex(30))  = 200
		m_MinSpacing(ratingToIndex(20))  = 250
		m_MinSpacing(ratingToIndex(10))  = 325
		m_MinSpacing(ratingToIndex(0))   = 365 ' Bomb rating
		m_MinSpacing(ratingToIndex(-1))  = 105 ' unknown rating
		m_MinSpacing(ratingToIndex(-2))  = 10  ' unskipped songs (SkipCount = 0)
		
		Set m_AllowedMoods = CreateObject("Scripting.Dictionary")
	End Sub
	
	' "Destructor"
	Private Sub Class_Terminate
		Set m_AllowedMoods = Nothing
	End Sub
	
	' Load settings from ini file.
	Public Sub loadFromFile
		Dim Ini : Set Ini = IniFile
		DbgMsg("Loading settings from " & Ini.Path)
		
		' Load version information TODO
		Dim saveVersion : saveVersion = IniFile.IntValue("VersionInfo", "SaveVersion")
		
		' Now load ini file values
		Dim i
		For i=0 To UBound(m_minSpacing)
			If Ini.ValueExists("Spacing", "MinSpacing" & i) Then
				m_minSpacing(i) = Ini.IntValue("Spacing", "MinSpacing" & i)
			End If
		Next
		
		m_allowedMoods.RemoveAll
				
		' Get all mood tags from the database and update allowed mood
		Dim MoodIter : Set MoodIter = SDB.Database.OpenSQL("SELECT Mood FROM Songs GROUP BY Mood")
		Do While Not MoodIter.EOF
			Dim mood : mood = MoodIter.ValueByIndex(0)
			If mood = "" Then mood = "<Unknown>"
			
			If Not Ini.ValueExists("AllowedMoods", mood) Then
				m_allowedMoods(mood) = True ' allow unknown/new moods to be played by default
			Else
				m_allowedMoods(mood) = Ini.BoolValue("AllowedMoods", mood)
			End If
			
			MoodIter.Next
		Loop
	End Sub
	
	' Save settings to ini file
	Public Sub saveToFile
		Dim Ini : Set Ini = IniFile
		DbgMsg("Saving settings to " & ini.Path)

		Ini.IntValue("VersionInfo", "SaveVersion") = m_SettingsVersion
		
		Dim i
		For i=0 To 12
			Ini.IntValue("Spacing", "MinSpacing" & i) = m_MinSpacing(i)
		Next
		
		Dim mood
		For Each mood In m_allowedMoods.Keys
			Ini.BoolValue("AllowedMoods", mood) = m_allowedMoods(mood)
		Next
	End Sub
	
	
	' Converts the rating of a song to an index used for minSpacing array.
	Private Function ratingToIndex(ByVal rating)
		If rating = -1 Then
			ratingToIndex = 12
		ElseIf rating = -2 Then
			ratingToIndex = 11
		Else
			ratingToIndex = Round((rating-1) / 10, 0)
		End If
	End Function
End Class



' UI stuff
Dim ControlPanel
Dim ShowPanelMenuItem ' Menu item to show / hide panel when clicked



' Creates mood checkboxes on the control panel.
Sub CreateMoodCheckboxes(ByRef X, ByRef Y)
	Dim allowedMoods : Set allowedMoods = SDB.Objects("APSettings").AllowedMoods
	
	' Now create check boxes
	Dim mood
	For Each mood In allowedMoods.Keys
		Dim ChkBox : Set ChkBox = SDB.UI.NewCheckBox(ControlPanel)
		With ChkBox
			.Checked = allowedMoods(mood)
			.Common.Visible = True
			.Caption = mood
			.Common.SetRect X, Y, 125, 20
		End With
		
		Script.UnRegisterEvents ChkBox
		Script.RegisterEvent ChkBox.Common, "OnClick", "OnCheckBoxToggled"
		
		Y = Y + 20
	Next
End Sub


'
' Main procedure of the script.
' Called on startup by AutoPlayerStarter.vbs and initializes all variables etc.
'
Sub OnStartupMain
	DbgMsg ScriptName & " starting..."
	
	Dim settings : Set settings = New APSettings
	Set SDB.Objects("APSettings") = settings
	settings.LoadFromFile
	
	'
	' UI stuff
	' 
	' Create quick options panel
	Set ControlPanel = SDB.UI.NewDockablePersistentPanel("APControlPanel")
	ControlPanel.Caption = ScriptName & " Control Panel"
	
	If ControlPanel.IsNew Then
		ControlPanel.Common.SetRect 10, 10, 200, 400
		ControlPanel.Common.Visible = True
		ControlPanel.DockedTo = 1 ' Left sidebar
	End If
	
	Script.RegisterEvent ControlPanel, "OnClose", "ControlPanelClose"
	
	' And add the necessary controls
	Dim X : X = 10
	Dim Y : Y = 10
	
	Dim PlayButton : Set PlayButton = SDB.UI.NewButton(ControlPanel)
	With PlayButton
		.Caption = SDB.Localize("Play something!")
		.Common.SetRect X, Y, 125, 25
		.Common.Visible = True
	End With

	Y = Y + 35
	
	Script.RegisterEvent PlayButton, "OnClick", "ClearAndRefillNowPlaying"
	
	' Add label "Allowed Moods:"
	Dim AllowedMoodsLabel : Set AllowedMoodsLabel = SDB.UI.NewLabel(ControlPanel)
	With AllowedMoodsLabel
		.Alignment = 0 ' Left
		.Caption = SDB.Localize("Allowed Moods:")
		.Multiline = False
		.Autosize = True
		.Common.SetRect X, Y, 125, 25
	End With
	Y = Y + 15
	
	CreateMoodCheckboxes X, Y
	
	' Add menu item to show / hide the control panel
	Dim Sep : Set Sep = SDB.UI.AddMenuItemSep(SDB.UI.Menu_View, 0, 0)
	Set ShowPanelMenuItem = SDB.UI.AddMenuItem(SDB.UI.Menu_View, 0, 0)
	ShowPanelMenuItem.Caption = ControlPanel.Caption
	ShowPanelMenuItem.Checked = ControlPanel.Common.Visible
	
	Script.RegisterEvent ShowPanelMenuItem, "OnClick", "ControlPanelShow"
	Script.RegisterEvent SDB, "OnShutdown", "HandleShutdown"
	
	DbgMsg(ScriptName & " started.")
End Sub


' Update allowed moods when the corresponding check box is clicked
Sub OnCheckBoxToggled(chkBox)
	' Update allowed moods from settings
	SDB.Objects("APSettings").AllowedMoods.Item(chkBox.Caption) = chkBox.Checked
End Sub


Sub ControlPanelShow(Item)
	ControlPanel.Common.Visible = Not ControlPanel.Common.Visible
	ShowPanelMenuItem.Checked   = ControlPanel.Common.Visible
End Sub


Sub ControlPanelClose(Item)
	ShowPanelMenuItem.Checked = False
End Sub


' Creates a button that shows detailed settings when clicked.
Sub InitConfigSheet(OptionsPanel)
	Dim BtnOptions : Set BtnOptions = SDB.UI.NewButton(OptionsPanel)
	BtnOptions.Common.SetRect 10, 10, 130, 21
	BtnOptions.Caption = "Change configuration"
	Script.RegisterEvent BtnOptions, "OnClick", "ShowDetailedOptions"
End Sub


' Closes the settings dialogue and optionally saves it.
Sub CloseConfigSheet(Panel, SaveConfig)
	Dim OptionsForm : Set OptionsForm = SDB.Objects("APOptsForm")
	If (Not OptionsForm Is Nothing) And SaveConfig Then
		Dim settings : Set settings = SDB.Objects("APSettings")
		
		' save spin edit values
		With OptionsForm.Common
			settings.MinSpacing(-2)  = .ChildControl("New").Value
			settings.MinSpacing(-1)  = .ChildControl("Unr").Value
			settings.MinSpacing(100) = .ChildControl("Five").Value
			settings.MinSpacing(90)  = .ChildControl("FourH").Value
			settings.MinSpacing(80)  = .ChildControl("Four").Value
			settings.MinSpacing(70)  = .ChildControl("ThreeH").Value
			settings.MinSpacing(60)  = .ChildControl("Three").Value
			settings.MinSpacing(50)  = .ChildControl("TwoH").Value
			settings.MinSpacing(40)  = .ChildControl("Two").Value
			settings.MinSpacing(30)  = .ChildControl("OneH").Value
			settings.MinSpacing(20)  = .ChildControl("One").Value
			settings.MinSpacing(10)  = .ChildControl("ZeroH").Value
			settings.MinSpacing(0)   = .ChildControl("Zero").Value
		End With
		
		settings.saveToFile
	End If
	
	Set SDB.Objects("APOptsForm") = Nothing
	Set OptionsForm = Nothing
End Sub


' Saves configuration on shutdown.
Sub HandleShutdown()
	DbgMsg "Shutting down ..."
	SDB.Objects("APSettings").SaveToFile
	Set SDB.Objects("APSettings") = Nothing
	DbgMsg "Shutdown finished."
End Sub


' Creates a line in the Options Sheet for a specific rating
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


' This function initializes the Options Widow for AutoPlayer.
Sub ShowDetailedOptions()
	Dim OptionsForm : Set OptionsForm = SDB.Objects("APOptsForm")
	If OptionsForm Is Nothing Then
	
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
		
		With SDB.Objects("APSettings")
			MinSpacingNewEdit.Value = .MinSpacing(-2)
			MinSpacingUnrEdit.Value = .MinSpacing(-1)
			MinSpacing50Edit.Value  = .MinSpacing(100)
			MinSpacing45Edit.Value  = .MinSpacing(90)
			MinSpacing40Edit.Value  = .MinSpacing(80)
			MinSpacing35Edit.Value  = .MinSpacing(70)
			MinSpacing30Edit.Value  = .MinSpacing(60)
			MinSpacing25Edit.Value  = .MinSpacing(50)
			MinSpacing20Edit.Value  = .MinSpacing(40)
			MinSpacing15Edit.Value  = .MinSpacing(30)
			MinSpacing10Edit.Value  = .MinSpacing(20)
			MinSpacing05Edit.Value  = .MinSpacing(10)
			MinSpacing00Edit.Value  = .MinSpacing(0)
		End With
		
		' Add OK button
		Dim OKButton : Set OKButton = SDB.UI.NewButton(OptionsForm)
		OKButton.Common.SetRect 300, 300, 130, 21
		OKButton.Caption = "&OK"
		OKButton.ModalResult = 1

		' Finally show the configuration dialogue
		Set SDB.Objects("APOptsForm") = OptionsForm
	End If
	
	OptionsForm.ShowModal
End Sub

'
' Writes a debug message if debug mode is enabled
' or when using the debug version of MediaMonkey.
' Does nothing otherwise.
'
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
	Str = Replace(Str, "'", "''")
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
	Dim settings : Set settings = SDB.Objects("APSettings")
	
	GetSpacingQuery = "(" &_
		"(SkipCount = 0 AND "                                  & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(-2)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND "            & "Rating  = -1 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(-1)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating >= 0 AND Rating <=  5 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(0)   * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating >  5 AND Rating <= 15 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(10)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 15 AND Rating <= 25 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(20)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 25 AND Rating <= 35 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(30)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 35 AND Rating <= 45 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(40)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 45 AND Rating <= 55 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(50)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 55 AND Rating <= 65 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(60)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 65 AND Rating <= 75 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(70)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 75 AND Rating <= 85 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(80)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 85 AND Rating <= 95 AND " & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(90)  * MinSpacingFactor & ") OR " &_
		"(SkipCount > 0 AND Rating > 95 AND "                  & CurrTime & "-LastTimePlayed > " & settings.MinSpacing(100) * MinSpacingFactor & ") )"
End Function


'
' Return query string determined by which moods are allowed.
' this limits the selection of songs to those with "allowed" mood tags.
'
Function GetAllowedMoodsString
	' Only select songs where the checkboxof the mood tag is checked
	Dim allowedMoods : Set allowedMoods = SDB.Objects("APSettings").AllowedMoods
	Dim QueryMoodString : QueryMoodString = "(0 " ' 0=False
	Dim mood
	
	For Each mood In allowedMoods.Keys
		If allowedMoods.Item(Mood) Then
			If mood = "<Unknown>" Then
				QueryMoodString = QueryMoodString & "OR Mood=''"
			Else
				QueryMoodString = QueryMoodString & "OR Mood='" & mood & "' "
			End If
		End If
	Next
	QueryMoodString = QueryMoodString & ")"
	GetAllowedMoodsString = QueryMoodString
End Function


'
' Generates a new track to be queued for Now Playing
'
Function GenerateNewTrack
	' Select only tracks that have not been played for some time
	Dim QueryString : QueryString = "Custom3 NOT LIKE '%Archive%' AND PlayCounter > 0 AND " &_
		GetSpacingQuery(1) & " AND " & GetAllowedMoodsString() & " ORDER BY RANDOM(*)"

	' Clear message queue first
	SDB.ProcessMessages
	
	' Now query the SQL DB
	Dim Iter : Set Iter = SDB.Database.QuerySongs(QueryString)
	
	Do Until Iter.EOF
		' Check tracks if they can be inserted into the Now Playing list
		DbgMsg("Considering '" & Iter.Item.ArtistName & " - " & Iter.Item.Title & "'")
		
		If IsTrackOK(Iter.Item) Then
			DbgMsg("NowPlayingAdd '" & Iter.Item.ArtistName & " - " & Iter.Item.Title & "'")
			DbgMsg("")
			
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
	Set Iter = SDB.Database.QuerySongs("1 ORDER BY RANDOM(*) LIMIT 1")
	If Iter.EOF Then
		' There is nothing we can do about it; there are probably no tracks in the library
		DbgMsg("Giving up: No suitable track has been found")
		
		Set Iter = Nothing
		Set GenerateNewTrack = Nothing
		Exit Function
	End If
	
	' All OK -> Tell about now playing song
	DbgMsg("NowPlayingAdd " & Iter.Item.ArtistName & " - " & Iter.Item.Title)
	DbgMsg("")
	
	Set GenerateNewTrack = Iter.Item
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


Sub BeginUninstall
	Dim ini : Set ini = SDB.IniFile
	Dim rootPath : rootPath = ini.StringValue(ScriptName, "RootPath")
	SDB.Tools.FileSystem.CopyFile rootPath & "APInstaller.vbs", SDB.ScriptsPath & "APInstaller.vbs"
End Sub
