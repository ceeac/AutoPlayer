
Option Explicit

' Get current time for use in SQL strings
Const CurrTimeUTC = "(JulianDay('now', 'utc')-2415018.5)"


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
End Sub


