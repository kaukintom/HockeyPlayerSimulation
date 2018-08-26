Option Explicit

Const year1 As Integer = 2017 'Change to Latest season start date

'***IMPORTANT***Before running macro, make sure to add the games, teams, and season to the _
'Adjustment worksheet

Sub All_Stats()
'Consolidates data across all seasons into one worksheet
'Creates three separate instances of the new consolidated data _
', which all represent a different way to look at the data
'Adjust past data to make it relevent to compare across seasons
Dim Nsheet As Worksheet 'NAMES worksheet
Dim names As Variant 'Array of NAMES worksheet
Dim HockeyStatsSheet As Worksheet 'Sheet for consolidated data
Dim Nrow As Integer 'Row number for names
Dim Name As String 'Name of hockey player
Dim Season As Variant 'Array of data for a specific season
Dim Srow As Integer 'Row number for Season
Dim Data As Variant 'Row of data for hockey player in a season
Dim Earliest As Integer 'Season  of data that is furthest in the past
Dim Year As Integer 'Season between Earliest and year1 --> for a For Loop
'Consolidated data arrays
Dim Asheet As Variant, Bsheet As Variant, Psheet As Variant, Ysheet As Variant
Dim StartTime As Double 'Variable for timing program
Dim Adjust As Variant 'Takes the goal and assist multipliers to adjusts the stats of _
'previous seasons
Dim HDataO As Workbook 'For opening another workbook
Dim SecondsElapsed As Double 'Variable for timing program
Dim Dcol As Integer 'Element number for Data array
Dim HSRow As Integer
Dim NameCol As Integer

StartTime = Timer

Application.ScreenUpdating = False
Set HockeyStatsSheet = Sheets("Hockey Stats")
Set Nsheet = Sheets("NAMES")
HockeyStatsSheet.UsedRange.delete
names = Nsheet.UsedRange
HockeyStatsSheet.Select
Earliest = 1979 'Earliest season of data collected
'*********************************************************
'Sets up Hockey Stats worksheet to have all the proper headings at the top _
'of the page
Season = Sheets(year1 & "|" & (year1 + 1)).UsedRange

ReDim Data(1 To UBound(Season, 2))

Sheets(year1 & "|" & (year1 + 1)).Select

Data = Range(Cells(1, 1), Cells(1, UBound(Season, 2)))
HockeyStatsSheet.Select
Range(Cells(1, 2), Cells(1, 1 + UBound(Season, 2))) = Data
Cells(1, 1).Value = "Year"

ReDim Data(1 To UBound(Season, 2))

HSRow = 1
NameCol = 2
'**********************************************************

'loops through all the names on the NAMES worksheet, which contains _
'all the players in the workbook
'for each player, their stats from every season they played are put _
'on a single worksheet
For Nrow = LBound(names) To UBound(names, 1)
    Name = names(Nrow, 1)
    'Loop through the worksheets
    For Year = Earliest To year1
    '2004 was a lockout season, so it doesnt exist
        If Year <> 2004 Then
            'put the latest season into an array and search through the array
            Season = Sheets(Year & "|" & (Year + 1)).UsedRange
            For Srow = LBound(Season) To UBound(Season, 1)
                'when the name is found, go to that sheet and put their stats _
                'into and array(data)
                If Season(Srow, 1) = Name Then
                    'Sheets(year & "|" & (year + 1)).Select
                    HSRow = HSRow + 1
                    For Dcol = LBound(Data) To UBound(Data)
                        HockeyStatsSheet.Cells(HSRow, Dcol + NameCol - 1).Value = Season(Srow, Dcol)
                    Next Dcol
                    'stats are put onto worksheet along with the season they played in
                    HockeyStatsSheet.Cells(HSRow, 1).Value = Year & "|" & (Year + 1)
                    'Stop
                    Exit For
                End If
            Next Srow
        End If
    Next Year
Next Nrow

Call CorrectSeasons 'Calculates goals, assists, and points for specified NHL seasons
Call Stat_Game_Team 'Determine Goals, Assists, and Points per game per team for each season
Call Adjust_Needed 'Determine multiplier for goals, asists, and points to match to current season

Call Adjust_All_Stats 'Use the calculated multipliers and apply them to consolidated Hockey Stats worksheet
Call differences_stats 'Finds yearly change in stats in three different formats for hockey players

Application.ScreenUpdating = True
MsgBox (Round(Timer - StartTime, 2) / 60 & " Minutes.")
End Sub

Sub differences_stats()
'Finds yearly change in stats in three different formats for hockey players
Dim names As Variant, Nsheet As Worksheet, HockeyStatsSheet As Worksheet
Dim X As Integer, i As Integer, shift As Integer, HockeyStats As Variant
Dim Name As String, Nrow As Integer, games1 As Integer, games2 As Integer
Dim Prow As Integer, adiff As Worksheet, pdiff As Worksheet, Per As Variant
Dim Bdiff As Worksheet, Actual As Double, Binomial As String, Percent As Double

Set HockeyStatsSheet = Sheets("Hockey Stats")
Set Nsheet = Sheets("NAMES")
Set adiff = Sheets("Actual Differences")
Set pdiff = Sheets("Percent Differences")
Set Bdiff = Sheets("Binomial Differences")
HockeyStats = HockeyStatsSheet.UsedRange
names = Nsheet.UsedRange

adiff.UsedRange.delete
Bdiff.UsedRange.delete
pdiff.UsedRange.delete

adiff.Select
Range(Cells(1, 1), Cells(UBound(HockeyStats, 1), UBound(HockeyStats, 2))) = HockeyStats
Bdiff.Select
Range(Cells(1, 1), Cells(UBound(HockeyStats, 1), UBound(HockeyStats, 2))) = HockeyStats
pdiff.Select
Range(Cells(1, 1), Cells(UBound(HockeyStats, 1), UBound(HockeyStats, 2))) = HockeyStats

HockeyStatsSheet.Select
Cells(1, 1).Select

Per = adiff.UsedRange

For Nrow = LBound(names) To UBound(names, 1)
    Name = names(Nrow, 1)
    For Prow = LBound(Per) To UBound(Per, 1)
        If Prow = UBound(Per, 1) Then
            Exit For
        ElseIf Per(Prow, 2) = Name And Per(Prow + 1, 2) <> Name Then
            Exit For
        End If
    Next Prow
    Cells(Prow, 2).Select
    X = 0
    Do Until Cells(Prow + X, 2).Value <> Name
        If Cells(Prow + X - 1, 2).Value = Name Then
            games1 = Cells(Prow + X, 4).Value
            games2 = Cells(Prow + X - 1, 4).Value
            If games1 = 0 Or games2 = 0 Then
                games1 = 1
                games2 = 1
            End If
            For i = 5 To 21
                If IsNumeric(Cells(Prow + X, i).Value) = True And IsNumeric(Cells(Prow + X - 1, i).Value) = True Then
                    Actual = Cells(Prow + X, i).Value / games1 - Cells(Prow + X - 1, i).Value / games2
                    adiff.Cells(Prow + X, i).Value = Actual
                    
                    If Cells(Prow + X - 1, i).Value = 0 Then Cells(Prow + X - 1, i).Value = 1
                    
                    If i <> 21 And i <> 14 Then
                        Percent = (Cells(Prow + X, i).Value / games1 - Cells(Prow + X - 1, i).Value / games2) / (Cells(Prow + X - 1, i).Value / games2)
                        pdiff.Cells(Prow + X, i).Value = Percent
                    Else
                        pdiff.Cells(Prow + X, i).Value = "-"
                    End If
                    
                    If Actual >= 0 Then
                        Binomial = "I"
                    Else
                        Binomial = "D"
                    End If
                    Bdiff.Cells(Prow + X, i).Value = Binomial
                Else
                    adiff.Cells(Prow + X, i).Value = "N"
                    pdiff.Cells(Prow + X, i).Value = "N"
                    Bdiff.Cells(Prow + X, i).Value = "N"
                End If
            Next i
        Else
            For i = 5 To 21
                adiff.Cells(Prow + X, i).Value = "-"
                pdiff.Cells(Prow + X, i).Value = "-"
                Bdiff.Cells(Prow + X, i).Value = "-"
            Next i
        End If
        X = X - 1
    Loop
Next Nrow

End Sub

Sub ALL()
Call CorrectSeasons
Call Stat_Game_Team
Call Adjust_Needed

End Sub

Sub CorrectSeasons()
'Calculates goals, assists, and points for specified NHL seasons
Dim SAsheet As Worksheet, Season As Integer
Dim goals As Integer, assists As Integer, points As Integer, Row As Integer

Set SAsheet = Sheets("Adjustment")

SAsheet.Cells(1, 1).Value = "Season"
SAsheet.Cells(1, 2).Value = "Goals"
SAsheet.Cells(1, 3).Value = "Assists"
SAsheet.Cells(1, 4).Value = "Points"
SAsheet.Cells(1, 5).Value = "Season Length"
SAsheet.Cells(1, 6).Value = "Number of Teams"

For Season = 1992 To year1
    If Season <> 2004 Then
        Row = Season - 1990
        Sheets(Season & "|" & (Season + 1)).Select
        goals = WorksheetFunction.Sum(Range(Cells(2, 4), Cells(2, 4).End(xlDown)))
        assists = WorksheetFunction.Sum(Range(Cells(2, 5), Cells(2, 5).End(xlDown)))
        points = WorksheetFunction.Sum(Range(Cells(2, 6), Cells(2, 6).End(xlDown)))
        SAsheet.Cells(Row, 1).Value = Season & "|" & (Season + 1)
        SAsheet.Cells(Row, 2).Value = goals
        SAsheet.Cells(Row, 3).Value = assists
        SAsheet.Cells(Row, 4).Value = points
    End If
Next Season

End Sub

Sub Stat_Game_Team()
'Determine Goals, Assists, and Points per game per team for each season
Dim SA As Variant, Srow As Integer

SA = Sheets("Adjustment").UsedRange
Sheets("Adjustment").Select


For Srow = (LBound(SA) + 1) To UBound(SA, 1)
    If IsEmpty(SA(Srow, 1)) = False Then
        Cells(Srow, 7).Value = Cells(Srow, 2) / (Cells(Srow, 5).Value * Cells(Srow, 6).Value)
        Cells(Srow, 8).Value = Cells(Srow, 3) / (Cells(Srow, 5).Value * Cells(Srow, 6).Value)
        Cells(Srow, 9).Value = Cells(Srow, 4) / (Cells(Srow, 5).Value * Cells(Srow, 6).Value)
    End If
Next Srow

End Sub

Sub Adjust_Needed()
'Determine multiplier for goals, asists, and points to match to current season
Dim SA As Variant, Srow As Integer
Dim FixedSeason As Integer
Dim goals As Double, assists As Double, points As Double

SA = Sheets("Adjustment").UsedRange
Sheets("Adjustment").Select

For Srow = LBound(SA) + 1 To UBound(SA, 1)
    If SA(Srow, 1) = year1 & "|" & (year1 + 1) Then
        goals = SA(Srow, 7)
        assists = SA(Srow, 8)
        points = SA(Srow, 9)
    End If
Next Srow

For Srow = (LBound(SA) + 1) To UBound(SA, 1)
    If IsEmpty(SA(Srow, 1)) = False Then
        Cells(Srow, 10).Value = goals / SA(Srow, 7)
        Cells(Srow, 11).Value = assists / SA(Srow, 8)
        Cells(Srow, 12).Value = points / SA(Srow, 9)
    End If
Next Srow

End Sub

Sub Adjust_All_Stats()
'Use the calculated multipliers and apply them to consolidated Hockey Stats worksheet
Dim Mult1 As Integer, HockeyStats As Variant, Adjust As Variant
Dim Row As Integer, Col As Integer, yrow As Integer, Ycol As Integer
Dim Gmult As Integer, Amult As Integer, Arow As Integer
Dim Season As String

Mult1 = 1992
Row = 2 'what row is first analyzed
Col = 5 ' what column is first analyzed
Gmult = 10 'column where goal multiplier is
Amult = 11 ' " Assist "

Adjust = Sheets("Adjustment").UsedRange
HockeyStats = Sheets("Hockey Stats").UsedRange
Sheets("Hockey Stats").Select
'goes through each row of entire sheet
'a multiplier value is determiend based on the season
For yrow = Row To UBound(HockeyStats, 1)
    Season = HockeyStats(yrow, 1)
    For Arow = LBound(Adjust) To UBound(Adjust, 1)
        If Adjust(Arow, 1) = Season Then
            Cells(yrow, Col).Value = Cells(yrow, Col).Value * Adjust(Arow, Gmult) 'Goals column
            Cells(yrow, Col + 1).Value = Cells(yrow, Col + 1).Value * Adjust(Arow, Amult) 'Assists Column
            Cells(yrow, Col + 2).Value = Cells(yrow, Col).Value + Cells(yrow, Col + 1).Value 'Points Column
            Exit For
        End If
    Next Arow
Next yrow

End Sub
