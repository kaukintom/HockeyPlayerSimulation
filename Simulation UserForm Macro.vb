Option Explicit On

Public numsim As Integer
Const Row As Integer = 20
Const StatRow As Integer = 3
Const year As Integer = 2016
Const Pyear As Integer = year - 1
Public EnableEvents As Boolean

Private Sub UserForm_Initialize()
    Me.EnableEvents = True
    Unload Sorting
'**************************************************************
    'UserForm Location
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + Application.Width - Me.Width - 25
    '**************************************************************
    If ActiveSheet.Name = "Fantasy Points Simulator" Then
        ScrollBar1.min = 50
        ScrollBar1.Max = 2000
        ScrollBar1.SmallChange = 50
        ScrollBar1.LargeChange = 50
        ScrollBar1.Value = 50
    Else
        ScrollBar1.min = 10000
        ScrollBar1.Max = 20000
        ScrollBar1.SmallChange = 500
        ScrollBar1.LargeChange = 500
        ScrollBar1.Value = 10000
    End If
    Simulation.Label2.Caption = ScrollBar1.Value
    numsim = ScrollBar1.Value
End Sub

Private Sub ScrollBar1_Change()

    numsim = ScrollBar1.Value
    Simulation.Label2.Caption = numsim
End Sub

Private Sub ToggleButton1_Click()
    Dim Time As Double

    Call Simulated_Values(Time)
    'Output to userform
    Simulation.Label3.Caption = numsim & " simulations took " & Time & " seconds to complete."

End Sub

'Simulation code begins
'*******************************************************************************************

Sub Simulated_Values(Time)
    Dim LastCol As Integer, fsheet As Worksheet, Asheet As Worksheet, Gsheet As Worksheet, Lsheet As Worksheet
    Dim FPS As Variant, Scol As Integer
    Dim Name As String, Sbegin As Integer, srow As Integer, GoalSim As Variant, AssistSim As Variant
    Dim Lrow As Integer, LastSeason As Variant, ADist As String, GDist As String
    Dim G As Integer, A As Integer, assist As Variant, Goal As Variant
    Dim Gmean As Double, Gsd As Double, Amean As Double, Asd As Double
    Dim Atot As Double, Gtot As Double, Acat As Integer, Gcat As Integer, games As Integer
    Dim Frow As Integer, GPG As Double, APG As Double, GperS As Integer
    Dim FPScol As Integer, FPSrow As Integer, PN As Double, Grnd As Double
    Dim ALoc As Integer, GLoc As Integer, NamesList As Variant
    Dim StartTime As Double, GP As Integer, TotStat As Double, AvgStat As Double
    Dim Tsim As Integer, ST1 As Variant, st2 As Double, x As Integer, st3 As Double
    Dim ElapsedTime As Double, StartTime2 As Double
    '*************************************************************
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Me.EnableEvents = False
    '*************************************************************
    StartTime = Timer
    ReDim ST1(1 To 10, 1 To 1)
    GLoc = 3 'Worksheet location of simulated Goals relative to FPS Array
    ALoc = 2
    Sbegin = 3 'where hockey players begin on Lsheet
    G = 21 'column for Goal data for Distributions
    A = 22 'column for Assist Data for distributions
    Gcat = 4 'Column for Goals
    Acat = 5 'column for assists
    games = 3 'column for games
    GperS = 82 'games played in hockey season
    TotStat = 2
    Tsim = 3
    AvgStat = 4

    '************************************************************
    'Setup up worksheets

    If ActiveSheet.Name = "Fantasy Points Simulator" Then
        Set fsheet = Sheets("Fantasy Points Simulator")
        Set Lsheet = Sheets("Latest Season Data")
    Else
        Set fsheet = Sheets("Prior Season Sim")
        Set Lsheet = Sheets("Prior Season Data")
    End If

    Set Gsheet = Sheets("Simulations(G)")
    Set Asheet = Sheets("Simulations(A)")

    fsheet.Select
    Cells(Row, 1).Select
    FPS = fsheet.Cells(Row, 1).CurrentRegion

    Application.CutCopyMode = False
    'copy the hockey player data to the simulation worksheets
    Call setup(Sbegin, Asheet, Gsheet, numsim, fsheet, FPS)
    'Get list of hockey players to be simulated and put them in the _
    'simulations worksheet
    fsheet.Select
    NamesList = Range(Cells(Row + 2, 1), Cells(UBound(FPS, 1) + Row - 1, 1))
    Gsheet.Select
    Range(Cells(3, 1), Cells(UBound(FPS, 1), 1)) = NamesList
    Asheet.Select
    Range(Cells(3, 1), Cells(UBound(FPS, 1), 1)) = NamesList

    '*******************************************************************
    'Define the arrays to be used
    GoalSim = Gsheet.UsedRange
    AssistSim = Asheet.UsedRange
    LastSeason = Lsheet.UsedRange
    FPS = fsheet.Cells(Row, 1).CurrentRegion
    'loop through each name in GoalSim Array, this array is identical to AssistSim and FPS
    'for each name, simulation takes place and the values are added to an array (FPS) and then _
    placed in the Fantasy Points Simulator worksheet

    For srow = LBound(GoalSim) + 2 To UBound(GoalSim, 1)
        Name = GoalSim(srow, 1)
        ' find name and values to be used in probabilities
        For Lrow = LBound(LastSeason) To UBound(LastSeason, 1)
            If LastSeason(Lrow, 1) = Name Then
                GDist = LastSeason(Lrow, G)
                ADist = LastSeason(Lrow, A)
                Exit For
            End If
        Next Lrow
        'Split arrays to find mean and standard deviation for each stat
        If GDist <> "|" And GDist <> "" And ADist <> "|" And ADist <> "" Then
            assist = Split(ADist, "|")
            Goal = Split(GDist, "|")
            Gmean = Goal(LBound(Goal))
            Gsd = Goal(UBound(Goal))
            Amean = assist(LBound(assist))
            Asd = assist(UBound(assist))
            'the following variables help find Sum of total differences
            Gtot = 0 ' need to be set to zero again for a new hockey player
            Atot = 0
            'Conduct specified number of simulations for current player

            Call Lookup(Gsd, Asd, Gmean, Amean, Gtot, Atot, Scol, srow, fsheet)

            'find the average change in a stat category
            '        If fsheet.Name = "Fantasy Points Simulator" Then
            '            GoalSim(srow, Scol) = Gtot / numsim
            '            AssistSim(srow, Scol) = Atot / numsim
            '        Else
            'outputs for Simulations results worksheets (Simulations(G) and Simulations (A))
            GoalSim(srow, TotStat) = Gtot 'Simulations summed
            GoalSim(srow, Tsim) = numsim 'Simulations conducted
            GoalSim(srow, AvgStat) = Gtot / numsim ' average value for simulations
            AssistSim(srow, TotStat) = Atot
            AssistSim(srow, Tsim) = numsim
            AssistSim(srow, AvgStat) = Atot / numsim
            '        End If
            'now the previous seasons stat category, is divided by the number of games _
            'so that the stat can be compared to simulated values
            For Frow = LBound(LastSeason) To UBound(LastSeason, 1)
                If LastSeason(Frow, 1) = Name Then
                    GP = LastSeason(Frow, games)
                    GPG = LastSeason(Frow, Gcat) / LastSeason(Frow, games)
                    APG = LastSeason(Frow, Acat) / LastSeason(Frow, games)
                    Exit For
                End If
            Next Frow
            'Add the simulated averages to the previous season stats per a game to get the simulated _
            'New season averages
            'these values are put into an array and will be displayed on the worksheet later
            '**********************************************************************************
            'Simulated values are put into the array here, using the locations provided _
            'by ALoc And GLoc
            If GP < 15 Then GPG = -1 And APG = -1 ' might be an inflated goals per a game and assists per a game _
            'If they Then dont play too many games And score a lot Of points
            'particularly concerning players that only play 1 game and get 1+ points
            '        If fsheet.Name = "Fantasy Points Simulator" Then
            '            If CInt((GPG + GoalSim(srow, Scol)) * GperS) < 0 Then
            '                FPS(srow, UBound(FPS, 2) - GLoc) = 0
            '            Else
            '                FPS(srow, UBound(FPS, 2) - GLoc) = CInt((GPG + GoalSim(srow, Scol)) * GperS)
            '            End If
            '
            '            If CInt((APG + AssistSim(srow, Scol)) * GperS) < 0 Then
            '                FPS(srow, UBound(FPS, 2) - ALoc) = 0
            '            Else
            '                FPS(srow, UBound(FPS, 2) - ALoc) = CInt((APG + AssistSim(srow, Scol)) * GperS)
            '            End If
            '            FPS(srow, UBound(FPS, 2) - ALoc + 1) = FPS(srow, UBound(FPS, 2) - ALoc) + FPS(srow, UBound(FPS, 2) - GLoc)
            '        Else
            If CInt((GPG + GoalSim(srow, AvgStat)) * GperS) < 0 Then
                FPS(srow, UBound(FPS, 2) - GLoc) = 0
            Else
                FPS(srow, UBound(FPS, 2) - GLoc) = CInt((GPG + GoalSim(srow, AvgStat)) * GperS)
            End If

            If CInt((APG + AssistSim(srow, AvgStat)) * GperS) < 0 Then
                FPS(srow, UBound(FPS, 2) - ALoc) = 0
            Else
                FPS(srow, UBound(FPS, 2) - ALoc) = CInt((APG + AssistSim(srow, AvgStat)) * GperS)
            End If
            FPS(srow, UBound(FPS, 2) - ALoc + 1) = FPS(srow, UBound(FPS, 2) - ALoc) + FPS(srow, UBound(FPS, 2) - GLoc)
            '        End If
            '***************************************************************************************
        End If
    Next srow

    'The Arrays containing all the simulated values is put into the worksheet
    For FPSrow = LBound(FPS) To UBound(FPS, 1)
        For FPScol = LBound(FPS) To UBound(FPS, 2)
            fsheet.Cells(FPSrow + Row - 1, FPScol).Value = FPS(FPSrow, FPScol)
        Next FPScol
    Next FPSrow
    'The Arrays containing the simulated values are put in their respective worksheets
    Gsheet.Select
    Range(Cells(LBound(GoalSim), LBound(GoalSim)), Cells(UBound(GoalSim, 1), UBound(GoalSim, 2))) = GoalSim
    Asheet.Select
    Range(Cells(LBound(AssistSim), LBound(AssistSim)), Cells(UBound(AssistSim, 1), UBound(AssistSim, 2))) = AssistSim


    Call New_FP(FPS, fsheet, GLoc, ALoc)

    Call ColourChangeResults(fsheet)

    Call Interior_Color(fsheet)

    Call SimulatedColours(fsheet)

    Call Stats_Section_Maintain(fsheet)

    Call Title_Maintain(fsheet)

    If fsheet.Name = "Prior Season Sim" Then Call PriorSeasonComparison()
    '**************************************************

    Call SecondsCount(srow, numsim, StartTime, fsheet, Time)
    '***********************************************
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Me.EnableEvents = True
    Exit Sub 'Workaround for disapearing userform when macro finishes
End Sub

Sub setup(Sbegin, Asheet, Gsheet, numsim, fsheet, FPS)
    Dim x As Integer, numFP As Integer
    'setup goal simulation sheet
    Gsheet.Select
    Gsheet.UsedRange.delete 'delete everything on sheet
    'Place titles on sheet
    Cells(Sbegin - 2, 1).Value = "Goals"
    Cells(Sbegin - 1, 1).Value = "Player"
    'If the simulation is for current season hockey players then _
    'all the simulations will be shown

    'If fsheet.Name = "Fantasy Points Simulator" Then
    '    For x = 1 To numsim
    '        Cells(Sbegin - 1, x + 1).Value = x
    '    Next x
    '    Cells(Sbegin - 1, numsim + 2).Value = "Average Goals"
    'Else

    'If the simulation is for the prior season, then more simulations are conducted _
    'which can cause issues with the workbook
    'only the number of simulations, total for the simulation, and average stat per _
    'game will be shown
    Cells(Sbegin - 1, 2).Value = "Total Goals for Sim"
    Cells(Sbegin - 1, 3).Value = "Total Simulations"
    Cells(Sbegin - 1, 4).Value = "Average Goals"
    'End If

    'setup assist simulation sheet (same as above setup)
    Asheet.Select
    Asheet.UsedRange.delete
    Cells(Sbegin - 2, 1).Value = "Assists"
    Cells(Sbegin - 1, 1).Value = "Player"

    'If fsheet.Name = "Fantasy Points Simulator" Then
    '    For x = 1 To numsim
    '        Cells(Sbegin - 1, x + 1).Value = x
    '    Next x
    '    Cells(Sbegin - 1, numsim + 2).Value = "Average Assists"
    'Else

    Cells(Sbegin - 1, 2).Value = "Total Assists for Sim"
    Cells(Sbegin - 1, 3).Value = "Total Simulations"
    Cells(Sbegin - 1, 4).Value = "Average Assists"

    'End If
    fsheet.Select
    'adds 4 more columns to Fantasy Points Simulator Worksheet
    'First column
    For x = 1 To 100
        'If the column for goals exists, then leave the current page setup, _
        'otherwise, add the column for goals
        If fsheet.Cells(Row + 1, x).Value = "Goals" Then
            Exit For
        ElseIf IsEmpty(fsheet.Cells(Row + 1, x).Value) = True Then
            fsheet.Cells(Row + 1, x).Value = "Goals"
            Range(Cells(Row + 1, x), Cells(Row + UBound(FPS, 1) - 1, x + 2)).BorderAround _
            ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
            Exit For
        End If
    Next x
    'Second column
    For x = 1 To 100
        If fsheet.Cells(Row + 1, x).Value = "Assists" Then
            Exit For
        ElseIf IsEmpty(fsheet.Cells(Row + 1, x).Value) = True Then
            fsheet.Cells(Row + 1, x).Value = "Assists"
            If fsheet.Name = "Fantasy Points Simulator" Then
                'for current season simulation
                fsheet.Cells(Row, x).Value = (year + 1) & "/" & (year + 2) & " Simulated Results"
            Else
                'for prior season simulation
                fsheet.Cells(Row, x).Value = (year) & "/" & (year + 1) & " Simulated Results"
            End If
            Range(Cells(Row, x - 1), Cells(Row, x + 2)).BorderAround _
            ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
            'merge the cells above the 4 new columns for simulated results, where the _
            'title Is located, And then add a border around it
            Range(Cells(Row, x - 1), Cells(Row, x + 2)).Merge
            Exit For
        End If
    Next x
    'Third column
    For x = 1 To 100
        If fsheet.Cells(Row + 1, x).Value = "Points" Then
            Exit For
        ElseIf IsEmpty(fsheet.Cells(Row + 1, x).Value) = True Then
            fsheet.Cells(Row + 1, x).Value = "Points"
            Exit For
        End If
    Next x
    'Fourth Column
    numFP = 0
    For x = 1 To 100
        If fsheet.Cells(Row + 1, x).Value = "Total FP" Then
            numFP = numFP + 1
            If numFP = 2 Then Exit For
        ElseIf IsEmpty(fsheet.Cells(Row + 1, x).Value) = True Then
            fsheet.Cells(Row + 1, x).Value = "Total FP"
            Range(Cells(Row + 1, x), Cells(Row - 1 + UBound(FPS, 1), x)).BorderAround _
            ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
            Exit For
        End If
    Next x

End Sub

Sub Lookup(Gsd, Asd, Gmean, Amean, Gtot, Atot, Scol, srow, fsheet)
'This Macro conducts the simulation for each player using random numbers _
And the NormInv function
Dim RanNum As Double

    'If fsheet.Name = "Fantasy Points Simulator" Then
    '    each simulation is shown and is put into an array to be placed on the worksheet at _
    '    a later stage of the macro
    '    Randomize 'makes sure number will be random
    '    For Scol = 2 To numsim + 1 'starting at from 2 and adding 1 to numsim makes for easy placement _
    '    on the worksheet
    '
    '        RanNum = Rnd()
    '
    '        Do Until RanNum > 0 And RanNum < 1
    '            Randomize
    '            RanNum = Rnd()
    '        Loop
    '
    '        GoalSim(srow, Scol) = Application.WorksheetFunction.NormInv(RanNum, Gmean, Gsd) 'Value placed into array
    '        Gtot = Gtot + GoalSim(srow, Scol) 'value from array added to a total
    '
    '        RanNum = Rnd()

    'Do Until RanNum > 0 And RanNum < 1
    'Randomize
    'RanNum = Rnd()
    'Loop

    'AssistSim(srow, Scol) = Application.WorksheetFunction.NormInv(RanNum, Amean, Asd)
    'Atot = Atot + AssistSim(srow, Scol)
    'Next Scol
    'Else
    Randomize
    For Scol = 2 To numsim + 1
        'unknown issues with this setup where the random number can be 0 or 1, so it is _
        'reset to 0 every loop And placed in a do loop until it Is Not 0 Or 1
        RanNum = Rnd()

        Do Until RanNum > 0 And RanNum < 1
            Randomize
            RanNum = Rnd()
        Loop

        Gtot = Gtot + Application.WorksheetFunction.NormInv(RanNum, Gmean, Gsd)

        RanNum = Rnd()

        Do Until RanNum > 0 And RanNum < 1
            Randomize
            RanNum = Rnd()
        Loop

        Atot = Atot + Application.WorksheetFunction.NormInv(RanNum, Amean, Asd)
    Next Scol
    'End If

End Sub


Sub ColourChangeResults(fsheet)
    Dim FPS As Variant, x As Integer, Frow As Integer
    Dim G As Integer, NSG As Integer
    Dim Goals As Integer, Assists As Integer
    Dim games As Integer

    fsheet.Select

    FPS = Cells(Row, 1).CurrentRegion
    NSG = 0
    'locates where the simulated data is, and the corresponding previous season data
    For x = 1 To 30
        If Cells(Row + 1, x).Value = "Goals" Then NSG = x
        If Cells(Row + 1, x).Value = "G" Then G = x
    Next x
    'if the simulated data is present (determined by whether the columns exist)
    If NSG <> 0 Then
        'for loop to go through every row of data
        For Frow = LBound(FPS) + 2 To UBound(FPS, 1)
            For x = 0 To 1
                'for the Assists and Goals columns of the simulated data
                If IsEmpty(Cells(Frow + Row - 1, NSG + x)) = False Then
                    'if there is a simulated value there
                    If x = 0 Then Goals = Cells(Frow + Row - 1, G + x).Value
                    If x = 1 Then Assists = Cells(Frow + Row - 1, G + x).Value
                    'store the values in a variables
                    'if the simulated values are greater than the 82 game adjusted values from the previous season _
                    'the simulated value Is coloured green, If its less than the 82 game adjusted values _
                    'the simulated value Is coloured red, If the previous And simulated values are identical _
                    'the coloured remains at black
                    If Cells(Frow + Row - 1, NSG + x).Value > CInt(82 * Cells(Frow + Row - 1, G + x) / Cells(Frow + Row - 1, G - 1).Value) Then
                        Cells(Frow + Row - 1, NSG + x).Font.Color = RGB(0, 170, 0)
                    ElseIf Cells(Frow + Row - 1, NSG + x).Value < CInt(82 * Cells(Frow + Row - 1, G + x) / Cells(Frow + Row - 1, G - 1).Value) Then
                        Cells(Frow + Row - 1, NSG + x).Font.Color = RGB(200, 0, 0)
                    Else
                        Cells(Frow + Row - 1, NSG + x).Font.Color = RGB(0, 0, 0)
                    End If
                    'total points are handled separately since there is no previous season points column
                    If x = 1 Then
                        If Cells(Frow + Row - 1, NSG + x + 1).Value > CInt(82 * (Goals + Assists) / Cells(Frow + Row - 1, G - 1).Value) Then
                            Cells(Frow + Row - 1, NSG + x + 1).Font.Color = RGB(0, 170, 0)
                        ElseIf Cells(Frow + Row - 1, NSG + x + 1).Value < CInt(82 * (Goals + Assists) / Cells(Frow + Row - 1, G - 1).Value) Then
                            Cells(Frow + Row - 1, NSG + x + 1).Font.Color = RGB(200, 0, 0)
                        Else
                            Cells(Frow + Row - 1, NSG + x + 1).Font.Color = RGB(0, 0, 0)
                        End If
                    End If
                End If
            Next x
        Next Frow
    End If
End Sub

Sub SecondsCount(srow, numsim, StartTime, fsheet, Time)
    'Finds how long it took to conduct simulation
    Dim Utime As Integer, Mtime As Integer, T As Integer

    Utime = 3
    Mtime = 2
    T = Utime
    Sheets("Simulation Time").Select
    Cells(1, 2).Select
    For srow = 1 To 101
        'If the macro has not been run before for the specified number of simulations, _
        'it will add the number of simulations at the end of the column as well as the time
        If IsEmpty(Cells(srow, 1)) = True Then
            Time = Round(Timer - StartTime, 2)
            Cells(srow, 1).Value = numsim
            Cells(srow, T).Value = Time
            Exit For
            'If the macro has been run for the specified number of simulations _
            'before, then it will place a time next to the number of simulations
        ElseIf Cells(srow, 1).Value = numsim Then
            Time = Round(Timer - StartTime, 2)
            Cells(srow, T).Value = Time
            Exit For
        End If
    Next srow
    fsheet.Select
End Sub

Sub SimulatedColours(fsheet)
    Dim Sim As Variant, srow As Integer

    Sheets("Simulations(A)").Select
    Sim = Sheets("Simulations(A)").UsedRange
    'Make sure everything initially defaults to white
    For srow = LBound(Sim) + 2 To UBound(Sim, 1) Step 1
        Range(Cells(srow, 1), Cells(srow, UBound(Sim, 2))).Interior.Color = RGB(255, 255, 255)
    Next srow
    'Then place orange row at the top where the simulation numbers are
    Range(Cells(2, 1), Cells(2, UBound(Sim, 2))).Interior.Color = RGB(255, 230, 200)
    'Afterwards every second row is coloured a light blue
    For srow = LBound(Sim) + 2 To UBound(Sim, 1) Step 2
        Range(Cells(srow, 1), Cells(srow, UBound(Sim, 2))).Interior.Color = RGB(200, 230, 255)
    Next srow
    ' Makes sure all information to the right of the first column is centred
    Range(Cells(2, 2), Cells(UBound(Sim, 1), UBound(Sim, 2))).HorizontalAlignment = xlCenter
    'Add a border around the cells containing the simulation number
    Range(Cells(2, 2), Cells(UBound(Sim, 1), UBound(Sim, 2))).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    'add a border around all the simulated values of the hockey players
    Range(Cells(3, 2), Cells(UBound(Sim, 1), UBound(Sim, 2))).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
'   add a border around the column containing the averages of the simulations
    Range(Cells(2, UBound(Sim, 2)), Cells(UBound(Sim, 1), UBound(Sim, 2))).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    'Same process for the other simulated page

    Sheets("Simulations(G)").Select

    Sim = Sheets("Simulations(G)").UsedRange
    For srow = LBound(Sim) + 2 To UBound(Sim, 1) Step 1
        Range(Cells(srow, 1), Cells(srow, UBound(Sim, 2))).Interior.Color = RGB(255, 255, 255)
    Next srow
    Range(Cells(2, 1), Cells(2, UBound(Sim, 2))).Interior.Color = RGB(255, 230, 200)
    For srow = LBound(Sim) + 2 To UBound(Sim, 1) Step 2
        Range(Cells(srow, 1), Cells(srow, UBound(Sim, 2))).Interior.Color = RGB(200, 230, 255)
    Next srow
    Range(Cells(2, 2), Cells(UBound(Sim, 1), UBound(Sim, 2))).HorizontalAlignment = xlCenter
    Range(Cells(2, 2), Cells(UBound(Sim, 1), UBound(Sim, 2))).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    Range(Cells(3, 2), Cells(UBound(Sim, 1), UBound(Sim, 2))).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    Range(Cells(2, UBound(Sim, 2)), Cells(UBound(Sim, 1), UBound(Sim, 2))).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium

End Sub


Sub New_FP(FPS, fsheet, GLoc, ALoc)
    'finds the fantasy points for the simulated values
    Dim TFP As Integer, Season As Variant
    Dim FP As Double, srow As Integer, x As Integer

    TFP = 1 'columns from the end of the current region

    Season = fsheet.Cells(StatRow, 1).CurrentRegion 'array for the Fantasy Points
    'first calculates the values in an array
    For x = LBound(FPS) + 2 To UBound(FPS, 1)
        FPS(x, UBound(FPS, 2)) = FPS(x, UBound(FPS, 2) - ALoc) * Season(3, 2) + FPS(x, UBound(FPS, 2) - GLoc) * Season(2, 2)
    Next x
    'puts the column from the array into the worksheet
    For x = LBound(FPS) + 2 To UBound(FPS, 1)
        fsheet.Cells(Row + x - 1, UBound(FPS, 2)).Value = FPS(x, UBound(FPS, 2))
    Next x

End Sub

Sub PriorSeasonComparison()
    Dim NextSeason As Worksheet, fsheet As Worksheet
    Dim NS As Variant, FS As Variant, col As Integer
    Dim NSrow As Integer, Name As String, FSrow As Integer
    Dim NSG As Double, NSA As Double, NSP As Double
    Dim FSG As Double, FSA As Double, FSP As Double
    Dim FScol As Integer
    Dim Err As Double, UnderErr As Double, OverErr As Double
    Dim Pcount As Integer, ErrorArray As Variant, Cat As Integer
    Dim Goals As Integer, Assists As Integer, Points As Integer
    Dim LastSeason As Worksheet, LS As Variant, LSrow As Integer
    Dim SimErr As Worksheet
    Dim GUP As Integer, AUP As Integer, PUP As Integer
    Dim GDOWN As Integer, ADOWN As Integer, PDOWN As Integer
    Dim GNeither As Integer, ANeither As Integer, PNeither As Integer
    Dim i As Integer, UP As Integer, DOWN As Integer, Neither As Integer
    Dim NSGP As Integer, ErrRow As Integer
    Dim UArray As Variant, OArray As Variant, x As Integer
    Dim LSIdentifier As Boolean

    col = 1 'Simple starting position to be referred to
    Pcount = 0 ' a counting variable for number of simulated results
    Cat = 0
    UP = 0
    DOWN = 0

    NSGP = 3
    NSG = 4
    NSA = 5
    NSP = 6

    'Defining the various worksheets and arrays
    Set NextSeason = Sheets("Latest Season Data")
    Set LastSeason = Sheets("Prior Season Data")
    Set fsheet = Sheets("Prior Season Sim")
    Set SimErr = Sheets("Simulation Error")

    LS = LastSeason.UsedRange
    NS = NextSeason.UsedRange
    FS = fsheet.Cells(Row, 1).CurrentRegion

    SimErr.UsedRange.delete 'delete what is currently on the output worksheet _
    'For this macro

    'The column location for the simulated stats need to be discovered , _
    'And once they are found, the next column can be discovered
    For FScol = LBound(FS) To UBound(FS, 2)
        If FS(LBound(FS) + 1, FScol) = "Goals" Then
            FSG = FScol 'location on fsheet
            Cat = Cat + 1 'signals the column is found
            Goals = 1 'location in a new array to be created
        ElseIf FS(LBound(FS) + 1, FScol) = "Assists" Then
            FSA = FScol
            Cat = Cat + 1
            Assists = 2
        ElseIf FS(LBound(FS) + 1, FScol) = "Points" Then
            FSP = FScol
            Cat = Cat + 1
            Points = 3
        End If
    Next FScol

    'There needs to be a mean and variance(M/V) that exists for the hockey player, for the _
    'New And previous seasons in order to find out the difference in the projection And _
    'actual performance
    'An array is constructed from counting the hockey players
    For NSrow = LBound(NS) + 2 To UBound(NS, 1) 'loops through each name from newest season data
        'Now the player is checked to see if the M/V exists
        If (NS(NSrow, UBound(NS, 2)) <> "|" And NS(NSrow, UBound(NS, 2) - 1) <> "|") Then
            Name = NS(NSrow, col) ' if M/V exists then keep track of name
            For LSrow = LBound(LS) To UBound(LS, 1) 'Find this player in the previous season
                If LS(LSrow, 1) = Name Then
                    'if the M/V exists for this season as well then add them to the total player count
                    If (LS(LSrow, UBound(LS, 2)) <> "|" And LS(LSrow, UBound(LS, 2) - 1) <> "|") Then
                        Pcount = Pcount + 1
                    End If
                    Exit For
                End If
            Next LSrow
        End If
    Next NSrow

    ReDim ErrorArray(1 To Pcount, 1 To Cat)
    Pcount = 0
    'There needs to be a mean and variance(M/V) that exists for the hockey player, for the _
    'New And previous seasons in order to find out the difference in the projection And _
    'actual performance
    For NSrow = LBound(NS) + 2 To UBound(NS, 1)
            If (NS(NSrow, UBound(NS, 2)) <> "|" And NS(NSrow, UBound(NS, 2) - 1) <> "|") Then
                Name = NS(NSrow, col)
                LSIdentifier = False 'Resets to false when M/V exists for new season
                For LSrow = LBound(LS) To UBound(LS, 1)
                    If LS(LSrow, 1) = Name Then
                        If (LS(LSrow, UBound(LS, 2)) <> "|" And LS(LSrow, UBound(LS, 2) - 1) <> "|") Then
                            LSIdentifier = True 'True when M/V exists for new season as well
                        End If
                        Exit For
                    End If
                Next LSrow
            'When LSIdentifier is true, then the difference in goals, assists, and points can be found _
            'between projection And actual performance
            If LSIdentifier = True Then
                For FSrow = LBound(FS) + 1 To UBound(FS, 1)
                    If FS(FSrow, col) = Name Then
                        Pcount = Pcount + 1
                        'need to convert from points per a game to total points for the season to compare the stats
                        ErrorArray(Pcount, Goals) = Round(NS(NSrow, NSG) / NS(NSrow, NSGP) * 82, 0) - FS(FSrow, FSG)
                        ErrorArray(Pcount, Assists) = Round(NS(NSrow, NSA) / NS(NSrow, NSGP) * 82, 0) - FS(FSrow, FSA)
                        ErrorArray(Pcount, Points) = Round(NS(NSrow, NSP) / NS(NSrow, NSGP) * 82, 0) - FS(FSrow, FSP)
                        'The difference in the prior season stats and projections are compared to the difference in the new season _
                        'And the prior season
                        For LSrow = LBound(LS) To UBound(LS, 1)
                            If LS(LSrow, col) = Name Then
                                'If both the projection and new season were a positive difference compared to _
                                'last season stat, then count this
                                If NS(NSrow, NSG) / NS(NSrow, NSGP) * 82 >= LS(LSrow, NSG) / LS(LSrow, NSGP) * 82 And FS(FSrow, FSG) >= LS(LSrow, NSG) / LS(LSrow, NSGP) * 82 Then
                                    GUP = GUP + 1
                                    'If both the projection and new season were a negative difference compared to _
                                    'last season stat, then count this
                                ElseIf NS(NSrow, NSG) / NS(NSrow, NSGP) * 82 < LS(LSrow, NSG) / LS(LSrow, NSGP) * 82 And FS(FSrow, FSG) < LS(LSrow, NSG) / LS(LSrow, NSGP) * 82 Then
                                    GDOWN = GDOWN + 1
                                Else
                                    'If the projection and new season dont have the same sign of difference _
                                    'compared to last season stat, then count this
                                    GNeither = GNeither + 1
                                End If

                                If NS(NSrow, NSA) / NS(NSrow, NSGP) * 82 >= LS(LSrow, NSA) / LS(LSrow, NSGP) * 82 And FS(FSrow, FSA) >= LS(LSrow, NSA) / LS(LSrow, NSGP) * 82 Then
                                    AUP = AUP + 1
                                ElseIf NS(NSrow, NSA) / NS(NSrow, NSGP) * 82 < LS(LSrow, NSA) / LS(LSrow, NSGP) * 82 And FS(FSrow, FSA) < LS(LSrow, NSA) / LS(LSrow, NSGP) * 82 Then
                                    ADOWN = ADOWN + 1
                                Else
                                    ANeither = ANeither + 1
                                End If

                                If NS(NSrow, NSP) / NS(NSrow, NSGP) * 82 >= LS(LSrow, NSP) / LS(LSrow, NSGP) * 82 And FS(FSrow, FSP) >= LS(LSrow, NSP) / LS(LSrow, NSGP) * 82 Then
                                    PUP = PUP + 1
                                ElseIf NS(NSrow, NSP) / NS(NSrow, NSGP) * 82 < LS(LSrow, NSP) / LS(LSrow, NSGP) * 82 And FS(FSrow, FSP) < LS(LSrow, NSP) / LS(LSrow, NSGP) * 82 Then
                                    PDOWN = PDOWN + 1
                                Else
                                    PNeither = PNeither + 1
                                End If

                                Exit For
                            End If
                        Next LSrow
                        Exit For
                    End If
                Next FSrow
            End If
        End If
    Next NSrow

    'Setup table for difference in simulated values and realized values
    SimErr.Select
    SimErr.Cells(LBound(ErrorArray) + 1, col).Resize(UBound(ErrorArray), UBound(ErrorArray, 2)).Value = ErrorArray
    SimErr.Cells(1, Goals) = "Difference in Goals"
    SimErr.Cells(1, Assists) = "Difference in Assists"
    SimErr.Cells(1, Points) = "Difference in Points"
    'The values from counting how hockey players behaved relative to the projections are placed _
    'onto the worksheet
    For i = 0 To 2
        If i = 0 Then
            'Goals
            UP = GUP 'Resued variable to place on the worksheet easier through a loop
            DOWN = GDOWN
            Neither = GNeither
            SimErr.Cells(1, Points + 2 + (4 * i)) = "Goals" 'Places title on worksheet relative to table
        ElseIf i = 1 Then
            UP = AUP
            DOWN = ADOWN
            Neither = ANeither
            SimErr.Cells(1, Points + 2 + (4 * i)) = "Assists"
        Else
            UP = PUP
            DOWN = PDOWN
            Neither = PNeither
            SimErr.Cells(1, Points + 2 + (4 * i)) = "Points"
        End If

        'setup for various tables for goals, assists, and points
        'shows how well the simulation predicted ONLY direction (whether _
        'player increased Or decreased in the stat)
        SimErr.Cells(1, Points + 3 + (4 * i)) = "Number" 'column titles
        SimErr.Cells(1, Points + 4 + (4 * i)) = "Percent of Total"
        'row titles, with the count and percent of total players
        SimErr.Cells(2, Points + 2 + (4 * i)) = "Same Change"
        SimErr.Cells(2, Points + 3 + (4 * i)) = (UP + DOWN)
        SimErr.Cells(2, Points + 4 + (4 * i)) = Round((UP + DOWN) / Pcount * 100, 2)
        SimErr.Cells(3, Points + 2 + (4 * i)) = "Same Change Increase"
        SimErr.Cells(3, Points + 3 + (4 * i)) = UP
        SimErr.Cells(3, Points + 4 + (4 * i)) = Round(UP / Pcount * 100, 2)
        SimErr.Cells(4, Points + 2 + (4 * i)) = "Same Change Decrease"
        SimErr.Cells(4, Points + 3 + (4 * i)) = DOWN
        SimErr.Cells(4, Points + 4 + (4 * i)) = Round(DOWN / Pcount * 100, 2)
        SimErr.Cells(5, Points + 2 + (4 * i)) = "Incorrectly Predicted"
        SimErr.Cells(5, Points + 3 + (4 * i)) = Neither
        SimErr.Cells(5, Points + 4 + (4 * i)) = Round(Neither / Pcount * 100, 2)
    Next i

    ' find the error of the simulation with the realized values in the form of standard deviation _
    'as well as looking at the average value of the error values
    SimErr.Cells(7, Points + 3) = "Standard Deviation (SD)"
    SimErr.Cells(7, Points + 4) = "SD Over-estimated"
    SimErr.Cells(7, Points + 5) = "Mean Over-estimate"
    SimErr.Cells(7, Points + 6) = "SD Under-estimate"
    SimErr.Cells(7, Points + 7) = "Mean Under-estimate"

    'Calculates the total standard devation of all the error terms for a stat
    SimErr.Cells(8, Points + 2) = "Goals"
    SimErr.Cells(8, Points + 3) = WorksheetFunction.StDev_S(Range(Cells(LBound(ErrorArray) + 1, Goals), Cells(UBound(ErrorArray, 1) + 1, Goals)))
    SimErr.Cells(9, Points + 2) = "Assists"
    SimErr.Cells(9, Points + 3) = WorksheetFunction.StDev_S(Range(Cells(LBound(ErrorArray) + 1, Assists), Cells(UBound(ErrorArray, 1) + 1, Assists)))
    SimErr.Cells(10, Points + 2) = "Points"
    SimErr.Cells(10, Points + 3) = WorksheetFunction.StDev_S(Range(Cells(LBound(ErrorArray) + 1, Points), Cells(UBound(ErrorArray, 1) + 1, Points)))

    SimErr.Cells(10, Points + 2).Interior.Color = RGB(200, 230, 255)
    SimErr.Cells(9, Points + 2).Interior.Color = RGB(200, 230, 255)
    SimErr.Cells(8, Points + 2).Interior.Color = RGB(200, 230, 255)
    'each stat will have the standard deviation and mean found for them for when the _
    'simulation over - estimated Or under - estimated the values compared to the realized values
    For x = 0 To 2 'helps find relative location of stat in ErrorArray
        ReDim OArray(1, 1) 'reset size of array
        ReDim UArray(1, 1)
        For ErrRow = 1 To UBound(ErrorArray, 1) 'check through error array
            'calls a sub that keeps track of the underestimated and overestimated values _
            'through 2 dynamically expanding arrays
            Call AssignErrors(ErrorArray, ErrRow, col, x, OArray, UArray)

        Next ErrRow
        'Standard deviations and means are placed on the worksheet from the arrays _
        'created From ErrorArray
        SimErr.Cells(8 + x, Points + 4) = Round(WorksheetFunction.StDev_S(OArray), 2)
        SimErr.Cells(8 + x, Points + 5) = Round(WorksheetFunction.Average(OArray), 2)
        SimErr.Cells(8 + x, Points + 6) = Round(WorksheetFunction.StDev_S(UArray), 2)
        SimErr.Cells(8 + x, Points + 7) = Round(WorksheetFunction.Average(UArray), 2)
    Next x

    'These macros maintain the worksheet visuals
    Call Interior_Color(fsheet) 'colours rows
    Call Title_Maintain(fsheet) 'maintains title locations and structure
    Call Stats_Section_Maintain(fsheet) 'table for user inputs has white background
    Call SimulationErrorMaintain() 'background colour and borders for Simulation Error Worksheet
End Sub

Sub AssignErrors(ErrorArray, ErrRow, col, x, OArray, UArray)
    'when the simulation over-estimated the values
    If ErrorArray(ErrRow, col + x) >= 0 Then
        If IsEmpty(OArray(1, 1)) = True Then 'fills in first value of array
            OArray(1, 1) = ErrorArray(ErrRow, col + x)
        Else
            'when the array is full, which will always happen once the first value _
            'Is placed inside, the array Is transposed And a column Is added, then transposed again
            'This allows rows to be added to the array
            OArray = Application.Transpose(OArray)
            ReDim Preserve OArray(1 To UBound(OArray, 1), 1 To UBound(OArray, 2) + 1)
            OArray = Application.Transpose(OArray)
            'the value is now placed inside the newly created row in the array
            OArray(UBound(OArray, 1), 1) = ErrorArray(ErrRow, col + x)
        End If
        'same process as above
    ElseIf ErrorArray(ErrRow, col + x) < 0 Then
        If IsEmpty(UArray(1, 1)) = True Then
            UArray(1, 1) = ErrorArray(ErrRow, col + x)
        Else
            UArray = Application.Transpose(UArray)
            ReDim Preserve UArray(1 To UBound(UArray, 1), 1 To UBound(UArray, 2) + 1)
            UArray = Application.Transpose(UArray)

            UArray(UBound(UArray, 1), 1) = ErrorArray(ErrRow, col + x)
        End If
    End If

End Sub

Sub Interior_Color(fsheet)
    Dim FPS As Variant
    Dim Frow As Integer
    '------------------------------------------

    '------------------------------------------
    fsheet.Select
    'reset worksheet to completely white and center alligns all columns
    FPS = ActiveSheet.UsedRange
    For Frow = LBound(FPS) To UBound(FPS, 1) Step 1
        Range(Cells(Frow, 1), Cells(Frow, UBound(FPS, 2))).Interior.ColorIndex = 0
    Next Frow
    Range(Cells(Row, 2), Cells(UBound(FPS, 1), UBound(FPS, 2))).HorizontalAlignment = xlCenter

    'every cell in every second row is made light blue
    FPS = Cells(Row, 1).CurrentRegion
    For Frow = Row + 1 To UBound(FPS, 1) + Row - 1 Step 2
        Range(Cells(Frow, 1), Cells(Frow, UBound(FPS, 2))).Interior.Color = RGB(200, 230, 255)
    Next Frow
    'every cell in every second row, starting one later, is made white
    For Frow = Row To UBound(FPS, 1) + Row - 1 Step 2
        Range(Cells(Frow, 1), Cells(Frow, UBound(FPS, 2))).Interior.Color = RGB(255, 255, 255)
    Next Frow
End Sub

Sub Title_Maintain(fsheet)
    'Makes sure the titles for the simulated values is maintained after a macro runs
    fsheet.Select
    Range(Cells(1, 1), Cells(1, 3)).Interior.Color = RGB(255, 255, 255)
    Range(Cells(1, 1), Cells(1, 3)).Merge
End Sub

Sub Stats_Section_Maintain(fsheet)
    Dim x As Integer, StatEnd As Integer
    'makes sure the table containing the user inputs has a white background
    StatEnd = 14 'where the last row of the table is
    fsheet.Select
    For x = StatRow To StatRow + StatEnd
        Range(Cells(x, 1), Cells(x, 2)).Interior.Color = RGB(255, 255, 255)
    Next x
End Sub

Sub SimulationErrorMaintain()
    Dim ErrSection As Variant, ColStart As Integer, RowStart As Integer
    Dim SimErr As Worksheet, x As Integer


    Set SimErr = Sheets("Simulation Error")
    SimErr.Select
    ErrSection = Cells(1, 1).CurrentRegion

    'Allign, Centre, Border, and Colour the table on the left side of the worksheet (ErrSection = ErrorArray)
    Range(Cells(LBound(ErrSection, 1), LBound(ErrSection, 2)), Cells(UBound(ErrSection, 1), UBound(ErrSection, 2))) _
    .HorizontalAlignment = xlCenter
    Range(Cells(LBound(ErrSection, 1), LBound(ErrSection, 2)), Cells(UBound(ErrSection, 1), UBound(ErrSection, 2))) _
    .BorderAround ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    Range(Cells(LBound(ErrSection, 1), LBound(ErrSection, 2)), Cells(LBound(ErrSection, 1), UBound(ErrSection, 2))) _
    .BorderAround ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    Range(Cells(LBound(ErrSection, 1), LBound(ErrSection, 2)), Cells(LBound(ErrSection, 1), UBound(ErrSection, 2))) _
    .Interior.Color = RGB(200, 230, 255)

    'the three tables next to ErrSection are bordered, alligned and the title is coloured
    For x = UBound(ErrSection, 2) + 2 To UBound(ErrSection, 2) + 2 + 4 * 2 Step 4
        ErrSection = Cells(1, x).CurrentRegion
        Cells(1, x).Interior.Color = RGB(200, 230, 255)
        Range(Cells(1, x), Cells(UBound(ErrSection, 1), x + UBound(ErrSection, 2) - 1)) _
        .BorderAround ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
        Range(Cells(1 + 1, x), Cells(UBound(ErrSection, 1), x + UBound(ErrSection, 2) - 1)) _
        .BorderAround ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
        Range(Cells(1, x + 1), Cells(UBound(ErrSection, 1), x + UBound(ErrSection, 2) - 1)) _
        .BorderAround ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
        Range(Cells(1, x + 1), Cells(UBound(ErrSection, 1), x + UBound(ErrSection, 2) - 1)) _
        .HorizontalAlignment = xlCenter
    Next x

    'the section showing the errors between the simulation and actual results, it is created _
    'relative to ErrSection
    ErrSection = Cells(1, 1).CurrentRegion
    ColStart = UBound(ErrSection, 2) + 2
    ErrSection = Cells(1, ColStart).CurrentRegion
    RowStart = UBound(ErrSection, 1) + 2
    ErrSection = Cells(8, UBound(ErrSection, 2) + 2).CurrentRegion

    Range(Cells(RowStart + 1, ColStart), Cells(RowStart + UBound(ErrSection, 1) - 1, ColStart + UBound(ErrSection, 2) - 1)) _
    .BorderAround ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    Range(Cells(RowStart, ColStart + 1), Cells(RowStart + UBound(ErrSection, 1) - 1, ColStart + UBound(ErrSection, 2) - 1)) _
    .BorderAround ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    Range(Cells(RowStart, ColStart + 1), Cells(RowStart + UBound(ErrSection, 1) - 1, ColStart + UBound(ErrSection, 2) - 1)) _
    .HorizontalAlignment = xlCenter

    SimErr.Cells.EntireColumn.AutoFit
End Sub
