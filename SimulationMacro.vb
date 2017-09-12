Option Explicit On
'Locations for tables on the Fantasy Points Simulator sheet
Const Row As Integer = 20
Const StatRow As Integer = 3
Public numsim As Integer 'number of simulations to conduct per a player


Sub Simulated_Values(Time)
    Dim LastCol As Integer, fsheet As Worksheet, Asheet As Worksheet, Gsheet As Worksheet, Lsheet As Worksheet
    Dim FPS As Variant, Scol As Integer
    Dim name As String, Sbegin As Integer, srow As Integer, GoalSim As Variant, AssistSim As Variant
    Dim Lrow As Integer, LastSeason As Variant, ADist As String, GDist As String
    Dim G As Integer, A As Integer, assist As Variant, Goal As Variant
    Dim Gmean As Double, Gsd As Double, Amean As Double, Asd As Double
    Dim Atot As Double, Gtot As Double, Acat As Integer, Gcat As Integer, games As Integer
    Dim Frow As Integer, GPG As Double, APG As Double, GperS As Integer
    Dim FPScol As Integer, FPSrow As Integer, PN As Double, Grnd As Double
    Dim year As Integer, ALoc As Integer, GLoc As Integer
    Dim StartTime As Double

    '*************************************************************
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    '*************************************************************
    StartTime = Timer

    GLoc = 3 'Worksheet location of simulated Goals relative to FPS Array
    ALoc = 2
    year = 2015
    Sbegin = 3 'where hockey players begin on Lsheet
    G = 21 'column for Goal data for Distributions
    A = 22 'column for Assist Data for distributions
    Gcat = 4 'Column for Goals
    Acat = 5 'column for assists
    games = 3 'column for games
    GperS = 82 'games played in hockey season
    '************************************************************
    'Setup up worksheets
    Set Lsheet = Sheets("Latest Season Data")
    Set fsheet = Sheets("Fantasy Points Simulator")
    Set Gsheet = Sheets("Simulations(G)")
    Set Asheet = Sheets("Simulations(A)")

    Cells(Row, 1).Select
    FPS = fsheet.Cells(Row, 1).CurrentRegion
    Application.CutCopyMode = False
    'copy the hockey player data to the simulation worksheets
    Call setup(Sbegin, Asheet, Gsheet, numsim, fsheet, FPS, year)

    fsheet.Select
    Range(Cells(Row + 2, 1), Cells(UBound(FPS, 1) + Row - 1, 1)).Copy
    Gsheet.Select
    Range(Cells(3, 1), Cells(UBound(FPS, 1), 1)).PasteSpecial
    Asheet.Select
    Range(Cells(3, 1), Cells(UBound(FPS, 1), 1)).PasteSpecial

    Application.CutCopyMode = False
    '*******************************************************************
    'Define the arrays to be used
    GoalSim = Gsheet.UsedRange
    AssistSim = Asheet.UsedRange
    LastSeason = Lsheet.UsedRange
    FPS = fsheet.Cells(Row, 1).CurrentRegion
    'loop through each name in GoalSim Array, this array is identical to AssistSim and FPS
    'for each name, simulation takes place and the values are added to an array (FPS) and then _
    'placed in the Fantasy Points Simulator worksheet
    For srow = LBound(GoalSim) + 2 To UBound(GoalSim, 1)
        name = GoalSim(srow, 1)
        ' find name and values to be used in probabilities
        For Lrow = LBound(LastSeason) To UBound(LastSeason, 1)
            If LastSeason(Lrow, 1) = name Then
                GDist = LastSeason(Lrow, G)
                ADist = LastSeason(Lrow, A)
                Exit For
            End If
        Next Lrow
        'Split arrays to find mean and variance for each stat
        If GDist <> "|" And GDist <> "" And ADist <> "|" And ADist <> "" Then
            assist = Split(ADist, "|")
            Goal = Split(GDist, "|")
            Gmean = Goal(LBound(Goal))
            Gsd = Goal(UBound(Goal))
            Amean = assist(LBound(assist))
            Asd = assist(UBound(assist))
            'the following variables help find Sum of total differences
            Gtot = 0
            Atot = 0
            'Conduct specified number of simulations for current player
            For Scol = 2 To numsim + 1
                Call Lookup(Gsd, Asd, Gmean, Amean, Gsheet, Asheet, GoalSim, AssistSim, Gtot, Atot, Scol, srow)
            Next Scol

            'find the average change in a stat category
            GoalSim(srow, Scol) = Gtot / numsim
            AssistSim(srow, Scol) = Atot / numsim
            'now the previous seasons stat category, is divided by the number of games _
            'so that the stat can be compared to simulated values
            For Frow = LBound(LastSeason) To UBound(LastSeason, 1)
                If LastSeason(Frow, 1) = name Then
                    GPG = LastSeason(Frow, Gcat) / LastSeason(Frow, games)
                    APG = LastSeason(Frow, Acat) / LastSeason(Frow, games)
                    Exit For
                End If
            Next Frow
            'Add the simulated averages to the previous season stats per a game to get the simulated _
            'New season averages
            'these values are put into an array and will be displayed on the worksheet later
            '**********************************************************************************
            'Simulated values are put into the aray here, using the locations provided _
            'by ALoc And GLoc
            'Simulation could produce negative point totals, which is impossible and must be corrected
            'to zero if they occur
            If CInt((GPG + GoalSim(srow, Scol)) * GperS) < 0 Then
                FPS(srow, UBound(FPS, 2) - GLoc) = 0
            Else
                FPS(srow, UBound(FPS, 2) - GLoc) = CInt((GPG + GoalSim(srow, Scol)) * GperS)
            End If

            If CInt((APG + AssistSim(srow, Scol)) * GperS) < 0 Then
                FPS(srow, UBound(FPS, 2) - ALoc) = 0
            Else
                FPS(srow, UBound(FPS, 2) - ALoc) = CInt((APG + AssistSim(srow, Scol)) * GperS)
            End If
            FPS(srow, UBound(FPS, 2) - ALoc + 1) = FPS(srow, UBound(FPS, 2) - ALoc) + FPS(srow, UBound(FPS, 2) - GLoc)
            '***************************************************************************************
        End If
    Next srow

    'The Arrays containing all the simulated values is put into the worksheet
    For FPSrow = LBound(FPS) To UBound(FPS, 1)
        For FPScol = LBound(FPS) To UBound(FPS, 2)
            fsheet.Cells(FPSrow + Row - 1, FPScol).Value = FPS(FPSrow, FPScol)
        Next FPScol
    Next FPSrow
    'The Arrays conataining the simulated values are put in their respective worksheets
    Gsheet.Select
    Range(Cells(LBound(GoalSim), LBound(GoalSim)), Cells(UBound(GoalSim, 1), UBound(GoalSim, 2))) = GoalSim
    Asheet.Select
    Range(Cells(LBound(AssistSim), LBound(AssistSim)), Cells(UBound(AssistSim, 1), UBound(AssistSim, 2))) = AssistSim
    '***********************************************
    'worksheet maintenance and visual macros
    Call New_FP(FPS, fsheet, GLoc, ALoc)
    Call ColourChangeResults()
    Call Interior_Color()
    Call SimulatedColours()
    Call Stats_Section_Maintain()
    Call Title_Maintain()
    Call SecondsCount(srow, numsim, StartTime, fsheet, Time)
    '**************************************************
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub Lookup(Gsd, Asd, Gmean, Amean, Gsheet, Asheet, GoalSim, AssistSim, Gtot, Atot, Scol, srow)
    'This Macro conducts the simulation for each player using random numbers _
    'with the NormInv function

    Randomize
    GoalSim(srow, Scol) = Application.WorksheetFunction.NormInv(Rnd(), Gmean, Gsd)
    Gtot = Gtot + GoalSim(srow, Scol)

    Randomize
    AssistSim(srow, Scol) = Application.WorksheetFunction.NormInv(Rnd(), Amean, Asd)
    Atot = Atot + AssistSim(srow, Scol)

End Sub