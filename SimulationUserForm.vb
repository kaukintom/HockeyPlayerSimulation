Option Explicit On
'Locations for Tables on the Fantasy Points Simulator sheet
Public numsim As Integer
Const Row As Integer = 20
Const StatRow As Integer = 3

Private Sub UserForm_Activate()
    Unload Sorting
    '**************************************************************
    'UserForm Location
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + Application.Width - Me.Width - 25
    '**************************************************************
    ScrollBar1.Value = 50
    Simulation.Label2.Caption = ScrollBar1.Value
    numsim = ScrollBar1.Value
End Sub


Private Sub ScrollBar1_Change()
    'when number of simulations changes, show the changes in a textbox
    'to the user and adjust the variable containing number of simulations
    numsim = ScrollBar1.Value
    Simulation.Label2.Caption = numsim
End Sub

Private Sub ToggleButton1_Click()
    Dim Time As Double
    Call Simulated_Values(Time) 'macro that conducts simulation
    'then the time it took to run the simulation macro is displayed
    Simulation.Label3.Caption = numsim & " simulations took " & Time & " seconds to complete."
End Sub

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

    Call New_FP(FPS, fsheet, GLoc, ALoc)
    Call ColourChangeResults()
    Call Interior_Color()
    Call SimulatedColours()
    Call Stats_Section_Maintain()
    Call Title_Maintain()

    '**************************************************
    Call SecondsCount(srow, numsim, StartTime, fsheet, Time)

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub setup(Sbegin, Asheet, Gsheet, numsim, fsheet, FPS, year)
    Dim x As Integer, numFP As Integer
    'setup goal simulation sheet
    Gsheet.Select
    Gsheet.UsedRange.delete
    Cells(Sbegin - 2, 1).Value = "Goals"
    Cells(Sbegin - 1, 1).Value = "Player"
    For x = 1 To numsim
        Cells(Sbegin - 1, x + 1).Value = x
    Next x
    Cells(Sbegin - 1, numsim + 2).Value = "Average Goals"
    'setup assist simulation sheet
    Asheet.Select
    Asheet.UsedRange.delete
    Cells(Sbegin - 2, 1).Value = "Assists"
    Cells(Sbegin - 1, 1).Value = "Player"
    For x = 1 To numsim
        Cells(Sbegin - 1, x + 1).Value = x
    Next x
    Cells(Sbegin - 1, numsim + 2).Value = "Average Assists"
    'adds three more columns to Fantasy Points Simulator Worksheet
    For x = 1 To 100
        If fsheet.Cells(Row + 1, x).Value = "Goals" Then
            Exit For
        ElseIf IsEmpty(fsheet.Cells(Row + 1, x).Value) = True Then
            fsheet.Cells(Row + 1, x).Value = "Goals"
            Exit For
        End If
    Next x
    fsheet.Select
    Range(Cells(Row + 1, x), Cells(Row + UBound(FPS, 1) - 1, x + 2)).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    For x = 1 To 100
        If fsheet.Cells(Row + 1, x).Value = "Assists" Then
            Exit For
        ElseIf IsEmpty(fsheet.Cells(Row + 1, x).Value) = True Then
            fsheet.Cells(Row + 1, x).Value = "Assists"
            Exit For
        End If
    Next x
    'Area Containing the headings for the simulated values
    '***********************************************************************************
    fsheet.Cells(Row, x).Value = (year + 1) & "/" & (year + 2) & " Simulated Results"
    Range(Cells(Row, x - 1), Cells(Row, x + 2)).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
    Range(Cells(Row, x - 1), Cells(Row, x + 2)).Merge
    '***********************************************************************************
    For x = 1 To 100
        If fsheet.Cells(Row + 1, x).Value = "Points" Then
            Exit For
        ElseIf IsEmpty(fsheet.Cells(Row + 1, x).Value) = True Then
            fsheet.Cells(Row + 1, x).Value = "Points"
            Exit For
        End If
    Next x

    numFP = 0
    For x = 1 To 100
        If fsheet.Cells(Row + 1, x).Value = "Total FP" Then
            numFP = numFP + 1
            If numFP = 2 Then Exit For
        ElseIf IsEmpty(fsheet.Cells(Row + 1, x).Value) = True Then
            fsheet.Cells(Row + 1, x).Value = "Total FP"
            Exit For
        End If
    Next x
    Range(Cells(Row + 1, x), Cells(Row - 1 + UBound(FPS, 1), x)).BorderAround _
    ColorIndex:=xlColorIndexAutomatic, Weight:=xlMedium
End Sub

Sub Lookup(Gsd, Asd, Gmean, Amean, Gsheet, Asheet, GoalSim, AssistSim, Gtot, Atot, Scol, srow)
    'This Macro conducts the simulation for each player using random numbers _
    'with the NormInv function
    Dim x As Double, Random As Double

    Randomize
    GoalSim(srow, Scol) = Application.WorksheetFunction.NormInv(Rnd(), Gmean, Gsd)
    Gtot = Gtot + GoalSim(srow, Scol)

    Randomize
    AssistSim(srow, Scol) = Application.WorksheetFunction.NormInv(Rnd(), Amean, Asd)
    Atot = Atot + AssistSim(srow, Scol)

End Sub

Sub Interior_Color()
    'Creates the alternating blue and white lines for the fantasy points simulator page
    Dim FPS As Variant
    Dim Frow As Integer
    '------------------------------------------
    Sheets("Fantasy Points Simulator").Select
    '-------------------------------------------

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

Sub ColourChangeResults()
    'Changes the colour of the font of the simulated goals, assists, and points to red, 
    'green or black based on simulated values
    Dim FPS As Variant, x As Integer, Frow As Integer
    Dim G As Integer, NSG As Integer, fsheet As Worksheet
    Dim Goals As Integer, Assists As Integer
    Dim games As Integer
        
    Set fsheet = Sheets("Fantasy Points Simulator")
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
    'Can change which column the time will go
    'The time it takes the user to conduct simulation would be Utime
    Utime = 3
    Mtime = 2
    T = Mtime

    Sheets("Simulation Time").Select
    Cells(1, 2).Select
    For srow = 1 To 101
        If IsEmpty(Cells(srow, 1)) = True Then
            Time = Round(Timer - StartTime, 2)
            Cells(srow, 1).Value = numsim
            Cells(srow, T).Value = Time
            Exit For
        ElseIf Cells(srow, 1).Value = numsim Then
            Time = Round(Timer - StartTime, 2)
            Cells(srow, T).Value = Time
            Exit For
        End If
    Next srow
    fsheet.Select
End Sub

Sub SimulatedColours()
    'For the worksheets containing the simulated values, this macro changes the colour of the
    'rows of the table, and adds borders around sections
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

    'add a border around the column containing the averages of the simulations
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

    Sheets("Fantasy Points Simulator").Select
End Sub

Sub Stats_Section_Maintain()
    Dim x As Integer, StatEnd As Integer
    'makes sure the table containing the user inputs has a white background
    StatEnd = 14 'where the last row of the table is

    For x = StatRow To StatRow + StatEnd
        Range(Cells(x, 1), Cells(x, 2)).Interior.Color = RGB(255, 255, 255)
    Next x
End Sub

Sub Title_Maintain()
    'Makes sure the titles for the simulated values is maintained after a macro runs
    Range(Cells(1, 1), Cells(1, 3)).Interior.Color = RGB(255, 255, 255)
    Range(Cells(1, 1), Cells(1, 3)).Merge
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
