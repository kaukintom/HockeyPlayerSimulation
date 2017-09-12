Option Explicit On
'Locations for Tables on the Fantasy Points Simulator sheet
Const Row As Integer = 20
Const StatRow As Integer = 3

Sub Fantasy_Points()
    'Finds the fantasy points per a game and total fantasy points for each player
    'and displays the user selected categories
    Dim Latest As Integer, Season As Variant, Fpoints As Worksheet
    Dim start As Range, x As Integer, Cat As String, score As Double
    Dim points As Variant, srow As Integer, Scol As Integer, tp As Double
    Dim Y As Integer, Prow As Integer, ls As Worksheet
    Dim Col As Integer
    Sheets("Latest Season Data").Visible = True
    Application.CutCopyMode = False
    Col = 4
    Latest = 2015
        
    Set ls = Sheets("Latest Season Data")
    Set start = Cells(Row, 1)
    Set Fpoints = Sheets("Fantasy Points simulator")


    Call clear()
    Application.CutCopyMode = False
    'copy the latest season season data to the Fantasy Simulator Page
    ls.Select
    Season = ActiveSheet.UsedRange
    ActiveSheet.Range(Cells(1, 1), Cells(UBound(Season, 1), LBound(Season, 2) + 2)).Copy
    Fpoints.Select
    Cells(Row, 1).Select
    Range(Cells(Row, 1), Cells(UBound(Season, 1), (LBound(Season, 2) + 2))).PasteSpecial

    'Determines which columns have the data the user wants to view
    For x = 2 To 15
        If Cells(x, 2).Value <> 0 Then
            Cat = Cells(x, 1).Value
            For Scol = 4 To 20
                If Season(2, Scol) = Cat Then
                    ls.Select
                    Range(Cells(2, Scol), Cells(UBound(Season), Scol)).Copy
                    Fpoints.Select
                    Cells(Row + 1, Col).Select
                    Range(Cells(Row + 1, Col), Cells(UBound(Season) + Row, Col)).PasteSpecial
                    Col = Col + 1
                    Exit For
                End If
            Next Scol
        End If
    Next x

    Cells(Row + 1, Col).Value = "Total FP"
    Cells(Row + 1, Col + 1).Value = "FP Per Game"

    'create an array for the fantasy point multipliers and the regular season stats
    Cells(Row, 1).Select
    Season = ActiveCell.CurrentRegion
    points = Cells(StatRow, 1).CurrentRegion

    'multiply the stats in an array
    For Scol = (LBound(Season) + 3) To (UBound(Season, 2) - 2)
        Cat = Season(2, Scol)
        For Prow = LBound(points) To UBound(points, 1)
            If points(Prow, 1) = Cat Then score = points(Prow, 2)
        Next Prow
        For srow = (LBound(Season) + 2) To UBound(Season, 1)
            Season(srow, Scol) = Season(srow, Scol) * score
        Next srow
    Next Scol

    Cells(Row, 1).Select

    ' add up the values from each stat to determine points
    For srow = (LBound(Season) + 2) To UBound(Season, 1)
        tp = 0
        For Scol = 4 To (UBound(Season, 2) - 2)
            tp = tp + Season(srow, Scol)
        Next Scol
        Season(srow, (UBound(Season, 2) - 1)) = tp
        If Season(srow, 3) > 0 Then
            Season(srow, UBound(Season, 2)) = Round(tp / (Season(srow, 3)), 3)
        Else
            Season(srow, UBound(Season, 2)) = tp
        End If
    Next srow

    'Place the Total Fantasy points and Fantasy points Per a Game in the Correct Columns (Last 2 Columns)
    For srow = (LBound(Season) + 1) To UBound(Season, 1)
        Cells(srow + Row - 1, UBound(Season, 2) - 1).Value = Season(srow, UBound(Season, 2) - 1)
        Cells(srow + Row - 1, UBound(Season, 2)).Value = Season(srow, UBound(Season, 2))
    Next srow
    Call Interior_Color()
    Call Stats_Section_Maintain()
    Call Title_Maintain()
    Application.CutCopyMode = False
    Sheets("Latest Season Data").Visible = False
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

Sub Stats_Section_Maintain()
    Dim x As Integer, StatEnd As Integer
    'makes sure the table containing the user inputs has a white background
    StatEnd = 14 'where the last row of the table is

    For x = StatRow To StatRow + StatEnd
        Range(Cells(x, 1), Cells(x, 2)).Interior.Color = RGB(255, 255, 255)
    Next x
End Sub

Sub Title_Maintain()
    'Main title on worksheet
    Range(Cells(1, 1), Cells(1, 3)).Interior.Color = RGB(255, 255, 255)
    Range(Cells(1, 1), Cells(1, 3)).Merge
    'Season Title
    Cells(Row, 1).Font.Size = 12
    Cells(Row, 1).Font.Bold = True
End Sub

Sub clear()
    Cells(Row, 1).Select
    ActiveCell.CurrentRegion.delete
    Sheets("Simulations(A)").UsedRange.delete
    Sheets("Simulations(G)").UsedRange.delete
End Sub
