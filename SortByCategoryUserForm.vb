Option Explicit On
'Locations for Tables on the Fantasy Points Simulator sheet
Const Row As Integer = 20
Const StatRow As Integer = 3
Public MinGames As Integer

Private Sub OB1_Click()
    Dim SelectCat As String

    SelectCat = OB1.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB2_Click()
    Dim SelectCat As String

    SelectCat = OB2.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB3_Click()
    Dim SelectCat As String

    SelectCat = OB3.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB4_Click()
    Dim SelectCat As String

    SelectCat = OB4.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB5_Click()
    Dim SelectCat As String

    SelectCat = OB5.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB6_Click()
    Dim SelectCat As String

    SelectCat = OB6.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB7_Click()
    Dim SelectCat As String

    SelectCat = OB7.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB8_Click()
    Dim SelectCat As String

    SelectCat = OB8.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB9_Click()
    Dim SelectCat As String

    SelectCat = OB9.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB10_Click()
    Dim SelectCat As String

    SelectCat = OB10.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB11_Click()
    Dim SelectCat As String

    SelectCat = OB11.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB12_Click()
    Dim SelectCat As String

    SelectCat = OB12.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB13_Click()
    Dim SelectCat As String

    SelectCat = OB13.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB14_Click()
    Dim SelectCat As String

    SelectCat = OB14.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB15_Click()
    Dim SelectCat As String

    SelectCat = OB15.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB16_Click()
    Dim SelectCat As String

    SelectCat = OB16.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB17_Click()
    Dim SelectCat As String

    SelectCat = OB17.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB18_Click()
    Dim SelectCat As String

    SelectCat = OB18.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub OB19_Click()
    Dim SelectCat As String
    'Option button corrosponding to a category
    SelectCat = OB19.Caption
    Call General_Sorting(SelectCat)
End Sub

Private Sub SpinButton1_Change()
    'this is solely for the FP per Game category
    MinGames = SpinButton1.Value
    Label3.Caption = MinGames
End Sub


Private Sub UserForm_Activate()
    'When userform is opened, immediately colours option box texts to show
    'sortable categories
    Dim Cat As Integer, FTable As Variant
    Dim OBnum As Integer
    'initial value of spin button = min number of games played
    SpinButton1.Value = 10
    MinGames = SpinButton1.Value
    Label3.Caption = MinGames

    FTable = Cells(Row, 1).CurrentRegion
    'First turn all the text to black
    For OBnum = 1 To 19
        Me.Controls("OB" & OBnum).ForeColor = RGB(0, 0, 0)
    Next OBnum

    'loop through each caption for the option boxes, if it exists in _
    'the table On the worksheet, then turn the font red for the caption,
    'If itThen doesnt exist, keep the font black
    For Cat = 3 To UBound(FTable, 2)
        For OBnum = 1 To 19 ' number of option boxes
            If Me.Controls("OB" & OBnum).Caption = FTable(2, Cat) Then
                Me.Controls("OB" & OBnum).ForeColor = RGB(255, 0, 0)
                If Me.Controls("OB18").ForeColor = RGB(255, 0, 0) Then
                    Me.Controls("OB19").ForeColor = RGB(255, 0, 0)
                End If
                Exit For
            End If
        Next OBnum
    Next Cat
    'where the startup position for the userform is
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + Application.Width - Me.Width - 25

End Sub



Sub General_Sorting(SelectCat)
    Dim FTable As Variant, fsheet As Worksheet
    Dim SortR As Integer, SortC As Integer, Cat As Integer

    Application.CutCopyMode = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set fsheet = Sheets("Fantasy Points Simulator")
    fsheet.Select
    FTable = Cells(Row, 1).CurrentRegion
    'Loop through each category, beginning with the third, in the table _
    'On the worksheet
    For Cat = 3 To UBound(FTable, 2)
        'if there are no parameters for to sort with(Ex. minimum number of games) then _
        'the category can be sorted with excel sort function
        If Cells(Row + 1, Cat).Value = SelectCat And SelectCat <> "FP Per Game" Then
            Range(Cells(Row + 1, 1), Cells(Row + UBound(FTable, 1) - 1, UBound(FTable, 2))).Sort key1:=Cells(Row + 1, Cat), order1:=xlDescending, Header:=xlYes
        ElseIf Cells(Row + 1, Cat).Value = SelectCat And SelectCat = "FP Per Game" Then
            'since a minimum number of games is required, the players who do not meet the criteria must be _
            'place at the bottom of the list
            SortC = Cat
            Call sort_by_FP_PG(SortC)
        End If
    Next Cat

    Call Interior_Color()
    Call ColourChangeResults()
    Call Stats_Section_Maintain()
    Call Title_Maintain()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Sub sort_by_FP_PG(SortC)
    'Sorts players based on Fantasy Points per a game
    Dim x As Integer, min As Integer
    Dim FTable As Variant, FPcol As Integer, Frow As Integer
    Dim VTable As Variant, Vrow As Integer, Vcol As Integer
    Dim Dtable As Variant, Drow As Integer, Dcol As Integer
    Dim name As String

    FTable = Cells(Row, 1).CurrentRegion
    min = MinGames
    FPcol = SortC

    'if the column for fantasy points per a game exists then the macro will proceed
    'sorts the range of values from largest value to smallest
    Range(Cells(Row + 1, 1), Cells(Row + UBound(FTable, 1) - 1, UBound(FTable, 2))).Sort key1:=Cells(Row + 2, SortC), order1:=xlDescending, Header:=xlYes

    FTable = Cells(Row, 1).CurrentRegion 'redefine array post-sorting
    'define a new array that will hold all the players that dont meet the minimum games played criteria
    ReDim VTable(1 To LBound(FTable), LBound(FTable) To UBound(FTable, 2))
    'The Ftable table array is static so that it can ensure each name is checked once
    'this for loop checks to see if the minimum number of games is meat for a player
    'if it is not, the player is and their data is put into the VTable array _
    'And Is deleted from the worksheet
    'before they are deleted, DTable is defined and used to find the players exact _
    'location on the worksheet
    For Frow = LBound(FTable) + 2 To UBound(FTable, 1)
        '*********************************************
        If FTable(Frow, 3) < min Then
            'a hockey player is found to have less than min number of games
            name = FTable(Frow, 1)
            Dtable = Cells(Row, 1).CurrentRegion
            '******************************************
            For Drow = LBound(Dtable) To UBound(Dtable, 1)
                'search through newly defined array for players exact location on worksheet
                If name = Dtable(Drow, 1) Then
                    'when found, VTable is expanded along x-axis by double transposing _
                    'to accommodate an additonal player
                    If IsEmpty(VTable(LBound(VTable), 1)) = False Then
                        VTable = Application.Transpose(VTable)
                        ReDim Preserve VTable(1 To UBound(VTable, 1), 1 To UBound(VTable, 2) + 1)
                        VTable = Application.Transpose(VTable)
                    End If
                    'hockey player data is put into the expanded array
                    For Dcol = LBound(Dtable) To UBound(Dtable, 2)
                        VTable(UBound(VTable, 1), Dcol) = Dtable(Drow, Dcol)
                    Next Dcol
                    Cells(Drow + Row - 1, 1).EntireRow.delete
                    Exit For
                End If
            Next Drow
        End If
    Next Frow
    '*************************************************************************
    'this section places all the players, and their data, at the bottom of the worksheet
    Dtable = Cells(Row, 1).CurrentRegion
    For Vrow = LBound(VTable) To UBound(VTable, 1)
        For Vcol = LBound(VTable) To UBound(VTable, 2)
            Cells(UBound(Dtable, 1) + Row + Vrow - 1, Vcol).Value = VTable(Vrow, Vcol)
        Next Vcol
    Next Vrow
    '***************************************************************************
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
    Range(Cells(1, 1), Cells(1, 3)).Interior.Color = RGB(255, 255, 255)
    Range(Cells(1, 1), Cells(1, 3)).Merge
End Sub