Option Explicit On

Const Row As Integer = 20
Const StatRow As Integer = 3
Public SheetName As String

Private Sub UserForm_Activate()
    Dim fsheet As Worksheet

    SheetName = ActiveSheet.Name
    Set fsheet = Sheets(ActiveSheet.Name)
    Sheets("Simulation Time").Select
    Cells(1, 1).Select
    fsheet.Select
    'UserForm location at top right of page
    SpinButton1.Value = 0
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + Application.Width - Me.Width - 25
    'Hide objects on current page
    fsheet.Shapes("TextBox 15").Visible = False
    fsheet.Shapes("TextBox 18").Visible = False
    fsheet.Shapes("TextBox 5").Visible = False
    fsheet.Shapes("TextBox 6").Visible = False
    fsheet.Shapes("TextBox 17").Visible = False
    fsheet.Shapes("TextBox 16").Visible = False

End Sub


Private Sub SpinButton1_Change()
    Dim FTable As Variant, fsheet As Worksheet, x As Integer
    Dim SimErr As Worksheet, ErrArray As Variant
    Dim psheet As Worksheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set fsheet = Sheets(SheetName)
    Set SimErr = Sheets("Simulation Error")
    Set psheet = Sheets("Prior Season Sim")
    FTable = fsheet.Cells(Row, 1).CurrentRegion

    If SpinButton1.Value = 0 Then
        Range(Cells(StatRow + 1, 2), Cells(StatRow + 14, 2)).Interior.Color = RGB(255, 255, 255)
        '**************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        ActiveSheet.Shapes("Right Arrow 10").Visible = False
        ActiveSheet.Shapes("Right Arrow 27").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 29").Visible = False
        ActiveSheet.Shapes("Right Arrow 12").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 30").Visible = False
        ActiveSheet.Shapes("TextBox 32").Visible = False
        ActiveSheet.Shapes("Right Arrow 35").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 33").Visible = False
        ActiveSheet.Shapes("Right Arrow 36").Visible = False
        Range(Cells(Row + 1, 1), Cells(Row + 1, UBound(FTable, 2))).Interior.Color = RGB(200, 230, 255)
        '**************************************************
        Sheets("Simulation Time").Shapes("TextBox 6").Visible = False
        fsheet.Select
    ElseIf SpinButton1.Value = 1 Then
        '**************************************************
        Range(Cells(StatRow + 1, 2), Cells(StatRow + 14, 2)).Interior.Color = RGB(255, 255, 0)
        ActiveSheet.Shapes("TextBox 3").Visible = True
        ActiveSheet.Shapes("Right Arrow 10").Visible = True
        ActiveSheet.Shapes("Right Arrow 27").Visible = True
        '**************************************************
        ActiveSheet.Shapes("TextBox 29").Visible = False
        ActiveSheet.Shapes("Right Arrow 12").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 30").Visible = False
        ActiveSheet.Shapes("TextBox 32").Visible = False
        ActiveSheet.Shapes("Right Arrow 35").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 33").Visible = False
        ActiveSheet.Shapes("Right Arrow 36").Visible = False
        Range(Cells(Row + 1, 1), Cells(Row + 1, UBound(FTable, 2))).Interior.Color = RGB(200, 230, 255)
        '**************************************************
        Sheets("Simulation Time").Shapes("TextBox 6").Visible = False
        fsheet.Select
    ElseIf SpinButton1.Value = 2 Then
        Range(Cells(StatRow + 1, 2), Cells(StatRow + 14, 2)).Interior.Color = RGB(255, 255, 255)
        '**************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        ActiveSheet.Shapes("Right Arrow 10").Visible = False
        ActiveSheet.Shapes("Right Arrow 27").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 29").Visible = True
        ActiveSheet.Shapes("Right Arrow 12").Visible = True
        '**************************************************
        ActiveSheet.Shapes("TextBox 30").Visible = False
        ActiveSheet.Shapes("TextBox 32").Visible = False
        ActiveSheet.Shapes("Right Arrow 35").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 33").Visible = False
        ActiveSheet.Shapes("Right Arrow 36").Visible = False
        Range(Cells(Row + 1, 1), Cells(Row + 1, UBound(FTable, 2))).Interior.Color = RGB(200, 230, 255)
        '**************************************************
        Sheets("Simulation Time").Shapes("TextBox 6").Visible = False
        fsheet.Select
    ElseIf SpinButton1.Value = 3 Then
        '**************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        ActiveSheet.Shapes("Right Arrow 10").Visible = False
        ActiveSheet.Shapes("Right Arrow 27").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 29").Visible = False
        ActiveSheet.Shapes("Right Arrow 12").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 30").Visible = True
        ActiveSheet.Shapes("TextBox 32").Visible = True
        ActiveSheet.Shapes("Right Arrow 35").Visible = True
        '**************************************************
        ActiveSheet.Shapes("TextBox 33").Visible = False
        ActiveSheet.Shapes("Right Arrow 36").Visible = False
        Range(Cells(Row + 1, 3), Cells(Row + 1, UBound(FTable, 2))).Interior.Color = RGB(200, 230, 255)
        '**************************************************
        Sheets("Simulation Time").Shapes("TextBox 6").Visible = False
        fsheet.Select
    ElseIf SpinButton1.Value = 4 Then
        fsheet.Select
        Range(Cells(Row + 1, 3), Cells(Row + 1, UBound(FTable, 2))).Interior.Color = RGB(255, 255, 0)
        '**************************************************
        fsheet.Shapes("TextBox 3").Visible = False
        fsheet.Shapes("Right Arrow 10").Visible = False
        fsheet.Shapes("Right Arrow 27").Visible = False
        '**************************************************
        fsheet.Shapes("TextBox 29").Visible = False
        fsheet.Shapes("Right Arrow 12").Visible = False
        '**************************************************
        fsheet.Shapes("TextBox 30").Visible = False
        fsheet.Shapes("TextBox 32").Visible = False
        fsheet.Shapes("Right Arrow 35").Visible = False
        '**************************************************
        fsheet.Shapes("TextBox 33").Visible = True
        fsheet.Shapes("Right Arrow 36").Visible = True
        '**************************************************
        Sheets("Simulation Time").Shapes("TextBox 6").Visible = False
        fsheet.Select
    ElseIf SpinButton1.Value = 5 Then
        psheet.Select
        ActiveSheet.Shapes("TextBox 8").Visible = False
        '**************************************************
        fsheet.Select
        '**************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        ActiveSheet.Shapes("Right Arrow 10").Visible = False
        ActiveSheet.Shapes("Right Arrow 27").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 29").Visible = False
        ActiveSheet.Shapes("Right Arrow 12").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 30").Visible = False
        ActiveSheet.Shapes("TextBox 32").Visible = False
        ActiveSheet.Shapes("Right Arrow 35").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 33").Visible = False
        ActiveSheet.Shapes("Right Arrow 36").Visible = False
        Range(Cells(Row + 1, LBound(FTable, 2)), Cells(Row + 1, UBound(FTable, 2))).Interior.Color = RGB(200, 230, 255)
        '**************************************************
        Sheets("Simulation Time").Select
        Cells(1, 1).Select
        ActiveSheet.Shapes("TextBox 6").Visible = True
    ElseIf SpinButton1.Value = 6 Then
        fsheet.Select
        '**************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        ActiveSheet.Shapes("Right Arrow 10").Visible = False
        ActiveSheet.Shapes("Right Arrow 27").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 29").Visible = False
        ActiveSheet.Shapes("Right Arrow 12").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 30").Visible = False
        ActiveSheet.Shapes("TextBox 32").Visible = False
        ActiveSheet.Shapes("Right Arrow 35").Visible = False
        '**************************************************
        ActiveSheet.Shapes("TextBox 33").Visible = False
        ActiveSheet.Shapes("Right Arrow 36").Visible = False
        Range(Cells(Row + 1, LBound(FTable, 2)), Cells(Row + 1, UBound(FTable, 2))).Interior.Color = RGB(200, 230, 255)
        '**************************************************
        Sheets("Simulation Time").Select
        ActiveSheet.Shapes("TextBox 6").Visible = False
        '***************************************************
        SimErr.Select
        ErrArray = Cells(1, 1).CurrentRegion
        ActiveSheet.Shapes("TextBox 1").Visible = False

        If ActiveSheet.Shapes("TextBox 1").Visible = True Then
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 255)
        End If
        '*****************************************************
        ActiveSheet.Shapes("TextBox 2").Visible = False
        If ActiveSheet.Shapes("TextBox 2").Visible = True Then
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 0)
            Next x
        Else
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 255)
            Next x
        End If
        '******************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        If ActiveSheet.Shapes("TextBox 3").Visible = True Then
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 255)
        End If
        '**************************************************
        psheet.Select
        CenterOnCell Range("A1")
        ActiveSheet.Shapes("TextBox 10").Visible = False
        ActiveSheet.Shapes("TextBox 11").Visible = False
        ActiveSheet.Shapes("TextBox 13").Visible = False
        ActiveSheet.Shapes("TextBox 14").Visible = False
        ActiveSheet.Shapes("TextBox 8").Visible = True

    ElseIf SpinButton1.Value = 7 Then
        psheet.Select

        ActiveSheet.Shapes("TextBox 8").Visible = False
        '***************************************************
        SimErr.Select
        CenterOnCell Range("A1")
        ErrArray = Cells(1, 1).CurrentRegion
        ActiveSheet.Shapes("TextBox 1").Visible = True

        If ActiveSheet.Shapes("TextBox 1").Visible = True Then
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 255)
        End If
        '*****************************************************
        ActiveSheet.Shapes("TextBox 2").Visible = False
        If ActiveSheet.Shapes("TextBox 2").Visible = True Then
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 0)
            Next x
        Else
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 255)
            Next x
        End If
        '******************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        If ActiveSheet.Shapes("TextBox 3").Visible = True Then
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 255)
        End If
    ElseIf SpinButton1.Value = 8 Then
        SimErr.Select
        ErrArray = Cells(1, 1).CurrentRegion
        ActiveSheet.Shapes("TextBox 1").Visible = False
        CenterOnCell Range("A1")
        If ActiveSheet.Shapes("TextBox 1").Visible = True Then
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 255)
        End If
        '*****************************************************

        ActiveSheet.Shapes("TextBox 2").Visible = True
        If ActiveSheet.Shapes("TextBox 2").Visible = True Then
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 0)
            Next x
        Else
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 255)
            Next x
        End If
        '******************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        If ActiveSheet.Shapes("TextBox 3").Visible = True Then
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 255)
        End If
    ElseIf SpinButton1.Value = 9 Then
        SimErr.Select
        ErrArray = Cells(1, 1).CurrentRegion
        ActiveSheet.Shapes("TextBox 1").Visible = False
        CenterOnCell Range("H1")
        If ActiveSheet.Shapes("TextBox 1").Visible = True Then
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 255)
        End If
        '*****************************************************
        ActiveSheet.Shapes("TextBox 2").Visible = False
        If ActiveSheet.Shapes("TextBox 3").Visible = True Then
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 0)
            Next x
        Else
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 255)
            Next x
        End If
        '******************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = True
        If ActiveSheet.Shapes("TextBox 3").Visible = True Then
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 255)
        End If
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim FTable As Variant, fsheet As Worksheet, x As Integer
    Dim SimErr As Worksheet, ErrArray As Variant, psheet As Worksheet

    Set psheet = Sheets("Prior Season Sim")
    Set fsheet = Sheets(SheetName)
    Set SimErr = Sheets("Simulation Error")
    FTable = Cells(Row, 1).CurrentRegion

    If CloseMode = 0 Then
        psheet.Select
        ActiveSheet.Shapes("TextBox 8").Visible = False
        ActiveSheet.Shapes("TextBox 10").Visible = True
        ActiveSheet.Shapes("TextBox 11").Visible = True
        ActiveSheet.Shapes("TextBox 13").Visible = True
        ActiveSheet.Shapes("TextBox 14").Visible = True
        SimErr.Select
        CenterOnCell Range("A1")
        ErrArray = Cells(1, 1).CurrentRegion
        ActiveSheet.Shapes("TextBox 1").Visible = False

        If ActiveSheet.Shapes("TextBox 1").Visible = True Then
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(LBound(ErrArray, 1) + 1, LBound(ErrArray, 2)), Cells(UBound(ErrArray, 1), UBound(ErrArray, 2))) _
            .Interior.Color = RGB(255, 255, 255)
        End If
        '*****************************************************
        ActiveSheet.Shapes("TextBox 2").Visible = False
        If ActiveSheet.Shapes("TextBox 2").Visible = True Then
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 0)
            Next x
        Else
            For x = UBound(ErrArray, 2) + 3 To UBound(ErrArray, 2) + 3 + 4 * 2 Step 4
                Range(Cells(2, x), Cells(5, x + 1)) _
                .Interior.Color = RGB(255, 255, 255)
            Next x
        End If
        '******************************************************
        ActiveSheet.Shapes("TextBox 3").Visible = False
        If ActiveSheet.Shapes("TextBox 3").Visible = True Then
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 0)
        Else
            Range(Cells(8, UBound(ErrArray, 2) + 3), Cells(10, UBound(ErrArray, 2) + 7)) _
            .Interior.Color = RGB(255, 255, 255)
        End If


        fsheet.Select
        Range(Cells(StatRow + 1, 2), Cells(StatRow + 14, 2)).Interior.Color = RGB(255, 255, 255)
        Sheets("Simulation Time").Shapes("TextBox 6").Visible = False
        '*************************************************
        fsheet.Select
        fsheet.Shapes("TextBox 15").Visible = True
        fsheet.Shapes("TextBox 18").Visible = True
        fsheet.Shapes("TextBox 5").Visible = True
        fsheet.Shapes("TextBox 6").Visible = True
        fsheet.Shapes("TextBox 17").Visible = True
        fsheet.Shapes("TextBox 16").Visible = True
        '**************************************************
        fsheet.Shapes("TextBox 37").Visible = False
        fsheet.Shapes("TextBox 3").Visible = False
        fsheet.Shapes("Right Arrow 10").Visible = False
        fsheet.Shapes("Right Arrow 27").Visible = False
        '**************************************************
        fsheet.Shapes("TextBox 29").Visible = False
        fsheet.Shapes("Right Arrow 12").Visible = False
        '**************************************************
        fsheet.Shapes("TextBox 30").Visible = False
        fsheet.Shapes("TextBox 32").Visible = False
        fsheet.Shapes("Right Arrow 35").Visible = False
        '**************************************************
        fsheet.Shapes("TextBox 33").Visible = False
        fsheet.Shapes("Right Arrow 36").Visible = False
        Range(Cells(Row + 1, 3), Cells(Row + 1, UBound(FTable, 2))).Interior.Color = RGB(200, 230, 255)
        fsheet.Select
    End If
End Sub

Sub CenterOnCell(OnCell As Range)
    'Centres the page on a specific cell
    Dim VisRows As Integer
    Dim VisCols As Integer

    Application.ScreenUpdating = False
    '
    ' Switch over to the OnCell's workbook and worksheet.
    '
    OnCell.Parent.Parent.Activate
    OnCell.Parent.Activate
    '
    ' Get the number of visible rows and columns for the active window.
    '
    With ActiveWindow.VisibleRange
        VisRows = .Rows.Count
        VisCols = .Columns.Count
    End With
    '
    ' Now, determine what cell we need to GOTO. The GOTO method will
    ' place that cell reference in the upper left corner of the screen,
    ' so that reference needs to be VisRows/2 above and VisCols/2 columns
    ' to the left of the cell we want to center on. Use the MAX function
    ' to ensure we're not trying to GOTO a cell in row <=0 or column <=0.
    '
    With Application
        .Goto reference:=OnCell.Parent.Cells(
            .WorksheetFunction.Max(1, OnCell.Row +
            (OnCell.Rows.Count / 2) - (VisRows / 2)),
            .WorksheetFunction.Max(1, OnCell.Column +
            (OnCell.Columns.Count / 2) -
            .WorksheetFunction.RoundDown((VisCols / 2), 0))),
         Scroll:=True
End With

    OnCell.Select
    Application.ScreenUpdating = True

End Sub
