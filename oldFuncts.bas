Attribute VB_Name = "oldFuncts"
Public r As Object
Private Const POS_PLR_ADD As String = "b1:b250"
Private Const HEADER_ROW As Long = 3


Private Sub gotoNotesbak()
On Error GoTo errorHandler
    Dim playerName As String
    Dim pos As String
    Dim playerRow As Integer
    Dim notesCol As Integer
    Dim posName As String
    
    playerName = ActiveCell.Value
    If TypeName(ActiveCell.Offset(0, -1).Value) = "String" Then
        pos = ActiveCell.Offset(0, -1).Value
    Else
        posName = ActiveCell.Offset(-1 * ActiveCell.Offset(0, -2) - 1, -3)
        Select Case posName
            Case " QUARTERBACKS"
                pos = "QB"
            Case " RUNNING BACKS"
                pos = "RB"
            Case " WIDE RECEIVERS"
                pos = "WR"
            Case " KICKERS"
                pos = "K"
            Case " TIGHT ENDS"
                pos = "TE"
            Case " DEFENSE / SPECIAL TEAMS"
                pos = "DEF"
            Case Else
                GoTo errorHandler
                pos = "DUNNO"
        End Select
    End If
    Worksheets(pos).Activate
    playerRow = Application.WorksheetFunction.Match(playerName, Range(POS_PLR_ADD), 0)
    notesCol = Application.WorksheetFunction.Match("User Notes", Rows(HEADER_ROW), 0)
    Rows(playerRow).Select
    ActiveWindow.ScrollRow = playerRow
    ActiveWindow.ScrollColumn = notesCol
    Selection.Cells(1, notesCol).Activate
Exit Sub
errorHandler:
    MsgBox pos & "||" & playerName
End Sub
Sub setCustoms()
    
    Const topRow = 5
    
    Dim pos As String
    Dim posI As Integer
    Dim lastRow As Integer
    Dim eq1, eq2, eq3 As String
    Dim nameCells, notes As Range
    Dim lastYearCol, lastYear2hCol
    Dim eqVL, eqNotes As String
        
    'pos = ActiveSheet.Name 'for testing a single sheet
    For posI = 1 To 6
        pos = Range("PosList").Cells(posI)
        Worksheets(pos).Activate
        If pos = "DEF" Then
            lastYearCol = "S"
            lastYear2hCol = "R"
        Else
            lastYearCol = "X"
            lastYear2hCol = "W"
        End If
        lastRow = Range(pos & "_Data_1").Rows.Count + topRow - 1
        Set notes = Range(Range(pos & "_notes"), Cells(lastRow, Range(pos & "_notes").Column + 3))
        notes.Select
        'EQUATIONS
        
        eq1 = "=" & lastYearCol & topRow
        eq2 = "=RANK(" & lastYearCol & topRow & ",$" & lastYearCol & "$" & topRow & ":$" & lastYearCol & "$" & lastRow & ")"
        eq3 = "=RANK(" & lastYear2hCol & topRow & ",$" & lastYear2hCol & "$" & topRow & ":$" & lastYear2hCol & "$" & lastRow & ")"
        Set nameCells = Range(notes.Cells(1, 6), Cells(lastRow, notes.Column + 5))
        eqVL = "VLOOKUP(" & nameCells.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ",[espn.xls]" & pos & "!$C$3:$P$499,14,0)"
        eqNotes = "=IF(ISERROR(" & eqVL & "),""""," & eqVL & ")"
        eqName = "=IF(RIGHT(B5,1)=""®"",LEFT(B5,LEN(B5)-2),IF(RIGHT(B5,1)="")"",LEFT(B5,FIND("" ("",B5)-1),B5))"
        notes.Cells(1, 1) = eq1
        notes.Cells(1, 2) = eq2
        notes.Cells(1, 3) = eq3
        notes.Cells(1, 4) = eqNotes
        
        'nameCells.Cells(1, 1) = eqName  'don't do these again because names are being manually entered for conflicts
    
        notes.FillDown
        'nameCells.FillDown             'ditto
                
    Next posI
    Set notes = Nothing
End Sub
Sub finishUp()
    
    Dim ESPNRange As Range
    Dim pos As String
    Dim posI As Integer
    Dim lastRow As Integer
    Dim oldcalc
    oldcalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    For posI = 1 To 6
        pos = Range("PosList").Cells(posI)
        Worksheets(pos).Activate
        lastRow = Range(pos & "_Data_1").Rows.Count + topRow - 1
        Set notes = Range(Range(pos & "_notes"), Cells(lastRow, Range(pos & "_notes").Column + 3))
        Set nameCells = Range(notes.Cells(1, 6), Cells(lastRow, notes.Column + 5))
        Set ESPNRange = notes.Columns(4)
        With ESPNRange
            .RowHeight = 100
            .ColumnWidth = 130
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
    notes.Select
        Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    nameCells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Next posI
    Application.Calculation = oldcalc
End Sub

