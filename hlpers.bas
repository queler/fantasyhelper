Attribute VB_Name = "hlpers"
Option Explicit
Const cCommandBarID = "DraftBar"
Const OFF_NOTES_SN As String = "off-notes"
Const POS_RG As String = "_poslist_main"
'Const WB_NAME As String = "ccdb_football_2011_2.1(2)"

Const POS_NAMES_NM As String = "_names"
Const FFURL_D As String = "http://www.fftoday.com/stats/players?TeamID="
Const FFURL_PL As String = "http://www.fftoday.com/stats/players/" '"http://www.fftoday.com/stats/players/"
Const LY_STAT_HR As String = "2015 Stats"
Const CALC_HR = "FF POINT CALCULATIONS"
Const CY_STAT_HR As String = "2016 Projected"
Const ESPN_SEARCH_URL As String = "http://games.espn.go.com/ffl/tools/projections?display=alt&avail=-1&search="

Sub CreateCM()
'routine to create the custom menu
Dim cbBar As CommandBar, cbCtrl As CommandBarControl
Dim posI As Integer
Dim pos As String
'dim pos as New range=("QB",RB")
        
DeleteCM
'call deletion routine first to ensure no duplicate menus created
Set cbBar = CommandBars.Add(name:=cCommandBarID, Position:=msoBarTop)
'create the custom menu
For posI = 1 To 6
        pos = Application.Range(POS_RG).Cells(posI)
    Set cbCtrl = cbBar.Controls.Add(Type:=msoControlButton)
    'add a control to the menu
    With cbCtrl
    'properties of the control
        .OnAction = "show" & pos & "s" 'Routine called by button
        .Style = msoButtonCaption  'Style of button
        .Caption = pos
    End With

Next posI
'notes
Set cbCtrl = cbBar.Controls.Add(Type:=msoControlButton)
cbCtrl.OnAction = "LaunchESPN"
cbCtrl.Caption = "ESPN"
cbCtrl.Style = msoButtonCaption
'make it visible
cbBar.Visible = True
Set cbCtrl = cbBar.Controls.Add(Type:=msoControlButton)
cbCtrl.OnAction = "fftoday"
cbCtrl.Caption = "FFToday"
cbCtrl.Style = msoButtonCaption
'make it visible
cbBar.Visible = True
Set cbBar = Nothing
Set cbCtrl = Nothing
'free memory

End Sub

Sub DeleteCM()
'routine to delete the custom menu
Dim cbBar As CommandBar

On Error Resume Next
'in case its already deleted
Set cbBar = CommandBars(cCommandBarID)
'reference the custom menu
cbBar.Delete
'delete the bar
On Error GoTo 0
'reset error checking
Set cbBar = Nothing
'free memory

End Sub

Sub CMVisibility(showme As Boolean)
'routine to switch visibility of the custom menu

On Error Resume Next
'in case its already gone
CommandBars(cCommandBarID).Visible = showme
'change the visibility status
On Error GoTo 0
'reset error checking

End Sub


Sub showQBs()
    gotoSelect ("b9")

    
End Sub

Sub showRBs()
    gotoSelect ("b71")

    
End Sub

Sub showWRs()
    gotoSelect ("b173")

    
End Sub

Sub showTEs()
    gotoSelect ("b275")

    
End Sub

Sub showKs()
    gotoSelect ("b315")

    
End Sub

Sub showDEFs()
    gotoSelect ("b349")
    
End Sub
Sub gotoSelect(where As String)
        Dim cell As Range
        Set cell = Sheets(OFF_NOTES_SN).Range(where)
        Application.GoTo cell, True
        cell.Offset(0, 3).Select
    'Application.Goto ActiveCell, True

End Sub
Sub pasterInternal(notes As Range, data As Range)
    
    Dim y As Integer, lastCol As Integer, notesCol As Integer
    Dim thisSheet As Worksheet
    Set thisSheet = data.Worksheet
    lastCol = lastYearCol(thisSheet)
    notesCol = notes.Column
    Dim oldcalc As XlCalculation
    Application.ScreenUpdating = False
    oldcalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    For y = notes.row To data.Rows(data.Rows.Count).row
        thisSheet.Cells(y, notesCol) = thisSheet.Cells(y, lastCol)
        thisSheet.Cells(y, notesCol + 1) = thisSheet.Cells(y, lastCol - 1)
                

    Next y
    Application.Calculation = oldcalc
    Application.ScreenUpdating = True
End Sub

Function lastYearCol(it As Worksheet)
    Dim i As Integer
    i = Range("qb_data").Column
    
    Do Until it.Cells(2, i) = CALC_HR
        i = i + 1
    Loop
    Do Until (it.Cells(3, i) = LY_STAT_HR) Or (it.Cells(3, i) = CY_STAT_HR)
        i = i - 1
    Loop
    If it.Cells(3, i) = CY_STAT_HR Then
        Do Until (it.Cells(3, i) = LY_STAT_HR)
            i = i + 1
        Loop
    End If
    lastYearCol = i
End Function
Sub AddPastYear(pos As String)
    Dim strNotes As String
    Dim strData As String
    strNotes = pos & "_Notes"
    strData = pos & "_Data"
    Call pasterInternal(Range(strNotes), Range(strData))
    
End Sub
Sub pastAll()
    Call AddPastYear("QB")
    Call AddPastYear("RB")
    Call AddPastYear("WR")
    Call AddPastYear("TE")
    Call AddPastYear("K")
    Call AddPastYear("Def")
    
    
End Sub
Sub LaunchESPN()
LaunchURLByLastName ESPN_SEARCH_URL, ""
End Sub

Sub LaunchURLByLastName(preurl As String, posturl As String)
    Dim lastname As String
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Pattern = "\s([-\w]+)(?:(?:\s\W)|$)"
    lastname = r.Execute(ActiveCell.Value).Item(0).SubMatches(0)
    ActiveWorkbook.FollowHyperlink preurl & lastname & posturl
End Sub
Sub fftodayOld()
    Dim compiler As Workbook
   
   ' On Error GoTo e
    Set compiler = ActiveWorkbook
    On Error GoTo 0
    Dim row As Integer
    Dim pos As String
    pos = "WR"
    row = GetRow(ActiveCell.Value, compiler.Names(pos & POS_NAMES_NM).RefersToRange)
    If row <> -1 Then GoTo Found
    pos = "RB"
    row = GetRow(ActiveCell.Value, compiler.Names(pos & POS_NAMES_NM).RefersToRange)
    If row <> -1 Then GoTo Found
    pos = "QB"
    row = GetRow(ActiveCell.Value, compiler.Names(pos & POS_NAMES_NM).RefersToRange)
    If row <> -1 Then GoTo Found
    pos = "TE"
    row = GetRow(ActiveCell.Value, compiler.Names(pos & POS_NAMES_NM).RefersToRange)
    If row <> -1 Then GoTo Found
    pos = "DEF"
    row = GetRow(ActiveCell.Value, compiler.Names(pos & POS_NAMES_NM).RefersToRange)
    If row <> -1 Then GoTo Found
    pos = "K"
    row = GetRow(ActiveCell.Value, compiler.Names(pos & POS_NAMES_NM).RefersToRange)
    If row <> -1 Then GoTo Found
    Err.Raise vbError + 9, , "not found"
       
Found:
    Dim col As Integer
    col = 5
    Dim id As String
    id = compiler.Names(pos & "_names").RefersToRange.Worksheet.Cells(row, col).Value
    If pos = "DEF" Then
        ActiveWorkbook.FollowHyperlink FFURL_D & id
    Else
        ActiveWorkbook.FollowHyperlink FFURL_PL & id
    End If
Exit Sub
e:
    'Set compiler = Workbooks.Open(WB_NAME)
   ' compiler.Windows(1).Visible = False
    Resume Next
End Sub
Function GetRow(name As String, r As Range) As Integer
    Dim i As Integer
    For i = 1 To r.Rows.Count
        If r.Cells(i, 1) = name Then
            GetRow = r.Rows(i).row
            Exit Function
        End If
    Next i
    GetRow = -1
End Function
Public Sub test()
    MsgBox ThisWorkbook.name
    
End Sub

Public Sub GetFFTLinks()
Const strFFT = "fftLinks"

    
On Error GoTo trap
    Dim html As New HTMLDocument
    Dim page As HTMLDocument
    Dim row As Long
    row = 1
    Dim sht As Worksheet
    Dim span As IHTMLElement, child As IHTMLElement
    Dim a As IHTMLElement
    'reg.Pattern = "(\S+),\s*(\S+)\s+([A-Z]{2,3})"
    Const LINKS_URL = "http://www.fftoday.com/stats/players?Pos="
    Dim posArray
    posArray = Array("QB", "RB", "WR", "TE", "K")
    Dim pos
'    On Error Resume Next
'    Set sht = Nothing
'    Set sht = Sheets(strFFT)
'        On Error GoTo 0
'    If sht Is Nothing Then
'        Set sht = Sheets.Add
'        sht.name = strFFT
'        sht.Move after:=Sheets(Sheets.Count)
'
'    End If
    Set sht = fft
    sht.Cells.Clear
    Application.ScreenUpdating = False
    Dim list As New Collection
    Dim comma As Integer
    Dim lastSpace As Integer
    Dim text As String
    For Each pos In posArray
        Set page = html.createDocumentFromUrl(LINKS_URL & pos, "")
        Do Until page.readyState = "complete": DoEvents: Loop
        For Each span In page.getElementsByClassName("bodycontent")
            If LCase$(span.tagName) = "span" Then
                
                For Each child In span.Children
                    If LCase$(child.tagName) = "a" Then
                        text = child.innerText
                        'Set matchName = reg.Execute(child.innerText).Item(0)
'                        sht.Cells(row, 1).Value = CStr(matchName.SubMatches(0))
'                        sht.Cells(row, 2).Value = matchName.SubMatches(1)
'                        sht.Cells(row, 3).Value = matchName.SubMatches(2)
'                        sht.Cells(row, 4).Value = pos
'                        sht.Cells(row, 5).Value = matchName.SubMatches(1) & " " & matchName.SubMatches(0)
'                        sht.Cells(row, 6).Value = child.getAttribute("href")
                        lastSpace = InStrRev(text, " ")
                        comma = InStr(text, ",")
                        list.Add Array(Mid(text, comma + 2, lastSpace - comma - 2), Left(text, comma - 1), _
                            Mid(text, lastSpace + 1), pos, _
                            child.getAttribute("href"))
                        
                        Exit For 'child
                        
                    End If 'a
                Next child
            End If 'span
        
        Next span
    
    Next pos
    Debug.Print "list loaded, converting to 2d-array"
    Dim arr As Long
    Dim col
    Application.ScreenUpdating = False
Dim OrigCalc As Excel.XlCalculation
OrigCalc = Application.Calculation
Application.Calculation = xlCalculationManual
    Dim arr2() As Variant
    Dim lb, ub
    lb = LBound(list(1))
    ub = UBound(list(1))
    ReDim arr2(1 To list.Count, 1 To ub - lb + 1)
    For arr = 1 To list.Count
        For col = lb To ub
            arr2(arr, col - lb + 1) = list.Item(arr)(col)
        Next col
    
    Next arr
    Debug.Print "array loaded"
    sht.Range(sht.Cells(1, 1), sht.Cells(list.Count, ub - lb + 1)).Value = arr2
Application.ScreenUpdating = True
Application.Calculation = OrigCalc
Exit Sub
trap:
    Resume
End Sub
Public Sub fftoday()
    Dim arr() As Variant
    arr = fft.Cells.CurrentRegion
    Dim sel As String
    sel = ActiveCell.Value
    On Error GoTo 0
        Dim row As Integer
        Dim pos As String
        pos = "WR"
        row = GetRow(sel, Names(pos & POS_NAMES_NM).RefersToRange)
        If row <> -1 Then GoTo Found
        pos = "RB"
        row = GetRow(sel, Names(pos & POS_NAMES_NM).RefersToRange)
        If row <> -1 Then GoTo Found
        pos = "QB"
        row = GetRow(sel, Names(pos & POS_NAMES_NM).RefersToRange)
        If row <> -1 Then GoTo Found
        pos = "TE"
        row = GetRow(sel, Names(pos & POS_NAMES_NM).RefersToRange)
        If row <> -1 Then GoTo Found
        pos = "DEF"
        row = GetRow(sel, Names(pos & POS_NAMES_NM).RefersToRange)
        If row <> -1 Then GoTo Found
        pos = "K"
        row = GetRow(sel, Names(pos & POS_NAMES_NM).RefersToRange)
        If row <> -1 Then GoTo Found
        Err.Raise vbError + 9, , "not found"
Found:
    Dim team As String
    Dim url As String
    Dim rookieless As String
    Dim rLoc As Integer
   'not worth it to do DEF
    rLoc = InStr(sel, " ®")
    If rLoc > 0 Then rookieless = Left(sel, rLoc - 1) Else rookieless = sel
    team = ActiveWorkbook.Names(pos & "_names").RefersToRange.Worksheet.Cells(row, 3).Value
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If pos = arr(i, 4) Then 'pos match
            If team = arr(i, 3) Then 'team match
                If InStr(arr(i, 1) & " " & arr(i, 2), rookieless) > 0 _
                    Or InStr(rookieless, arr(i, 1) & " " & arr(i, 2)) > 0 Then
                    url = arr(i, 5)
                    Exit For
                End If 'name
                
            End If 'team
        End If 'pos
        
    Next i
    If url <> "" Then ActiveWorkbook.FollowHyperlink url
End Sub
