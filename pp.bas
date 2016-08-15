Attribute VB_Name = "pp"
Option Explicit
Public oHtml As New HTMLDocument
Public Const ROW_CLASS = "pncPlayerRow"
Public Const URL_PROJ = "http://games.espn.com/ffl/tools/projections?&slotCategoryId=$POS$&startIndex=$I40$"
Sub waitdoc(doc)
Do While doc.readyState <> "complete": DoEvents: Loop
End Sub
Function htmlTableFromRowClass(url As String, class As String) As Variant
    Dim doc As HTMLDocument
    Set doc = getDocUrl(url)
    'Call waitdoc(doc)
    
    Dim table As Object
    Set table = doc.getElementsByClassName(class)
    Dim arr As Variant
    ReDim arr(0 To table.Length - 1, 0 To 23)
    Dim row As HTMLTableRow, i As Integer, j As Integer
    Dim drift As Integer
    For i = 0 To table.Length - 1
        drift = 0
        Set row = table.Item(i)
        For j = 0 To row.Cells.Length - 1
        With row.Cells.Item(j)
            Select Case j
                Case 1:
                    arr(i, j + drift) = Left(.innerText, InStr(.innerText, ",") - 1)
                Case 2:
                    Dim ca As Variant
                    ca = Split(.innerText, "/")
                    arr(i, j + drift) = ca(0)
                    drift = drift + 1
                    arr(i, j + drift) = ca(1)
                    
            Case Else
            arr(i, j + drift) = .innerText
            End Select
        End With
        Next j
        
    Next i
    htmlTableFromRowClass = arr
End Function

Function getDocUrl(url As String)
    Dim xml As New XMLHTTP
    xml.Open "GET", url, False
    Dim doc As Object
    xml.send
    Set doc = CreateObject("htmlfile")
    doc.Open
    doc.write xml.responseText
    Set getDocUrl = doc
   ' Set doc = Nothing
End Function
