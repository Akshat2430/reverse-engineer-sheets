Option Explicit

' FY26 Q4 accelerator (sales ops). Run once after Q4 close, before payroll export.
' Applies the 1.15x accelerator to Q4 deals and locks the number in Final Commission.

Public Sub ApplyAccelerator()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim colDate As Long, cNet As Long, cFinal As Long

    Set ws = ThisWorkbook.Worksheets("Calc")
    colDate = HeaderCol(ws, "Close Date")
    cNet = HeaderCol(ws, "Net Commission")
    cFinal = HeaderCol(ws, "Final Commission")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If IsDate(ws.Cells(r, colDate).Value) Then
            If Month(ws.Cells(r, colDate).Value) >= 10 Then
                ws.Cells(r, cFinal).Value = ws.Cells(r, cNet).Value * 1.15
            End If
        End If
    Next r
End Sub

Private Function HeaderCol(ws As Worksheet, headerText As String) As Long
    Dim c As Range
    Set c = ws.Rows(1).Find(What:=headerText, LookAt:=xlWhole)
    If c Is Nothing Then Err.Raise vbObjectError + 513, , "Header not found: " & headerText
    HeaderCol = c.Column
End Function