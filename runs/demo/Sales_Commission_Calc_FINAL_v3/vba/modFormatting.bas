Option Explicit

' Tidies the Summary tab before it goes into the monthly deck.

Public Sub FormatSummary()
    With ThisWorkbook.Worksheets("Summary")
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Columns("A:F").AutoFit
    End With
End Sub

Public Sub ClearFilters()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then ws.AutoFilter.ShowAllData
    Next ws
End Sub