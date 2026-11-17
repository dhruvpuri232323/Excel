Attribute VB_Name = "Module1"
Sub FormatDailyReport()
' Formats the day's orders for Farah: header, numbers, a total row
    With Worksheets("Orders")
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(221, 235, 247)
        .Range("G2:H13").NumberFormat = "#,##0"
        .Range("A14").Value = "Total"
        .Range("A14").Font.Bold = True
        .Range("H14").Formula = "=SUM(H2:H13)"
        .Range("H14").Font.Bold = True
        .Columns("A:H").AutoFit
    End With
End Sub
