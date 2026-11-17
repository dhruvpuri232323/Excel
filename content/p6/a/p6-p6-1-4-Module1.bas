Attribute VB_Name = "Module1"
Sub FormatReport()
    Range("A1:H1").Font.Bold = True
    Range("A1:H1").Interior.Color = 16247773
    Range("A1").CurrentRegion.Columns("G:H").NumberFormat = "#,##0"
    Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
End Sub
