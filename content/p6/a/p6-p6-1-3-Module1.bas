Attribute VB_Name = "Module1"
Sub FormatHeader()
    Range("A1:H1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .Color = 16247773
    End With
End Sub

Sub StampChecked()
    ActiveCell.FormulaR1C1 = "Checked"
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub
