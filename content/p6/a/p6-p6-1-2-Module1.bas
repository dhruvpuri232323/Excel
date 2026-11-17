Attribute VB_Name = "Module1"
Sub MarkDispatched()
' Marks every order as dispatched today, for the dispatch register
    Dim r As Long
    With Worksheets("Orders")
        .Range("I1").Value = "Dispatched"
        .Range("I1").Font.Bold = True
        For r = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
            .Cells(r, 9).Value = "Yes"
        Next r
    End With
End Sub
