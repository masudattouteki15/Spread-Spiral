Sub mergeEachRColumn()
    Dim i As Integer
    For i = 1 To Selection(Selection.Count).Column
        Range(Cells(Selection(1).Row, i), Cells(Selection(1).Row + 1, i)).Merge
    Next
End Sub
