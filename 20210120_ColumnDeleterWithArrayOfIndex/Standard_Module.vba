
Sub DeleteColumn(col_indexes() As Variant)
    Dim col As Variant
    Dim i As Variant
    i = 1
    Dim j As Variant
    For j = 1 To UBound(col_indexes)
        If IsEmpty(col_indexes(j)) Then
            Exit For
        End If
        ActiveSheet.Range(Columns(col_indexes(j) - i + 1), Columns(col_indexes(j) - i + 1)).Delete
        i = i + 1
    Next
End Sub

Function ColumnsValuesRead(ByVal name_of_sheet As String) As Variant
    
    Dim i As Variant
    Dim j As Variant
    ' 消したい項目の一覧を取ってくる。（別のシートに置く？）
    Dim values() As Variant
    i = NumCounted(name_of_sheet)
    
    ReDim values(i - 1)
    For j = 1 To i - 1
        values(j) = Worksheets(name_of_sheet).Cells(1, j)
    Next
    
    ColumnsValuesRead = values
End Function

Function ColumnsIndexRead(ByRef values() As Variant) As Variant
    ' その項目と当該行の値が同じだったら、消す。
    Dim row_midashi As Variant '項目名（ex.［経由］）
    row_midashi = 20
    Dim i As Variant
    Dim j As Variant
    j = 1
    Dim num_all_values As Variant
    num_all_values = NumCounted(ActiveSheet.Name)
    Dim indexes() As Variant
    ReDim indexes(num_all_values)
    For i = 1 To num_all_values
        If values(j) = ActiveSheet.Cells(row_midashi, i).Value Then
            indexes(j) = i
            j = j + 1
            If j > UBound(values) Then
                Exit For
            End If
        End If
    Next
    
    ColumnsIndexRead = indexes
End Function

Function NumCounted(ByVal name_of_sheet As String) As Variant
    Dim i As Variant
    i = 1
    While Worksheets(name_of_sheet).Cells(1, i).Value <> ""
        i = i + 1
    Wend
    NumCounted = i
End Function
