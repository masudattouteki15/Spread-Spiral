
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

Function RangeValuesRead(ByVal name_of_sheet As String, ByVal num_of_rows As Variant, ByVal num_of_columns As Variant) As Variant
    Dim i As Variant
    Dim j As Variant
    
    Dim values() As Variant
    ReDim values(num_of_rows, num_of_columns)
    row_index_wanted_to_escape = 1
    column_index_wanted_to_escape = 1
    
    For i = row_index_wanted_to_escape To num_of_rows
        For j = column_index_wanted_to_escape To num_of_columns
            values(i, j) = ActiveSheet.Cells(i, j).Value
        Next
    Next
    
    RangeValuesRead = values
    
End Function

Function ColumnsValuesRead(ByVal name_of_sheet As String) As Variant
    Dim i As Variant
    Dim j As Variant
    Dim row_index_wanted_to_delete As Variant
    
    Dim values() As Variant
    row_index_wanted_to_delete = 1
    column_index_wanted_to_escape = 1
    i = NumOfColumnCounted(name_of_sheet, row_index_wanted_to_delete)
    
    ReDim values(i - 1)
    For j = column_index_wanted_to_escape To i - 1
        values(j) = Worksheets(name_of_sheet).Cells(1, j)
    Next
    
    ColumnsValuesRead = values
End Function

Function ColumnsIndexRead(ByVal row_index_of_midashi As Variant, ByRef values() As Variant) As Variant
    ' その項目と当該行の値が同じだったら、消す。
    ' row_index_of_midashi：項目名（ex.［経由］）が載っている行インデックス
    ' values     ：消したい項目名が格納された配列
    Dim i As Variant
    Dim j As Variant
    j = 1
    
    Dim num_all_values_wanted_to_delete As Variant '消したい項目の数
    num_all_values_wanted_to_delete = NumOfColumnCounted(ActiveSheet.Name, row_index_of_midashi)
    
    Dim indexes() As Variant
    ReDim indexes(num_all_values_wanted_to_delete)
    
    For i = 1 To num_all_values_wanted_to_delete
        If values(j) = ActiveSheet.Cells(row_index_of_midashi, i).Value Then
            indexes(j) = i
            j = j + 1
            If j > UBound(values) Then
                Exit For
            End If
        End If
    Next
    
    ColumnsIndexRead = indexes
End Function

Sub WriteRangeValues(ByRef values() As Variant)
    Dim i As Variant
    Dim j As Variant
    
    ' 書き始める行インデックスと列インデックス
    row_index_wanted_to_start = 1
    column_index_wanted_to_start = 1
    
    Dim columns_of_values() As Variant
    ' redim columns_of_values()
    ' columns_of_values() = values(column_index_wanted_to_start)
    
    For i = row_index_wanted_to_start To UBound(values, 1)
        For j = column_index_wanted_to_start To UBound(values, 2)
            ActiveSheet.Cells(i, j).Value = values(i, j)
        Next
    Next
End Sub

Function NumOfColumnCounted(ByVal name_of_sheet As String, ByVal row_index As Variant) As Variant
    Dim i As Variant
    i = 1
    While Worksheets(name_of_sheet).Cells(row_index, i).Value <> ""
        i = i + 1
    Wend
    NumOfColumnCounted = i
End Function
