
Sub Button_Del_Click()
    Dim name_of_sheet_2 As String
    name_of_sheet_2 = "消したい項目を1行目に貼る"
    Dim row_index_of_midashi As Variant
    row_index_of_midashi = 20
    
    Dim num_of_rows_wanted_to_escape As Variant
    Dim num_of_columns_wanted_to_escape As Variant
    num_of_rows_wanted_to_escape = row_index_of_midashi - 1
    num_of_columns_wanted_to_escape = 3
    
    ' 消したくないヘッダ情報を配列へ避難させる。
    'Dim values_wanted_to_escape(num_of_rows_wanted_to_escape, num_of_columns_wanted_to_escape) As Variant
    Dim values_wanted_to_escape() As Variant
    ReDim values_wanted_to_escape(num_of_rows_wanted_to_escape, num_of_columns_wanted_to_escape)
    values_wanted_to_escape() = RangeValuesRead(ActiveSheet.Name, num_of_rows_wanted_to_escape, num_of_columns_wanted_to_escape)
    
    ' 消したい項目を配列へ取得する。
    Dim values_wanted_to_delete() As Variant
    values_wanted_to_delete() = ColumnsValuesRead(name_of_sheet_2)
    
    ' 消したい項目が位置するインデックスを特定して配列へ取得する。
    Dim col_indexes_wanted_to_delete() As Variant
    col_indexes_wanted_to_delete() = ColumnsIndexRead(row_index_of_midashi, values_wanted_to_delete)
    
    ' 取得した配列を基に列を削除する。
    Call DeleteColumn(col_indexes_wanted_to_delete)
    
    ' 退避させたヘッダ情報を書き戻す。
    Call WriteRangeValues(values_wanted_to_escape)
End Sub

