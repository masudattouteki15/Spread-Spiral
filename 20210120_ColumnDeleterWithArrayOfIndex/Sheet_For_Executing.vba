
Sub Button_Del_Click()
    Dim col_indexes() As Variant
    Dim values()      As Variant
    Dim name_of_sheet As String
    name_of_sheet = "消したい項目を1行目に貼る"
    values()      = ColumnsValuesRead(name_of_sheet)
    col_indexes() = ColumnsIndexRead(values)
    Call DeleteColumn(col_indexes)
End Sub

