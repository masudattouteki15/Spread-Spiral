' 「標準モジュール」として作らないと、Excel内の関数として利用できない。
' 空白文字を読み取りたかったが、文字化けしているっぽくて、断念した。
' なので、未完成。
Function COUNTSPACE(ByVal r As Range)
    Dim count As Integer
    Dim ary As Variant
    ' VBAの場合、イタレーターはVariant型じゃないとダメ。
    Dim i As Variant
    ' 「Set」を付けないとオブジェクトは変換できない。
    Set ary = r
    For Each i In ary
        If i Like " *" Then
            count = count + 1
        End If
    Next i
    COUNTSPACE = count
End Function
