Dim firstRow As Integer
Dim mochidashibangou_column As Integer

Dim sKeizokubun As String, sKonnendo As String


Private Sub CheckBox1_Change()
    'Dim i As Integer
    'Dim bZennendo As Boolean
    'bZennendo = False
    
    '色付けしない行(大項目が記述されている行)の設定 ---START---
    'If CheckBox1.Value = True Then
    '    i = firstRow
    '    While Cells(i, mochidashibangou_column).Value <> ""
    '        If Cells(i, mochidashibangou_column).Interior.Color = RGB(252, 228, 214) Then
    '            If bZennendo = False Then
    '                sKeizokubun = Cells(i, mochidashibangou_column).Value
    '                bZennendo = True
    '            Else
    '                sKonnendo = Cells(i, mochidashibangou_column).Value
    '            End If
    '        End If
    '        i = i + 1
    '    Wend
    'End If
    '色付けしない行(大項目が記述されている行)の設定 ---END---
End Sub

Private Sub CheckBox2_Change() 'フォームをしましまに塗り直す
    Dim a As Boolean
    a = True '　※　塗りたい場合は、Falseにする。　※
    If a = False Then
        If CheckBox2.Value = True Then
        
            Dim i As Integer, j As Integer
            Dim midashi_row As Integer '見出しが載っている行インデックス
            
            firstRow = 9 '走査し始める行インデックス　※複数個所で宣言されています
            i = firstRow
            mochidashibangou_column = 2 '持出番号の列インデックス　※複数個所で宣言されています
            midashi_row = 5
            
            While Cells(i, mochidashibangou_column).Value <> "" '持出番号が打ち込まれている間
                j = mochidashibangou_column '持出番号の列からスタート
                If i Mod 2 = 1 Then '奇数行に対して
                    While Cells(midashi_row, j).Value <> "" '見出しが空欄ではない間
                        If Cells(midashi_row, j).Value <> "持出日" And Cells(midashi_row, j).Value <> "持帰日" Then '見出しが「持出日」か「持帰日」ではない場合
                            Cells(i, j).Interior.Color = RGB(230, 255, 230)
                        Else
                            ' そのままの色
                        End If
                        j = j + 1
                    Wend
                Else
                    While Cells(midashi_row, j).Value <> "" '見出しが空欄ではない間
                        If Cells(midashi_row, j).Value <> "持出日" And Cells(midashi_row, j).Value <> "持帰日" Then '見出しが「持出日」か「持帰日」ではない場合
                            Cells(i, j).Interior.Color = RGB(255, 255, 255)
                        Else
                            ' そのままの色
                        End If
                        j = j + 1
                    Wend
                End If
                i = i + 1
            Wend
            
        End If
    End If
End Sub

Public Sub Worksheet_Active()
    Call mochi_check
    CheckBox1.Caption = "年度更新用" '色付けしない行(大項目が記述されている行)の設定用
    CheckBox2.Caption = "しましま化" '色付けしない行(大項目が記述されている行)の設定用
    
    '色をRGBの数値に変換して出力する ---START---
        'myColor = Cells(7, 2).Interior.Color
        'myR = myColor Mod 256
        'myG = Int(myColor / 256) Mod 256
        'myB = Int(myColor / 256 / 256)
        'Cells(8, 2).Value = "#" & Right("0" & Hex(myR), 2) & _
         '   Right("0" & Hex(myG), 2) & _
         '   Right("0" & Hex(myB), 2)
    '色をRGBの数値に変換して出力する ---END---

End Sub

Public Sub Worksheet_Change(ByVal Target As Range)
    Call choufuku_check
    Call mochi_check
End Sub

Public Sub choufuku_check() '重複する持出番号が入力されたかチェック。
    Dim i As Integer, j As Integer
    
    firstRow = 9 '走査し始める行インデックス　※複数個所で宣言されています
    i = firstRow
    mochidashibangou_column = 2 '持出番号の列インデックス　※複数個所で宣言されています
    
    While Cells(i, mochidashibangou_column).Value <> "" '持出番号が打ち込まれている間
        j = i + 1 '繰り返し変数を初期化
        While Cells(j, mochidashibangou_column).Value <> "" '持出番号が打ち込まれている間
            If Cells(i, mochidashibangou_column).Value = Cells(j, mochidashibangou_column).Value Then '重複している持出番号がある場合
                'MsgBox "重複している持出番号があります。", vbOKOnly
            Else
                'そのままの色
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend
    
End Sub

Public Sub mochi_check() '所定の列を走査する。
    Dim i As Integer
    Dim yotei_to_column As Integer, mochikaeri_column As Integer
    Dim yotei_from_column As Integer, mochidashi_column As Integer
    
    firstRow = 9 '走査し始める行インデックス　※複数個所で宣言されています
    
    mochidashibangou_column = 2 '持出番号の列インデックス　※複数個所で宣言されています
    
    sKeizokubun = "【2018年度申請継続分】"
    sKonnendo = "2019年度"

    yotei_to_column = 7 '予定期間(TO)の列インデックス
    mochikaeri_column = 12 '持帰日の列インデックス
    
    yotei_from_column = 6 '予定期間(FROM)の列インデックス
    mochidashi_column = 11 '持出日の列インデックス
    
    i = firstRow
    While Cells(i, mochidashibangou_column).Value <> "" '持出番号が打ち込まれている間
        If Cells(i, mochidashibangou_column).Value = sKeizokubun Then '行の2列目のセルの値がsKeizokubunだったら、
            
        ElseIf Cells(i, mochidashibangou_column).Value = sKonnendo Then '行の2列目のセルの値がsKonnendoだったら、
            
        Else
            '持帰日が記入されているか、もしくは予定期間(TO)と相違がないか。
            Call mochikaeri_color(i, yotei_to_column, mochikaeri_column)
            '持出日が記入されているか、もしくは予定期間(FROM)と相違がないか。
            Call mochidashi_color(i, yotei_from_column, mochidashi_column)
        End If
        i = i + 1
    Wend
End Sub

Public Sub mochikaeri_color(row As Integer, yotei_column As Integer, mochi_column As Integer) '場合に合わせてセルを塗る
    If Cells(row, yotei_column).Value <> "" Then '予定(TO)が打ち込まれていない場合
        If Cells(row, yotei_column).Value < Date Then
            If Cells(row, mochi_column).Value = "" Then '日付が打ち込まれていない場合、
                Cells(row, mochi_column).Interior.Color = RGB(255, 210, 0)
            ElseIf Cells(row, mochi_column).Value > Cells(row, yotei_column).Value Then '予定(TO)より持帰日が遅い場合、
                Cells(row, mochi_column).Interior.Color = RGB(255, 70, 70)
            Else
                Cells(row, mochi_column).Interior.Color = Cells(row, mochi_column + 1).Interior.Color '問題なし、右隣のセルと同じ色に塗る
            End If
        Else
            Cells(row, mochi_column).Interior.Color = Cells(row, mochi_column + 1).Interior.Color  '問題なし、右隣のセルと同じ色に塗る
        End If
    Else
        Cells(row, mochi_column).Interior.Color = Cells(row, mochi_column + 1).Interior.Color  '問題なし、右隣のセルと同じ色に塗る
    End If
End Sub

Public Sub mochidashi_color(row As Integer, yotei_column As Integer, mochi_column As Integer) '場合に合わせてセルを塗る
    If Cells(row, yotei_column).Value <> "" Then '予定(FROM)が打ち込まれていない場合
        If Cells(row, yotei_column).Value < Date Then '今日が予定(FROM)を過ぎている場合、
            If Cells(row, mochi_column).Value = "" Then '日付が打ち込まれていない場合
                Cells(row, mochi_column).Interior.Color = RGB(255, 210, 0)
            ElseIf Cells(row, mochi_column).Value < Cells(row, yotei_column).Value Then '予定(FROM)より持出日が早い場合、
                Cells(row, mochi_column).Interior.Color = RGB(255, 70, 70)
            Else
                Cells(row, mochi_column).Interior.Color = Cells(row, mochi_column - 1).Interior.Color '問題なし、左隣のセルと同じ色に塗る
            End If
        Else
            Cells(row, mochi_column).Interior.Color = Cells(row, mochi_column - 1).Interior.Color  '問題なし、左隣のセルと同じ色に塗る
        End If
    Else
        Cells(row, mochi_column).Interior.Color = Cells(row, mochi_column - 1).Interior.Color  '問題なし、左隣のセルと同じ色に塗る
    End If
End Sub
