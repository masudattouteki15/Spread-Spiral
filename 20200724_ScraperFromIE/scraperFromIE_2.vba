Sub testIE()

    '1 start---------------------------------------------
    ' Dim objIE As InternetExplorer
    ' Set objIE = CreateObject("Internetexplorer.Application")
    
    ' ' objIE.Visible = True
    
    ' objIE.navigate "http://<hostname>/idworkflow/servlet"
    
    '1 end---------------------------------------------
    
    '2 start---------------------------------------------
    
    Dim shl As Object
    Set shl = CreateObject("Shell.Application")
    
    Dim targetTitle As String
    targetTitle = "IDワークフロー - 承認"
    
    Dim win As Object, getFlag As Boolean
    For Each win In shl.Windows
        If TypeName(win.document) = "HTMLDocument" Then
            If win.document.Title = targetTitle Then
                Dim objIE As InternetExplorer
                Set objIE = win
                
                getFlag = True
                Exit For
            End If
        End If
    Next
    
    If getFlag = False Then
        MsgBox "目的の画面が開かれていません。", vbExclamation
        Exit Sub
    End If
    
    '2 end---------------------------------------------
    
    '3 start---------------------------------------------
    '「HTMLDocument」を使う時は、参照設定の「Microsoft HTML Object Library」にチェック
    Dim htmlDoc As HTMLDocument
    Set htmlDoc = objIE.document
    
    
    ' 入力しよう
    
    Cells(1, 2).Value = Cells(1, 2).Value + 1
    Dim count As Integer
    count = Cells(1, 2).Value
    firstCount = Cells(1, 2).Value
    
    Dim i As Integer
    Dim reqMax As Integer
    reqMax = 10
    Dim toStr As String
    For i = 0 To reqMax
        ' MsgBox htmlDoc.getElementsById("maincontents").innerHTML
        If htmlDoc.getElementsByTagName("td")(4 + i * 6) Is Nothing Then
            Exit For
        Else
            toStr = CStr(htmlDoc.getElementsByTagName("td")(4 + i * 6).innerHTML)
            Cells(i + count, 1).Value = toStr
            Cells(i + count, 1).Value = Replace(Cells(i + count, 1).Value, "<br>", "")
            Cells(i + count, 2).Value = htmlDoc.getElementsByClassName("center")(1 + i).innerHTML ' 済
            ' Cells(count, 3).Value = htmlDoc.getElementsByClassName("editAttrVal")(0).innerHTML
        End If
        
    Next
    Cells(1, 2).Value = i + count - 1
    
    
    ' Sort
    Dim endCount As Integer
    endCount = Cells(1, 2).Value
    
    Call sortAsc(Cells(firstCount, 1), Cells(endCount, 2), Cells(firstCount, 2))
    Call sortAsc(Cells(firstCount, 1), Cells(endCount, 2), Cells(firstCount, 1))
    
    '3 end---------------------------------------------

End Sub

Sub onlySort()
    Dim i As Integer
    
    Call sortAsc(Cells(Selection(1).Row, Selection(1).Column), _
                    Cells(Selection(Selection.count).Row, Selection(Selection.count).Column), _
                    Cells(Selection(1).Row, Selection.Columns.count))
    Call sortAsc(Cells(Selection(1).Row, Selection(1).Column), _
                    Cells(Selection(Selection.count).Row, Selection(Selection.count).Column), _
                    Cells(Selection(1).Row, Selection(1).Column))
End Sub

