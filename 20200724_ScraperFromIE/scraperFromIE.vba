Sub testIE()

    '1 start---------------------------------------------
    ' Dim objIE As InternetExplorer
    ' Set objIE = CreateObject("Internetexplorer.Application")
    
    ' ' objIE.Visible = True
    
    ' objIE.navigate "http://piopal015/idworkflow/servlet"
    
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
    Dim htmlDoc As HTMLDocument
    Set htmlDoc = objIE.document
    
    
    ' 入力しよう
    Cells(1, 2).Value = Cells(1, 2).Value + 1
    Dim count As Integer
    count = Cells(1, 2).Value
    
    Dim i As Integer
    Dim reqMax As Integer
    reqMax = 10
    Dim toStr As String
    
    ' MsgBox htmlDoc.getElementsById("maincontents").innerHTML
    ' Cells(count, 1).Value = htmlDoc.getElementsByClassName("editAttrVal")(0).innerHTML
    ' Cells(count, 2).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 3).Value = htmlDoc.getElementsByClassName("editAttrVal")(0).innerHTML ' 済
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(0).innerHTML
    Cells(count, 3).Value = Replace(toStr, vbLf, "") ' 改行を削除
    ' Cells(count, 4).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 5).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 6).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 7).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 8).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 9).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 10).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 11).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    Cells(count, 12).Value = "新規" ' 済
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(7).innerHTML
    Cells(count, 13).Value = Replace(Left(toStr, InStr(toStr, "<") - 1), vbLf, "") ' 改行と後ろのタグを削除
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(8).innerHTML
    Cells(count, 14).Value = Replace(Left(toStr, InStr(toStr, "<") - 1), vbLf, "") ' 改行と後ろのタグを削除
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(9).innerHTML
    Cells(count, 15).Value = Replace(Left(toStr, InStr(toStr, "<") - 1), vbLf, "") ' 改行と後ろのタグを削除
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(10).innerHTML
    Cells(count, 16).Value = Replace(Left(toStr, InStr(toStr, "<") - 1), vbLf, "") ' 改行と後ろのタグを削除
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(11).innerHTML
    Cells(count, 17).Value = Replace(Left(toStr, InStr(toStr, "<") - 1), vbLf, "") ' 改行と後ろのタグを削除
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(12).innerHTML
    Cells(count, 18).Value = Replace(Left(toStr, InStr(toStr, "<") - 1), vbLf, "") ' 改行と後ろのタグを削除
    ' Cells(count, 19).Value = htmlDoc.getElementsByClassName("editAttrVal")(0).innerHTML
    toStr = CStr(htmlDoc.getElementsByClassName("editAttrVal")(3).innerHTML)
    Cells(count, 20).Value = toStr ' 0落ちしている
    MsgBox TypeName(Cells(count, 20).Value)
    Cells(count, 21).Value = htmlDoc.getElementsByClassName("editAttrVal")(4).innerHTML ' 済
    toStr = CStr(htmlDoc.getElementsByClassName("editAttrVal")(5).innerHTML)
    Cells(count, 22).Value = toStr ' 0落ちしている
    MsgBox TypeName(Cells(count, 22).Value)
    
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(1).innerHTML
    Cells(count, 23).Value = Replace(Left(toStr, InStr(toStr, "&") - 1), vbLf, "") ' 改行と後ろの空白コード以降を削除
    
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(6).innerHTML
    Cells(count, 24).Value = Replace(Replace(toStr, " ", ""), vbLf, "") ' 改行と先頭の半角スペースを削除
    
    toStr = htmlDoc.getElementsByClassName("editAttrVal")(0).innerHTML
    Cells(count, 25).Value = Replace(toStr, vbLf, "") ' 改行を削除
    ' Cells(count, 26).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 27).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 28).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 29).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    ' Cells(count, 30).Value = htmlDoc.getElementsByClassName("contentsTitle clear")(0).innerHTML
    
    '3 end---------------------------------------------

End Sub
