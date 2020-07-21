Sub ScrapeYetOpenedPage()

  Dim objIE As InternetExplorer 'IEオブジェクトを準備
  Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット

  objIE.Visible = True 'IEを表示

  objIE.navigate "https://secure.goldpoint.co.jp/gpm/authentication/index.html" 'IEでURLを開く

  Call WaitResponse(objIE) '読み込み待ち 

 End Sub


Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub

Sub ScrapeOpenedPage()
 
    Dim shl As Object 'シェルオブジェクト生成
    Set shl = CreateObject("Shell.Application")
    
    Dim targetTitle As String '取得したいウィンドウのタイトルを設定
    targetTitle = "コミュニティ「ノンプログラマーのためのスキルアップ研究会」についてのお知らせ #ノンプロ研"
    
    Dim win As Object, getFlag As Boolean
    For Each win In shl.Windows '起動中のウィンドウを順番にみていく
        
        'IEとエクスプローラがシェルで取得されるため、IEのみ処理
        If TypeName(win.document) = "HTMLDocument" Then
            If win.document.Title = targetTitle Then
    
                Dim objIE As New InternetExplorer
                Set objIE = win
                
                getFlag = True '正しく取得できた
                Exit For
            End If
        End If
        
    Next
    
    If getFlag = False Then
        MsgBox "目的の画面が開かれていません。", vbExclamation
        Exit Sub
    End If
    
    '目的の画面のHTMLを読み込む
    Dim htmlDoc As HTMLDocument
    Set htmlDoc = objIE.document
    
    ' 欲しいクラス名の中の文字列を取得する。
    'MsgBox htmlDoc.getElementsByClassName("license")(0).innerHTML
    Cells(1, 1).Value = htmlDoc.getElementsByClassName("license")(0).innerHTML
    
 
End Sub
