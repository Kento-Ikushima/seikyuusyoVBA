Attribute VB_Name = "Module1"
Option Explicit

Sub 請求書作成()
    Const HAN_HIDUKE_CLM As Long = 1 '「販売」ワークシートの「日付」の列
    Const HAN_KOKYAKU_CLM As Long = 2 '「販売」ワークシートの「顧客」の列
    Const HAN_SYOHIN_CLM As Long = 3 '「販売」ワークシートの「商品」の列
    Const HAN_TANKA_CLM As Long = 4 '「販売」ワークシートの「単価」の列
    Const HAN_SURYO_CLM As Long = 5 '「販売」ワークシートの「数量」の列
    Const HAN_KINGAKU_CLM As Long = 6 '「販売」ワークシートの「金額」の列
    
    Const SEI_HIDUKE_CLM As Long = 1 '請求書のワークシートの「日付」の列
    Const SEI_SYOHIN_CLM As Long = 2 '請求書のワークシートの「商品」の列
    Const SEI_TANKA_CLM As Long = 3 '請求書のワークシートの「単価」の列
    Const SEI_SURYO_CLM As Long = 4 '請求書のワークシートの「数量」の列
    Const SEI_KINGAKU_CLM As Long = 5 '請求書のワークシートの「金額」の列
    
    Const SEITP_WSNM As String = "請求書雛形" '請求書テンプレートのワークシート名
    Const ATESAKI_ADRS As String = "A6" '請求書の宛先のセル番地
    Const HAKKOBI_ADRS As String = "E2" '請求書の発行日のセル番地
    

    Dim i As Long '「販売」ワークシートの表の処理用カウント変数
    Dim Cnt As Long '請求書のワークシートの表の処理用変数
    Dim Kokyaku As String '請求書を作成する顧客名
    Dim HanKiten As Range '「販売」ワークシートの表の基点セル
    Dim SeiKiten As Range '請求書のワークシートの表の基点セル
    Dim sheetExists As Boolean '既存のシートが存在するかのフラグ
    Dim ws As Object 'ワークシートを走査するための変数
    
    
    Cnt = 1 '請求書のワークシートの表の先頭行の値に初期化
    Kokyaku = myForm.myComboBox.Value 'フォームのドロップダウンで選んだ顧客を設定
    sheetExists = False '存在しない、に初期化
  
    '既存のシートが存在するか確認
    For Each ws In Worksheets
        If ws.Name = Kokyaku Then
            sheetExists = True
            Exit For
        End If
    Next ws

    '既存のシートが存在する場合、メッセージを表示して終了
    If sheetExists Then
        MsgBox "その請求書はすでに発行済みです。"
        Exit Sub
    End If
    
    
    'ワークシート「請求書雛形」を末尾にコピー
    Worksheets("請求書雛形").Copy After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = Kokyaku      'ワークシート名設定
    Worksheets(Kokyaku).Range(ATESAKI_ADRS).Value = Kokyaku  '宛先の設定
    Worksheets(Kokyaku).Range(HAKKOBI_ADRS).Value = Date     '発行日の入力
    
    Set HanKiten = Worksheets("販売").Range("A4") '「販売」ワークシートの表の基点セルを設定
    Set SeiKiten = Worksheets(Kokyaku).Range("A12") '請求書のワークシートの表の基点セルを設定
    
    '指定した販売データを請求書へコピー
    For i = 1 To HanKiten.CurrentRegion.Rows.Count - 1
      If HanKiten.Cells(i, HAN_KOKYAKU_CLM).Value = Kokyaku Then
        SeiKiten.Cells(Cnt, SEI_HIDUKE_CLM).Value = HanKiten.Cells(i, HAN_HIDUKE_CLM).Value '日付
        SeiKiten.Cells(Cnt, SEI_SYOHIN_CLM).Value = HanKiten.Cells(i, HAN_SYOHIN_CLM).Value '商品
        SeiKiten.Cells(Cnt, SEI_TANKA_CLM).Value = HanKiten.Cells(i, HAN_TANKA_CLM).Value '単価
        SeiKiten.Cells(Cnt, SEI_SURYO_CLM).Value = HanKiten.Cells(i, HAN_SURYO_CLM).Value '数量
        SeiKiten.Cells(Cnt, SEI_KINGAKU_CLM).Value = HanKiten.Cells(, HAN_KINGAKU_CLM).Value '金額
        
        Cnt = Cnt + 1 '請求書のワークシートの表のコピー先を1つ進める
      End If
    Next
    
    'フォームをアンロードする
    Unload myForm

    'PDF作成の確認メッセージを表示
    Dim response As Integer
    response = MsgBox("続けてPDFを作成しますか？", vbQuestion + vbYesNo, "確認")
    
    'ユーザーの選択に応じて処理を実行
    If response = vbYes Then
        'PDF作成処理を実行する関数を呼び出す
        CreatePDF Kokyaku
    End If
End Sub

Sub CreatePDF(ByVal Kokyaku As String)
    ' 請求書ワークシートの範囲を選択
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Kokyaku)
    Dim lastRow As Long
    Const SEI_HIDUKE_CLM As Long = 1 '請求書のワークシートの「日付」の列
    Const SEI_KINGAKU_CLM As Long = 5 '請求書のワークシートの「金額」の列
    lastRow = ws.Cells(ws.Rows.Count, SEI_HIDUKE_CLM).End(xlUp).Row
    Dim printRange As Range
    Set printRange = ws.Range(ws.Cells(SEI_HIDUKE_CLM, 1), ws.Cells(lastRow, SEI_KINGAKU_CLM))
    
    ' PDFファイルを保存するパスを指定
    Dim tempFilePath As String
    tempFilePath = "C:\個人\生島\install\VBA勉強用\sample\TempPDF_" & Format(Now(), "yyyymmdd_hhmmss") & ".pdf"

    '「Microsoft Print to PDF」プリンターを設定
    Application.ActivePrinter = "Microsoft Print to PDF on Ne01:"
    printRange.ExportAsFixedFormat Type:=xlTypePDF, Filename:=tempFilePath, Quality:=xlQualityStandard
    
    ' PDF作成後、メッセージを表示
    MsgBox "PDF作成が完了しました。"
End Sub

Sub フォーム用意() 'ボタンに埋め込む用
    myForm.Show
End Sub
Sub 請求書削除()
    ' 確認メッセージを表示し、削除を選択した場合のみ削除処理を実行する
    Dim response As Integer
    response = MsgBox("作成した請求書を一括で削除しますか？", vbQuestion + vbYesNo, "確認")
    
    If response = vbYes Then
        Dim ws As Worksheet
        Application.DisplayAlerts = False ' 確認メッセージを非表示にする
        
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> "販売" And ws.Name <> "請求書雛形" And ws.Name <> "設定" Then
                ws.Delete
            End If
        Next ws
        
        Application.DisplayAlerts = True ' 確認メッセージを再表示する
        
        MsgBox "削除が完了しました。"
    End If
End Sub
'文字化けが治っているかの実験




