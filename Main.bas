Attribute VB_Name = "Main"
Dim jColumn As Long     '中学の身長列
Dim jCell As range      '中学の身長セル
Dim hCell As range      '高校の名前セル

Sub Main()
    '---変数群---
    Dim hname As Variant        '高校生の名前
    Dim resultSheet As String   '結果出力シートの名前
    
    Dim hlist() As Variant      '高校名簿の配列
    
    Dim row As Integer          '出力行数カウンタ
    
    '-----ここから処理開始-----
    '出力するシート名の格納
    resultSheet = "出力結果"
    
    'フォームを可視化する
    Form.Show
    
    '中学シートをアクティブにし、身長先頭セル、列インデックスを取得する
    Worksheets(Form.jSheet.Text).Activate
    On Error GoTo SelectError
    Set jCell = SelectCell("中学シートの身長先頭セル（項目名を除く）を選択してください。")
    jColumn = jCell.Column
    
    '高校シートをアクティブにし、名簿先頭セルを取得する
    Worksheets(Form.hSheet.Text).Activate
    On Error GoTo SelectError
    Set hCell = SelectCell("高校名簿の氏名先頭セル（項目名を除く）を選択してください。")
    
    '画面更新停止
    Application.ScreenUpdating = False
    
    '高校名簿の配列を取得する
    hlist() = gethNames()
    
    '結果出力シートをアクティブにする
    On Error GoTo Sheetrack
    
    '出力先頭行を格納する
    row = 2
    
    '各高校生徒名で検索メソッドを呼び出し、戻り値があったものについてシートに出力する
    For Each hname In hlist
        Dim std As Student
        '検索メソッド呼び出し
        Set std = getStudent(hname)
        
        If Not std Is Nothing Then
            '一列目に名前を格納
            Worksheets(resultSheet).Cells(row, 1).Value = std.name
            '二列目以降に身長体重を格納
            range(Worksheets(resultSheet).Cells(row, 2), Worksheets(resultSheet).Cells(row + 1, 10)).Value = std.getData
            '行数を更新
            row = row + 2
        End If
    Next
    '出力結果シートをアクティブにする
    Worksheets(resultSheet).Activate
    
    'フォームを閉じる
    Unload Form
    
    MsgBox "処理が完了しました。"
    Exit Sub
'セルが選択されなかった場合、処理を終了する
SelectError:
    MsgBox "セルの取得に失敗しました。"
    End
'出力シートがない場合、作成する
Sheetrack:
    Worksheets.Add.name = resultSheet
    Worksheets(resultSheet).range("A1:J1").Value = Array("氏名", "小1", "小2", "小3", "小4", "小5", "小6", "中1", "中2", "中3")
    Resume
End Sub
'セルを選択させ、その左上のセルを返却する
Function SelectCell(ByVal txt As String) As range
    Dim rng As range
    Set rng = Application.InputBox(txt, "セル選択", Type:=8)
    Set SelectCell = rng.Cells(1)
End Function
'高校名簿の全名前を配列で取得する
Function gethNames() As Variant()
    Dim nameCells As range
    Dim arr() As Variant
    
    Set nameCells = range(hCell, hCell.End(xlDown))
    arr() = nameCells.Value
    gethNames = arr()
End Function
'中学シートの検索を行い、名前があればStudentオブジェクトを生成、なければNothingを格納
Function getStudent(name As Variant) As Student
    Dim rng As range
    
    '中学シートを引数で検索する
    Set rng = Worksheets(Form.jSheet.Text).Cells.Find(name, Cells(1, 1), xlValues, xlWhole, xlByColumns, xlNext, False, False, False)
    
    '検索結果があった場合、Studentオブジェクトを生成する
    If Not rng Is Nothing Then
        Dim i As Integer    'カウンタ
        Dim row As Long     '行値
        Dim dataCell As range   '身長体重の先頭セル
        Dim person As Student   '返却するオブジェクト
        Dim arr(17) As Variant  '身長体重配列
        Dim datas() As Variant  'データ取得用配列
        
        'データの行値を取得
        row = rng.row
        
        '身長体重の先頭セルを取得
        Set dataCell = Worksheets(Form.jSheet.Text).Cells(row, jColumn)
        
        '範囲を中3まで拡大し、値を取得
        datas = dataCell.Resize(1, 18).Value
        
        '値を1次元配列に変換する
        For i = 0 To 17
            arr(i) = datas(1, i + 1)
        Next
        
        'Studentオブジェクトを生成して、名前と身長体重を格納する
        Set person = New Student
        person.name = rng.Value
        Call person.setData(arr)
        
        'オブジェクトの返却
        Set getStudent = person
    Else
        '検索結果がない場合、Nothingを返却
        Set getStudent = Nothing
    End If
End Function
