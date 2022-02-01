Option Explicit

Sub EachRowsPaste()

    Dim selectedRange As Range
    Set selectedRange = Application.InputBox _
                        ( _
                        prompt:="式を挿入する範囲をマウスで選択してください。", _
                        Title:="対象範囲選択", _
                        Type:=8 _
                        )

    '入力のエラー処理
    If Err.Number <> 0 Then
        MsgBox "キャンセルされました。"
        Exit Sub
    End If
    If selectedRange.Columns.Count <> 1 Then
        MsgBox "１列のみ選択してください。"
        Exit Sub
    End If

    '選択した範囲を配列に格納する → Rangeオブジェクトはもともとコレクションなのでやる必要ない
    'Dim selectedRangeColl As Collection: Set selectedRangeColl = New Collection

    'パスと実行するブック名を取得する
    Dim path As String: path = ThisWorkbook.path

    Dim thisBookName As String: thisBookName = ThisWorkbook.Name
    Dim prevName As String: prevName = left(thisBookName, InStr(thisBookName, "（"))
    Dim folName As String: folName = Mid(thisBookName, InStr(thisBookName, "）"))
    '回答元様式の拡張子を選ばせる
    Dim extension As String: extension = Application.InputBox _
                        ( _
                        prompt:="回収した様式の拡張子を選択してください。（例：xlsx）", _
                        Title:="拡張子選択", _
                        Type:=2 _
                        )
    folName = left(folName, InStr(folName, ".")) & extension

    '式を生成して選択範囲へ代入する
    Dim generatedFormula As String
    Dim i As Long
    For i = 1 To selectedRange.Count
        generatedFormula = "='" & path & "[" & prevName & selectedRange(i).Offset(0, -1).Value & folName & "]" & _
                           "Sheet1" & "'!" & selectedRange(i).Address(False, False)
        selectedRange(i).Value = generatedFormula

        MsgBox generatedFormula
    Next i

End Sub
