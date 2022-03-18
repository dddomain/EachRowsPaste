Option Explicit

Sub EachRowsPaste()

Dim selectedRange As Range
Set selectedRange = Application.InputBox _
     ( _
    prompt:="式を挿入する範囲をマウスで選択してください。", _
    Title:="対象範囲選択", _
    Type:=8 _
    )

If Err.Number <> 0 Then
    MsgBox "キャンセルされました。"
    Exit Sub
End If
If selectedRange.Columns.Count <> 1 Then
    MsgBox "１列のみ選択してください。"
    Exit Sub
End If

Dim path As String: path = ThisWorkbook.path
Dim thisBookName As String: thisBookName = ThisWorkbook.Name
Dim prevName As String: prevName = left(thisBookName, InStr(thisBookName, "（"))
Dim folName As String: folName = Mid(thisBookName, InStr(thisBookName, "）"))
Dim extension As String: extension = Application.InputBox _
    ( _
    prompt:="回収した様式の拡張子を選択してください。（例：xlsx）", _
    Title:="拡張子選択", _
    Default:="xlsx", _
    Type:=2 _
    )
folName = left(folName, InStr(folName, ".")) & extension

Dim sheetName As String: sheetName = Application.InputBox _
    ( _
    prompt:="回答元ブックのシート名を入力してください。", _
    Title:="シート名入力", _
    Default:=ActiveSheet.Name, _
    Type:=2 _
    )

Dim targetWorkbook As Workbook
Dim targetBookName As String
Dim fullPath As String
Dim uncaughtBooks As Collection: Set uncaughtBooks = New Collection

Dim generatedFormula As String

Dim i As Long
For i = 1 To selectedRange.Count
    targetBookName = prevName & selectedRange(i).Offset(0, -1).Value & folName
    fullPath = path & "/" & targetBookName
    If Dir(fullPath) = "" Then
        uncaughtBooks.Add targetBookName
        GoTo Continue
    End If
    Set targetWorkbook = Workbooks.Open(fullPath)
    generatedFormula = "='[" & targetBookName & "]" & sheetName & "'!" & selectedRange(i).Address(False, False)
    selectedRange(i).Value = generatedFormula
    Workbooks(targetBookName).Close
Continue:
Next i

If uncaughtBooks.Count > 0 Then
    Dim alertMessage As String
    Dim uncaughtBook As Variant
    For Each uncaughtBook In uncaughtBooks
        alertMessage = alertMessage & uncaughtBook & vbCrLf
    Next
    MsgBox alertMessage & vbCrLf & uncaughtBooks.Count & "件はブックが見つかりませんでした。"
End If
End Sub
