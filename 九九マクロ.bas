Attribute VB_Name = M99_Kuku
Option Explicit

' 九九表（1〜9）をアクティブシートへ出力するマクロ
Public Sub CreateKukuTable()
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long

    Set ws = ActiveSheet
    ws.Cells.Clear

    ws.Range(A1).Value = 九九表
    ws.Range(A1).Font.Bold = True
    ws.Range(A1).Font.Size = 14

    ' 列ヘッダー（1〜9）
    For j = 1 To 9
        ws.Cells(2, j + 1).Value = j
    Next j

    ' 行ヘッダー（1〜9）と掛け算結果
    For i = 1 To 9
        ws.Cells(i + 2, 1).Value = i
        For j = 1 To 9
            ws.Cells(i + 2, j + 1).Value = i  j
        Next j
    Next i

    ' 書式調整
    With ws.Range(A2J11)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ws.Columns(AJ).AutoFit

    MsgBox 九九表を作成しました。, vbInformation
End Sub
