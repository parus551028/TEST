Attribute VB_Name = "M04_PrintGenbaKinmu"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  P04_PrintGenbaKinmu
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'// 現場勤務表をPDF化し、その結果を出力する。対象の現場勤務表は指定されたパスから取得し、
'// 処理が完了したらそのワークブックを閉じる。処理時間をメッセージボックスで表示する。
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub P04_PrintGenbaKinmu()
   
    '------------------------------------------------------
    '// 初期処理
    '------------------------------------------------------
    On Error GoTo Err

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Const PROCEDUR_NAME As String = "P04_PrintGenbaKinmu"
    
    Dim lngProcedureLen As Long
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
    Dim tmStartTime As Double                          '処理開始時間
    Dim tmEndTime As Double                            '処理終了時間
    
    tmStartTime = Now()                                '処理開始時間
    
    Call objCmn.WriteLog("", 0)
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start ")

    '------------------------------------------------------
    '// 対象月の取得
    '------------------------------------------------------
    Dim strRunDate As String
    strRunDate = ThisWorkbook.Worksheets("勤務表 打込み用 (IT)").Range(STR_MONTH).Value


    '------------------------------------------------------
    '// 処理対象の現場勤務表のフルパス
    '------------------------------------------------------
    Dim strOutputFileName As String                 '出力ファイル名
    strOutputFileName = Replace(STR_GENBA_KINMU_PATH, "yyyy", Format(strRunDate, "yyyy"))
    strOutputFileName = Replace(strOutputFileName, "MM", Format(strRunDate, "M"))
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]処理対象ファイル：" & strOutputFileName)
    
    
    '------------------------------------------------------
    '// 現場勤務表をオープン
    '------------------------------------------------------
    If IsWorkbookOpen(objCmn.GetFileNames(strOutputFileName)) Then
        Workbooks(objCmn.GetFileNames(strOutputFileName)).Close
    End If
    
    Dim wbGenbaKinmu As Workbook
        Set wbGenbaKinmu = Workbooks.Open(strOutputFileName)
        Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]現場勤務表を開きました。：" & strOutputFileName)
    Dim objGenbaCmn As clsCommon
        Set objGenbaCmn = New clsCommon
        Call objCmn.GetInfo(wbGenbaKinmu.Worksheets(STR_GENBA_SHEET_NAME))


    '------------------------------------------------------
    '// 現場勤務表のPDF化
    '------------------------------------------------------
    Call P04_001_GenbaPDFCreate(wbGenbaKinmu)
    
    
    '------------------------------------------------------
    '// 現場勤務表のクローズ
    '------------------------------------------------------
    Workbooks(wbGenbaKinmu.Name).Close True
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]現場勤務表を閉じました。")


    '------------------------------------------------------
    '// 後処理
    '------------------------------------------------------
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End ")
 
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    tmEndTime = Now()                                '処理終了時間
    'MsgBox "【処理時間】" & Format(tmEndTime - tmStartTime, "hh:mm:ss") & vbNewLine & _
            "処理が正常終了しました。", vbInformation, "正常終了"
Exit Sub


'------------------------------------------------------
'// エラーハンドリング
'------------------------------------------------------
Err:
    
    '// エラーログの出力
    Call objCmn.WriteLog("【プロシージャ名 】　" & strErrProcedureName)
    Call objCmn.WriteLog("【エラーNo       】　" & Err.Number, 3)
    Call objCmn.WriteLog("【エラー内容     】　" & Err.Description, 3)
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End ")
    
    '// エラーメッセージ出力
    tmEndTime = Now()
    'MsgBox "処理が異常終了しました。" & vbNewLine & vbNewLine & _
            "【処理時間       】" & Format(tmEndTime - tmStartTime, "hh:mm:ss") & vbNewLine & _
            "【プロシージャ名 】　" & strErrProcedureName & vbNewLine & _
            "【エラーNo       】　" & Err.Number & vbNewLine & _
            "【エラー内容     】　" & Err.Description, vbCritical, "異常終了"
            
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  P04_001_GenbaPDFCreate
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 引数① ： 現場勤務表のWorkbookオブジェクト
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'// 現場勤務表をPDF化し、所定のフォルダへ出力する。
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub P04_001_GenbaPDFCreate(wbGenbaKinmu)

    '------------------------------------------------------
    '// エラー発生時、Errラベルへ
    '------------------------------------------------------
    On Error GoTo Err:
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// 定数・変数宣言
    Const PROCEDUR_NAME As String = "P04_001_GenbaPDFCreate"
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start ")
    
    '// PDFファイルの出力先パスの取得
    Dim strOutputFilePath As String               'PDFファイルアウトプットパス
    strOutputFilePath = ThisWorkbook.Worksheets("勤務表 打込み用 (IT)").Range(STR_OUTPUT_PATH).Value
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]現場勤務表出力パス：" & strOutputFilePath)
    
    '// 対象月の取得
    Dim strRunDate As String
    strRunDate = ThisWorkbook.Worksheets("勤務表 打込み用 (IT)").Range(STR_MONTH).Value
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]" & strRunDate & "分の処理を行います")

    '//対象月にリネームしたPDFファイル名の作成
    Dim strOutputFileName As String                 '出力ファイル名
    strOutputFileName = Replace(STR_OUTPUT_GENBA_PDF_NAME, "yyyy", Format(strRunDate, "yyyy"))
    strOutputFileName = Replace(strOutputFileName, "MM", Format(strRunDate, "MM"))
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]出力先パス：" & strOutputFilePath & "\" & strOutputFileName)
    
    'PDF出力
    Workbooks(wbGenbaKinmu.Name).Worksheets(STR_GENBA_SHEET_NAME).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strOutputFilePath & "\" & strOutputFileName
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]現場勤務表の出力が完了しました。")

    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End ")

Exit Sub

    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'エラー処理
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Err:
    Err.Raise Err.Number, , "【" & PROCEDUR_NAME & "】" & Err.Description

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

End Sub




'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  IsWorkbookOpen
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 引数1 ：確認したいワークブックの名前（文字列）
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'// 指定したワークブックが開かれているかどうかを確認する。ワークブックが開かれている場合はTrueを、そうでない場合はFalseを返す。
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function IsWorkbookOpen(WorkbookName As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name = WorkbookName Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
    IsWorkbookOpen = False
End Function

