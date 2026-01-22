Attribute VB_Name = "M02_PrintKinmu"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  P02_PrintKinmu
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'// 　勤務表をPDF出力する。
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub P02_PrintKinmu()
   
    '------------------------------------------------------
    '// 初期処理
    '------------------------------------------------------
    
    '// エラーハンドリング
    On Error GoTo Err

    '// 時短処理
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '// プロシージャ名の設定
    Const PROCEDUR_NAME As String = "P02_PrintKinmu"
    
    '// 共通クラスの設定
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
    
    '// 処理開始時間
    Dim tmStartTime As Double
        tmStartTime = Now()
    
    '// プロシージャ開始ログの出力
    Call objCmn.WriteLog("", 0)
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start")


    '------------------------------------------------------
    '// 勤務表PDF化処理
    '------------------------------------------------------
    Call P02_001_PDFCreate
    

    '------------------------------------------------------
    '// 後処理
    '------------------------------------------------------
    
    '// プロシージャ終了ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End")
 
    '// 時短解除
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    '// 処理終了時間
    Dim tmEndTime As Double
        tmEndTime = Now()
    
    '// 正常終了メッセージの出力
    'MsgBox "【処理時間】" & Format(tmEndTime - tmStartTime, "hh:mm:ss") & vbNewLine & _
            "処理が正常終了しました。", vbInformation, "正常終了"
Exit Sub
   
   
'------------------------------------------------------
'// エラーハンドリング
'------------------------------------------------------
Err:

    '// エラー発生プロシージャ名の設定
    If strErrProcedureName = "" Then strErrProcedureName = PROCEDUR_NAME
    
    '// 発生したエラーを再初させ、呼び出し元プロシージャへ戻す
    Err.Raise Err.Number, , Err.Description
    
    GoTo Finish

Finish:
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  PDFCreate
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 引数① ：なし
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'// 　勤務表をPDF化し、Configで設定したフォルダに出力する。
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub P02_001_PDFCreate()

    '------------------------------------------------------
    '// 初期処理
    '------------------------------------------------------

    '// エラーハンドリング
    On Error GoTo Err:
    
    '// 時短設定
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// プロシージャ名の設定
    Const PROCEDUR_NAME As String = "P02_001_PDFCreate"
    
    '// 共通クラスの設定
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
    
    '// プロシージャ開始ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start")
    
    
    '------------------------------------------------------
    '// 各情報の取得処理
    '------------------------------------------------------
    
    '// PDFファイルアウトプットパス
    Dim strOutputFilePath As String
        strOutputFilePath = ThisWorkbook.Worksheets("勤務表 打込み用 (IT)").Range(STR_OUTPUT_PATH).Value

    '// 対象月の取得
    Dim strRunDate As String
        strRunDate = ThisWorkbook.Worksheets("勤務表 打込み用 (IT)").Range(STR_MONTH).Value

    '// 出力ファイル名の設定
    Dim strOutputFileName As String
        strOutputFileName = Replace(STR_OUTPUT_FILE_NAME, "yyyy", Format(strRunDate, "yyyy"))
        strOutputFileName = Replace(strOutputFileName, "MM", Format(strRunDate, "MM"))
    
    
    '------------------------------------------------------
    '// 勤務表PDF出力処理
    '------------------------------------------------------
    ThisWorkbook.Worksheets(Array("勤務表 打込み用 (IT)", "電車運行表(定期)", "電車運行表", "車両運行表")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strOutputFilePath & "\" & strOutputFileName
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]PDF出力先：" & strOutputFilePath & "\" & strOutputFileName)
    
    
    '------------------------------------------------------
    '// 後処理
    '------------------------------------------------------
    
    '// 「勤務時間報告書(入力用)_大関」シートをアクティブ化
    ThisWorkbook.Worksheets("勤務表 打込み用 (IT)").Select
    
    '// プロシージャ終了ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End")

    '// 時短解除
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    GoTo Finish

Exit Sub

    
'------------------------------------------------------
'// エラーハンドリング
'------------------------------------------------------
Err:

    '// エラー発生プロシージャ名の設定
    If strErrProcedureName = "" Then strErrProcedureName = PROCEDUR_NAME
    
    '// 発生したエラーを再発させ、呼び出し元プロシージャへ戻す
    Err.Raise Err.Number, , Err.Description

Finish:
End Sub



