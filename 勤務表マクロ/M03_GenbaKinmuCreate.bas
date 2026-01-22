Attribute VB_Name = "M03_GenbaKinmuCreate"
Option Explicit

Public Sub P03_GenbaKinmuCreateMain()
   
    '------------------------------------------------------
    '// 初期処理
    '------------------------------------------------------
    
    '// エラーハンドリング
    On Error GoTo Err

    '// 時短設定
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '// プロシージャ名の設定
    Const PROCEDUR_NAME As String = "P03_GenbaKinmuCreateMain"
    
    '// 共通クラスを設定
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon

    '// 処理開始時間取得
    Dim tmStartTime As Double
        tmStartTime = Now()
        
    '// 処理開始ログ
    Call objCmn.WriteLog("", 0)
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start")

    
    '------------------------------------------------------
    '// 現場勤務表作成処理
    '------------------------------------------------------
    Call P03_001_GenbaKinmuCreate

    
    '------------------------------------------------------
    '// 後処理
    '------------------------------------------------------
    
    '// プロシージャ終了ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End")
 
    '// A1セル選択
    objCmn.Range1
    
    '// 時短解除
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    '// 処理終了時間の設定
    Dim tmEndTime As Double
        tmEndTime = Now()
        
    '// 正常終了メッセージの出力
    'MsgBox "【処理時間】" & Format(tmEndTime - tmStartTime, "hh:mm:ss") & vbNewLine & _
            "処理が正常終了しました。" & vbNewLine & _
            vbNewLine & _
            "日付、備考等の確認をして下さい。", vbInformation, "正常終了"

GoTo Finish


'------------------------------------------------------
'// エラーハンドリング
'------------------------------------------------------
Err:
    
    '// エラーログの出力
    Call objCmn.WriteLog("【プロシージャ名 】　" & strErrProcedureName)
    Call objCmn.WriteLog("【エラーNo       】　" & Err.Number, 3)
    Call objCmn.WriteLog("【エラー内容     】　" & Err.Description, 3)
    Call objCmn.WriteLog(PROCEDUR_NAME & "プロシージャ　End--")
    
    '// エラーメッセージ出力
    tmEndTime = Now()
    'MsgBox "処理が異常終了しました。" & vbNewLine & vbNewLine & _
            "【処理時間       】" & Format(tmEndTime - tmStartTime, "hh:mm:ss") & vbNewLine & _
            "【プロシージャ名 】　" & strErrProcedureName & vbNewLine & _
            "【エラーNo       】　" & Err.Number & vbNewLine & _
            "【エラー内容     】　" & Err.Description, vbCritical, "異常終了"
    
    '// A1セルを選択
    objCmn.Range1
    
    '// 時短解除
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

Finish:
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  GenbaKinmuCreate
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 引数① ：なし
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'// 　現場勤務表を作成
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub P03_001_GenbaKinmuCreate()

    '------------------------------------------------------
    '// 初期処理
    '------------------------------------------------------

    '// エラーハンドリング
    On Error GoTo Err

    '// プロシージャ名の設定
    Const PROCEDUR_NAME As String = "P03_001_GenbaKinmuCreate"
    
    '// 共通クラスの設定
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
        Call objCmn.GetInfo(ThisWorkbook.Worksheets(KINMU_SHEET))
    
    '// プロシージャ開始ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start")

    
    '------------------------------------------------------
    '// 現場勤務表の存在確認
    '------------------------------------------------------
    
    '// 勤務表が存在しない場合、処理中断
    If Dir(STR_GENBA_KINMU_PATH) = "" Then
        Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]現場勤務表のフォーマットが存在しない為、処理を中断します。", 2)
        Err.Raise 999, , "現場勤務表のフォーマットが存在しない為、処理を中断します。"
    End If
    
    
    '------------------------------------------------------
    '// アイトレ勤務表からデータを取得
    '------------------------------------------------------
    
    '// アイトレ勤務表データを格納する配列を作成
    Dim arrKinmuDate(4, LNG_END_ROW - LNG_START_ROW) As String
    
    '// アイトレ勤務表から各データを取得
    Call P03_002_GetValueKinmu(arrKinmuDate)
    
    
    '------------------------------------------------------
    '// 現場勤務表へアイトレ勤務表のデータを反映
    '------------------------------------------------------
    
    '// 現場勤務表の共通クラスを設定
    Dim objGenbaCmn As clsCommon
        Set objGenbaCmn = New clsCommon
    
    '// 現場勤務表の設定
    Dim wbGenbaKinmu As Workbook
        Set wbGenbaKinmu = Workbooks.Open(STR_GENBA_KINMU_PATH)
    
    '// 勤務表データを現場勤務表に反映
    Call P03_003_OutputValueKinmu(arrKinmuDate, wbGenbaKinmu)


    '------------------------------------------------------
    '// 現場勤務表の保存
    '------------------------------------------------------
    
    '// 対象月の取得
    Dim strRunDate As String
        strRunDate = ThisWorkbook.Worksheets("勤務表 打込み用 (IT)").Range(STR_MONTH).Value
    
    '//対象月にリネームしたPDFファイル名の作成
    Dim strOutputFileName As String                 '出力ファイル名
        strOutputFileName = Replace(STR_GENBA_KINMU_PATH, "yyyy", Format(strRunDate, "yyyy"))
        strOutputFileName = Replace(strOutputFileName, "MM", Format(strRunDate, "M"))
            
    '// 現場勤務表を対象月ファイル名に設定しExcel出力
    If IsWorkbookOpen(objCmn.GetFileNames(strOutputFileName)) Then
        Workbooks(objCmn.GetFileNames(strOutputFileName)).Close
    End If
    
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]現場勤務表保存処理")
    Application.DisplayAlerts = False
    wbGenbaKinmu.SaveAs strOutputFileName
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]現場勤務表を保存しました。：" & strOutputFileName)
    
    
    '------------------------------------------------------
    '// 後処理
    '------------------------------------------------------
    
    '// プロシージャ終了ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End")

GoTo Finish


'------------------------------------------------------
'// エラーハンドリング
'------------------------------------------------------
Err:

    '// エラー発生プロシージャ名の設定
    If strErrProcedureName = "" Then strErrProcedureName = PROCEDUR_NAME
    
    '// 発生したエラーを再発させ、呼び出し元プロシージャへ戻す
    Err.Raise Err.Number, , Err.Description
    
    GoTo Finish

Finish:
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  GetValueKinmu
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 引数① ：勤務表データを格納する配列
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'// アイトレ勤務表から各データを取得する。
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub P03_002_GetValueKinmu(arrKinmuDate)

    On Error GoTo Err

    '定数・変数宣言
    Const PROCEDUR_NAME As String = "P03_002_GetValueKinmu"
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
        Call objCmn.GetInfo(ThisWorkbook.Worksheets(KINMU_SHEET))

    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start")

    '// アイトレ勤務表から処理対象日付を取得
    arrKinmuDate(4, 0) = objCmn.WS.Range(STR_ITRE_DAY_CELL).Value
    
    '// 勤務表データの取得
    Dim lngArrCnt As Long                                                                       ' 配列カウント
    Dim lngLoopCnt As Long                                                                      ' ループカウント
    For lngLoopCnt = LNG_START_ROW To LNG_END_ROW
    
        arrKinmuDate(0, lngArrCnt) = _
            objCmn.WS.Cells(lngLoopCnt, KinmuCol.StartTime).Value                               ' 始業時間の取得
            
        arrKinmuDate(1, lngArrCnt) = _
            objCmn.WS.Cells(lngLoopCnt, KinmuCol.EndTime).Value                                 ' 終業時間の取得
        
        '//　休憩時間は1時間固定で入力
        If objCmn.WS.Cells(lngLoopCnt, KinmuCol.IntermMission).Value <> "" Then
            arrKinmuDate(2, lngArrCnt) = "1:00"                                                 ' 休憩時間の取得
        End If
        
        arrKinmuDate(3, lngArrCnt) = _
            objCmn.WS.Cells(lngLoopCnt, KinmuCol.ReMarks).Value                                 ' 備考の取得
            
            
        '// 配列カウントの更新
        lngArrCnt = lngArrCnt + 1
    
    Next lngLoopCnt

    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End")

GoTo Finish

Err:
    If strErrProcedureName = "" Then strErrProcedureName = PROCEDUR_NAME
    Err.Raise Err.Number, , Err.Description
    
    GoTo Finish

Finish:
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


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  OutputValueKinmu
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 引数1 ：勤務表データを格納されている配列
'// 引数2 ：現場勤務表のBookオブジェクト
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'// 現場勤務表にアイトレ勤務表のデータを反映する。
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub P03_003_OutputValueKinmu(arrKinmuDate, wbGenbaKinmu)

    On Error GoTo Err

    '定数・変数宣言
    Const PROCEDUR_NAME As String = "P03_003_OutputValueKinmu"
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
        Call objCmn.GetInfo(Workbooks(wbGenbaKinmu.Name).Worksheets(STR_GENBA_SHEET_NAME))

    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start")

    '// 現場勤務表に処理対象日付を反映
    objCmn.WS.Range(STR_GENBA_DAY_CELL).Value = arrKinmuDate(4, 0)
    
    '// 勤務表データの転記
    Dim lngArrCnt As Long                                                   ' 配列カウント
    Dim lngLoopCnt As Long                                                  ' ループカウント
    
    For lngLoopCnt = LNG_GENBA_START_ROW To LNG_GENBA_END_ROW
    
        objCmn.WS.Cells(lngLoopCnt, GenbaKinmuCol.StartTime).Value = _
            arrKinmuDate(0, lngArrCnt)                                      ' 始業時間の入力
        
        objCmn.WS.Cells(lngLoopCnt, GenbaKinmuCol.EndTime).Value = _
            arrKinmuDate(1, lngArrCnt)                                      ' 終業時間の入力
            
        objCmn.WS.Cells(lngLoopCnt, GenbaKinmuCol.IntermMission).Value = _
            arrKinmuDate(2, lngArrCnt)                                      ' 休憩時間の入力
            
        objCmn.WS.Cells(lngLoopCnt, GenbaKinmuCol.ReMarks).Value = _
            arrKinmuDate(3, lngArrCnt)                                      ' 備考の入力

        '// 配列カウントの更新
        lngArrCnt = lngArrCnt + 1
    
    Next lngLoopCnt

    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End")

GoTo Finish

Err:
    If strErrProcedureName = "" Then strErrProcedureName = PROCEDUR_NAME
    Err.Raise Err.Number, , Err.Description
    
    GoTo Finish

Finish:
End Sub
