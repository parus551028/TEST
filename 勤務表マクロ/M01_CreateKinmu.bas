Attribute VB_Name = "M01_CreateKinmu"
Option Explicit

Sub P01_CreateKinmu()

    '------------------------------------------------------
    '// 初期処理
    '------------------------------------------------------
    
    '// エラーハンドリング
'    On Error GoTo Err

    '// 時短設定
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    '// プロシージャ名の設定
    Const PROCEDUR_NAME As String = "P01_CreateKinmu"
    
    '// 共通クラスの設定
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
    
    '// 処理開始時間の取得
    Dim tmStartTime As Double
        tmStartTime = Now()
    
    '// プロシージャ開始ログの出力
    Call objCmn.WriteLog("", 0)
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start")


    '------------------------------------------------------
    '// 勤務表PDF化処理
    '------------------------------------------------------
    Call P01_001_CreateKinmu


    '------------------------------------------------------
    '// 後処理
    '------------------------------------------------------
    
    '// プロシージャ終了ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End")
 
    '// 時短解除
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    '// 処理終了時間の取得
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
    
    '// エラーログの出力
    Call objCmn.WriteLog("【プロシージャ名 】　" & strErrProcedureName)
    Call objCmn.WriteLog("【エラーNo       】　" & Err.Number, 3)
    Call objCmn.WriteLog("【エラー内容     】　" & Err.Description, 3)
    Call objCmn.WriteLog(PROCEDUR_NAME & "プロシージャ　End--")
    
    '// エラーメッセージ出力
    tmEndTime = Now()
    MsgBox "処理が異常終了しました。" & vbNewLine & vbNewLine & _
            "【処理時間       】" & Format(tmEndTime - tmStartTime, "hh:mm:ss") & vbNewLine & _
            "【プロシージャ名 】　" & strErrProcedureName & vbNewLine & _
            "【エラーNo       】　" & Err.Number & vbNewLine & _
            "【エラー内容     】　" & Err.Description, vbCritical, "異常終了"
    
    Application.Calculation = xlCalculationAutomatic

End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//  M01_001_CreateKinmu
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 引数① ：なし
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// 概要
'//
'// 「M1」セルに入力した月の勤務表フォーマットを作成。
'// 「祝日」シートを参照にし、祝日も考慮し作成。
'//
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub P01_001_CreateKinmu()

    '------------------------------------------------------
    '// 初期処理
    '------------------------------------------------------
    
    '// エラーハンドリング
    On Error GoTo Err

    '// 時短設定
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    '// プロシージャ名の設定
    Const PROCEDUR_NAME As String = "P01_001_CreateKinmu"
    
    '// 共通クラスの設定
    Dim objCmn As clsCommon
        Set objCmn = New clsCommon
    
    '// 現場クラスの設定
    Dim objKinmu As ClsKinmu
        Set objKinmu = New ClsKinmu

    '// プロシージャ開始ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]Start")
    
    
    '------------------------------------------------------
    '// フォーマット作成処理開始確認
    '------------------------------------------------------
    
    '// 本処理開始確認
    Dim lngStartFlg As Long
        lngStartFlg = MsgBox("処理を開始しますか？" & vbNewLine & _
                            "処理を開始すると現在入力されているデータが全て初期化されます。", vbYesNo + vbInformation, "処理実行確認")
                     
    '// 「いいえ」が応答された場合、処理終了
    If lngStartFlg = 7 Then GoTo Finish
    
    
    '------------------------------------------------------
    '// 前月分の有休を反映得
    '------------------------------------------------------
    ThisWorkbook.Worksheets(KINMU_SHEET).Range(STR_LAST_YUKYU).Value = ThisWorkbook.Worksheets(KINMU_SHEET).Range(STR_REMAIND_YUKYU).Value
   
    
    '------------------------------------------------------
    '// 祝日・開始終了時間などの取得
    '------------------------------------------------------
    
    '// 勤務表のフォーマット作業
    Call objCmn.GetInfo(ThisWorkbook.Worksheets(KINMU_SHEET))
    Call objKinmu.KinmuInitialize(objCmn.WS)
    
    '// 処理対象ファイルログ出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]" & Format(objCmn.WS.Range(STR_MONTH), "yyyy年MM月分 処理開始"))
    
    
    '------------------------------------------------------
    '// 転記先の初期化
    '------------------------------------------------------
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]初期化処理")
    Dim rngTargetRange As Range
    Set rngTargetRange = Range(objCmn.WS.Cells(LNG_START_ROW, KinmuCol.Holiday), objCmn.WS.Cells(LNG_END_ROW, KinmuCol.Holiday))
        rngTargetRange = "-"
    Set rngTargetRange = Range(objCmn.WS.Cells(LNG_START_ROW, KinmuCol.StartTime), objCmn.WS.Cells(LNG_END_ROW, KinmuCol.StartTime))
        rngTargetRange = ""
    Set rngTargetRange = Range(objCmn.WS.Cells(LNG_START_ROW, KinmuCol.EndTime), objCmn.WS.Cells(LNG_END_ROW, KinmuCol.EndTime))
        rngTargetRange = ""
    Set rngTargetRange = Range(objCmn.WS.Cells(LNG_START_ROW, KinmuCol.IntermMission), objCmn.WS.Cells(LNG_END_ROW, KinmuCol.IntermMission))
        rngTargetRange = ""
    Set rngTargetRange = Range(objCmn.WS.Cells(LNG_START_ROW, KinmuCol.NightIntermMission), objCmn.WS.Cells(LNG_END_ROW, KinmuCol.NightIntermMission))
        rngTargetRange = ""
    Set rngTargetRange = Range(objCmn.WS.Cells(LNG_START_ROW, KinmuCol.ReMarks), objCmn.WS.Cells(LNG_END_ROW, KinmuCol.ReMarks))
        rngTargetRange = ""
    Set rngTargetRange = Nothing
    
    
    '------------------------------------------------------
    '// 処理月の月末日を求める
    '------------------------------------------------------
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]末日計算処理")
    
    '// 対象年月１日の翌月を取得
    Dim datNextMonth As Date
        datNextMonth = DateAdd("m", 1, objCmn.WS.Range(STR_MONTH))
    
    '// 末日を取得
    Dim datLastDay As Date
        datLastDay = DateAdd("d", -1, datNextMonth)
    
    '// 末日を取得（数値）
    Dim lngLastDay As Long
        lngLastDay = Format(datLastDay, "dd")
    
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]処理末日：" & lngLastDay & "日")
    
    
    '------------------------------------------------------
    '// フォーマット作成
    '------------------------------------------------------
    
    '// ループ処理最終行を取得
    Dim lngLastRow As Long
        lngLastRow = LNG_END_ROW - (31 - lngLastDay)
        
    '// 月末までループ
    Dim i As Long
    For i = LNG_START_ROW To lngLastRow
            
            
        '------------------------------------------------------
        '// 祝日判定フラグの更新
        '------------------------------------------------------
        
        '// 処理対象日付の取得
        Dim datTargetDay As Date
            datTargetDay = objCmn.WS.Cells(i, KinmuCol.Hiduke).Value
        
        '// 祝日の場合は「祝日」カラムに「祝」を転記
        If WorksheetFunction.CountIf(Worksheets("祝日").Range("A:A"), CDate(datTargetDay)) > 0 Then
            objCmn.WS.Cells(i, KinmuCol.Holiday).Value = "祝"
            Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]" & CDate(datTargetDay) & " 祝日")
        End If
        
        
        '------------------------------------------------------
        '// 休祝日の判断
        '------------------------------------------------------
        
        '// 祝日フラグInitialize
        Dim blnHoridayFlg As Boolean
            blnHoridayFlg = False
            
        '// 対象日が「土・日・祝」の場合は祝日フラグをTrueに更新
        If objCmn.WS.Cells(i, KinmuCol.Weekend).Value = "土" Or _
            objCmn.WS.Cells(i, KinmuCol.Weekend).Value = "日" Or _
            objCmn.WS.Cells(i, KinmuCol.Holiday).Value = "祝" Then
            blnHoridayFlg = True
        End If
        
        
        '------------------------------------------------------
        '// 始終業時間の転記
        '------------------------------------------------------
        
        '// 平日の場合、始終業時間などを打刻
        If Not blnHoridayFlg Then
            Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]" & objCmn.WS.Cells(i, KinmuCol.Hiduke).Value & " 平日")
            objCmn.WS.Cells(i, KinmuCol.StartTime).Value = objKinmu.StartTime
            objCmn.WS.Cells(i, KinmuCol.EndTime).Value = objKinmu.EndTime
            objCmn.WS.Cells(i, KinmuCol.IntermMission).Value = objKinmu.InterMission
            objCmn.WS.Cells(i, KinmuCol.NightIntermMission).Value = objKinmu.NightInterMission
        
        '// 祝日の場合、始終業時間などを空白
        Else
            Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]" & CDate(datTargetDay) & " 休祝日")
            objCmn.WS.Cells(i, KinmuCol.StartTime).Value = ""
            objCmn.WS.Cells(i, KinmuCol.EndTime).Value = ""
            objCmn.WS.Cells(i, KinmuCol.IntermMission).Value = ""
            objCmn.WS.Cells(i, KinmuCol.NightIntermMission).Value = ""
        End If
        
                    
    Next i
    
    
    '------------------------------------------------------
    '// 後処理
    '------------------------------------------------------
    
    '// プロシージャ終了ログの出力
    Call objCmn.WriteLog("[" & PROCEDUR_NAME & "]End")
    
    '// 時短解除
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

Exit Sub


'------------------------------------------------------
'// エラーハンドリング
'------------------------------------------------------
Err:

    '// エラー発生プロシージャ名の設定
    If strErrProcedureName = "" Then strErrProcedureName = PROCEDUR_NAME
    
    '// 発生したエラーを再出力させ、呼び出し元プロシージャへ戻す
    Err.Raise Err.Number, , Err.Description
    
    GoTo Finish

Finish:
End Sub

