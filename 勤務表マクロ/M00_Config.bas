Attribute VB_Name = "M00_Config"
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//
'//　全処理共通
'//
'//
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public strErrProcedureName As String                                    ' エラー発生時のプロシージャ名

Public Const STR_MONTH As String = "M1"                                 ' アイトレ勤務表の処理月入力セル
Public Const STR_LAST_YUKYU As String = "AJ14"                          ' アイトレ勤務表の先月分有給セル
Public Const STR_REMAIND_YUKYU As String = "AJ16"                       ' アイトレ勤務表の有休残数セル

Public Const LNG_START_ROW As Long = 7                                  ' アイトレ勤務表の勤務表の入力開始行
Public Const LNG_END_ROW As Long = 37                                   ' アイトレ勤務表の勤務表の入力終了行
Public Const KINMU_SHEET As String = "勤務表 打込み用 (IT)"             ' アイトレ勤務表の勤務表入力シート名

'// アイトレ勤務表のカラム数
Enum KinmuCol
    Hiduke = 1                                                          ' 日付カラム
    Weekend = 3                                                         ' 週カラム
    Holiday                                                             ' 休日カラム
    StartTime                                                           ' 始業時間カラム
    EndTime = 9                                                         ' 終業時間カラム
    IntermMission = 12                                                  ' 通常休憩時間カラム
    NightIntermMission = 14                                             ' 深夜休憩時間カラム
    ReMarks = 26                                                        ' 備考カラム
End Enum



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//
'//　M02_PrintKinmu
'//
'//
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Const STR_OUTPUT_PATH As String = "AJ9"                                      ' PDFアウトプットパス記載セル
Public Const STR_OUTPUT_FILE_NAME As String = "（アイトレ用）【大関洋平】yyyy年MM月_勤務表.pdf"   ' PDFアウトプットファイル名


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//
'//　M03_GenbaKinmuCreate
'//
'//
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Const STR_GENBA_KINMU_PATH As String = "Z:\000_マクロ\005_書類送付\02_勤務表\yyyy年MM月勤怠.xlsx"                                ' 現場勤務表のフルパス
Public Const LNG_GENBA_START_ROW As Long = 5                                                                                            ' 現場勤務表の開始行
Public Const LNG_GENBA_END_ROW As Long = 35                                                                                             ' 現場勤務表のフルパス
Public Const STR_GENBA_SHEET_NAME As String = "大関"                                                                                    ' 現場勤務表のシート名
Public Const STR_GENBA_DAY_CELL As String = "B2"                                                                                        ' 現場勤務表の処理対象月記載セル
Public Const STR_ITRE_DAY_CELL As String = "M1"                                                                                         ' アイトレ勤務表の処理対象月記載セル

'// 現場勤務表のカラム数
Enum GenbaKinmuCol
    StartTime = 4                                                                                                                       ' 始業時間カラム
    EndTime                                                                                                                             ' 終業時間カラム
    IntermMission                                                                                                                       ' 通常休憩時間カラム
    ReMarks = 8                                                                                                                         ' 備考カラム
End Enum

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//
'//　M04_PrintGenbaKinmu
'//
'//
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Const STR_OUTPUT_GENBA_PDF_NAME As String = "（ITSO様向け）【大関洋平】yyyy年MM月_勤務表.pdf"                                    ' PDFアウトプットファイル名

