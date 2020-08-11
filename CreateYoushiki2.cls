VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateYoushiki2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@folder("view")
Option Explicit

Implements ICreateYoushiki

'対象ファイル関係
Private Const YOUSHIKI_EXCEL_NAME As String = "000639683.xls"
Private Const TARGET_BASE_SHEET_NAME = "様式新第２号（3）"

Private m_kyugyo_workers As Workers '休業対象労働者を格納する

Private m_wb実績一覧表 As Workbook
Private m_ws一覧表 As Worksheet
Private m_Company As Company

Private m_currentTargetWs As Worksheet '現在入力対象になるシート
Private m_createSheetCnt As Long '作成したシート枚数を保持
Private m_mustCreateSheetNum As Long

'出力先の行と列番号
Private Enum col
    '↓鑑部分
    通し番号 = 1
    開始日_年 = 3
    開始日_月 = 5
    開始日_日 = 7
    終了日_年 = 11
    終了日_月 = 13
    終了日_日 = 15
    作成日_年 = 2
    作成日_月 = 4
    作成日_日 = 6
    会社名 = 3
    事業所番号_左 = 5
    事業所番号_中 = 7
    事業所番号_右 = 10
    事業主氏名 = 3
    労働者代表氏名 = 3
    枚目 = 17
    枚中 = 19
    '↓対象者部分
    対象者名 = 2
    保険番号_左 = 8
    保険番号_中 = 12
    保険番号_右 = 17
    月所定日数 = 18
    代表的な所定労働時間 = 18
    全日休業 = 19
    短時間休業 = 20
    対象者総計 = 19
End Enum

Private Enum row
    '↓鑑部分
    判定期間 = 6
    合計 = 45
    代表的な所定労働時間 = 48
    作成日 = 53
    会社名 = 54
    事業所番号 = 55
    事業主氏名 = 56
    労働者代表氏名 = 58
    枚数 = 58
    休業対象者合計 = 54
    '↓対象者部分
    対象者名_開始 = 11
End Enum

Private Const 対象者行間隔 = 3


Private Sub Class_Initialize()
    Set m_Company = New Company
    Set m_kyugyo_workers = New Workers
End Sub

Private Sub ICreateYoushiki_原紙エクセルを開く()
    Dim youshiki2 As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .AllowMultiSelect = False
        .Title = "様式2号を選択して下さい"
        .InitialFileName = YOUSHIKI_EXCEL_NAME
        .Filters.Add "Excel", "*.xls?"
        If .Show Then
            youshiki2 = .SelectedItems(1)
        Else
            End
        End If
    End With
    
    'あらかじめ印刷範囲を設定し、G:3とG:4にある□にチェックを入れておきましょう。
    Set m_wb実績一覧表 = Workbooks.Open(youshiki2)
    Set m_ws一覧表 = m_wb実績一覧表.Worksheets(TARGET_BASE_SHEET_NAME)
    ICreateYoushiki_原紙ファイル初期化
End Sub

Private Sub ICreateYoushiki_原紙ファイル初期化()
    If m_wb実績一覧表.Worksheets.Count > 1 Then
        Dim msg As String: msg = "原紙以外のシートは全て削除してから一覧表を作成します。よろしいですか？"
        If MsgBox(msg, vbExclamation + vbYesNo) = vbYes Then
            Application.DisplayAlerts = False
            Dim ws As Worksheet
            For Each ws In m_wb実績一覧表.Worksheets
                If ws.name <> TARGET_BASE_SHEET_NAME Then ws.Delete
            Next
            Application.DisplayAlerts = True
        Else
            End
        End If
    End If
    m_ws一覧表.Range("A11:A40").ClearContents
End Sub

Public Sub ICreateYoushiki_一覧作成(company_ As Company, workers_ As Workers)
    Set m_Company = company_
    Set m_kyugyo_workers = workers_

    m_mustCreateSheetNum = WorksheetFunction.RoundUp(m_kyugyo_workers.休業_教育訓練対象者合計, -1) / 10

    ICreateYoushiki_原紙エクセルを開く
    ICreateYoushiki_元シートコピー
    
    Dim outputRowOffset As Long
    Dim worker_cnt As Long
    For worker_cnt = 1 To m_kyugyo_workers.休業_教育訓練対象者合計
        ICreateYoushiki_対象者出力 worker_cnt, outputRowOffset
        outputRowOffset = outputRowOffset + 対象者行間隔
        
        'kyugyo_workersをループして10人単位でシートを増やす
        If outputRowOffset Mod 10 = 0 Then
            outputRowOffset = 0
            ICreateYoushiki_元シートコピー
            DoEvents
        End If
        
        If worker_cnt = m_kyugyo_workers.休業_教育訓練対象者合計 Then ICreateYoushiki_最終ページ作成
    Next
    MsgBox "実績一覧表の作成が完了しました。", vbInformation
End Sub

Private Sub ICreateYoushiki_元シートコピー()
    m_createSheetCnt = m_wb実績一覧表.Sheets.Count
    
    m_ws一覧表.Copy After:=m_wb実績一覧表.Sheets(m_wb実績一覧表.Sheets.Count)
    Set m_currentTargetWs = ActiveSheet
    m_currentTargetWs.name = m_createSheetCnt
    
    ICreateYoushiki_鑑作成
End Sub

Private Sub ICreateYoushiki_鑑作成()
    With m_currentTargetWs
        .Cells(row.判定期間, col.開始日_年) = Format(m_Company.判定期間_開始日, "e")
        .Cells(row.判定期間, col.開始日_月) = Format(m_Company.判定期間_開始日, "m")
        .Cells(row.判定期間, col.開始日_日) = Format(m_Company.判定期間_開始日, "d")
        .Cells(row.判定期間, col.終了日_年) = Format(m_Company.判定期間_終了日, "e")
        .Cells(row.判定期間, col.終了日_月) = Format(m_Company.判定期間_終了日, "m")
        .Cells(row.判定期間, col.終了日_日) = Format(m_Company.判定期間_終了日, "d")
    End With
End Sub

Private Sub ICreateYoushiki_最終ページ作成()
    With m_currentTargetWs
        .Cells(row.合計, col.月所定日数) = m_kyugyo_workers.月間所定労働日数合計
        .Cells(row.合計, col.全日休業) = m_kyugyo_workers.全日休業日数合計
        .Cells(row.合計, col.短時間休業) = m_kyugyo_workers.短時間休業時間合計
        .Cells(row.代表的な所定労働時間, col.代表的な所定労働時間) = m_Company.所定労働時間
        
        .Cells(row.作成日, col.作成日_年) = Format(m_Company.作成日, "e")
        .Cells(row.作成日, col.作成日_月) = Format(m_Company.作成日, "m")
        .Cells(row.作成日, col.作成日_日) = Format(m_Company.作成日, "d")
                
        .Cells(row.会社名, col.会社名) = "（名称） " & m_Company.会社名
        .Cells(row.事業所番号, col.事業所番号_左) = m_Company.事業所番号_左
        .Cells(row.事業所番号, col.事業所番号_中) = m_Company.事業所番号_中
        .Cells(row.事業所番号, col.事業所番号_右) = m_Company.事業所番号_右
        .Cells(row.事業主氏名, col.事業主氏名) = "（氏名） " & m_Company.事業者氏名
        .Cells(row.労働者代表氏名, col.労働者代表氏名) = "（氏名） " & m_Company.労働者代表氏名
        
        .Cells(row.休業対象者合計, col.対象者総計) = m_kyugyo_workers.休業_教育訓練対象者合計
    End With
End Sub

Private Sub ICreateYoushiki_対象者出力(ByVal worker_index As Long, ByVal offset As Long)
    With m_currentTargetWs
        .Cells(row.対象者名_開始 + offset, col.通し番号) = worker_index
        .Cells(row.対象者名_開始 + offset, col.対象者名) = m_kyugyo_workers.氏名(worker_index)
        .Cells(row.対象者名_開始 + offset, col.対象者名).Select
        .Cells(row.対象者名_開始 + offset, col.保険番号_左) = m_kyugyo_workers.保険番号_左(worker_index)
        .Cells(row.対象者名_開始 + offset, col.保険番号_中) = m_kyugyo_workers.保険番号_中(worker_index)
        .Cells(row.対象者名_開始 + offset, col.保険番号_右) = m_kyugyo_workers.保険番号_右(worker_index)
        .Cells(row.対象者名_開始 + offset, col.月所定日数) = m_kyugyo_workers.所定日数(worker_index)
        .Cells(row.対象者名_開始 + offset, col.全日休業) = m_kyugyo_workers.全日休業(worker_index)
        .Cells(row.対象者名_開始 + offset, col.短時間休業) = m_kyugyo_workers.短時間休業(worker_index)
        
        '何枚目か出力
        .Cells(row.枚数, col.枚目) = m_createSheetCnt
        .Cells(row.枚数, col.枚中) = m_mustCreateSheetNum
    End With
End Sub
