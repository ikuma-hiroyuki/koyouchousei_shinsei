Attribute VB_Name = "Main"
'@folder("control")
Option Explicit

Sub 被保険者用()
    If MsgBox("被保険者用の一覧表を作成しますか？", vbInformation + vbYesNo) = vbYes Then
        Dim youshiki As ICreateYoushiki: Set youshiki = New CreateYoushiki9
        実績一覧表作成 youshiki, True
    End If
End Sub

Sub 被保険者以外用()
    If MsgBox("被保険者用「以外」の一覧表を作成しますか？", vbInformation + vbYesNo) = vbYes Then
        Dim youshiki As ICreateYoushiki: Set youshiki = New CreateYoushiki2
        実績一覧表作成 youshiki, False
    End If
End Sub

Sub 実績一覧表作成(youshiki As ICreateYoushiki, ByVal Is被保険者 As Boolean)
    Dim workers_ As Workers: Set workers_ = New Workers
    workers_.targetCsvPath = GetTargetCsvPath
    If workers_.targetCsvPath <> vbNullString Then
        workers_.GetWorker Is被保険者
        
        If workers_.休業_教育訓練対象者合計 > 0 Then
            Dim c As Company:    Set c = New Company
            c.GetCompanyInfo
            youshiki.一覧作成 c, workers_
        Else
            Dim str As String
            str = "対象者が一人もいません。" & vbNewLine & _
                  "CSVファイルの列の並びが正しくないかもしれません。" & vbNewLine & _
                  "あるいは、被保険者用のデータで被保険保険者「以外」ボタンを押した(or その逆)のかもしれません。"
            MsgBox str, vbExclamation
        End If
    End If
End Sub

Private Function GetTargetCsvPath() As String
    Dim fso As Object:    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    ChDir ThisWorkbook.Path
    On Error GoTo 0
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .InitialFileName = fso.GetParentFolderName(ThisWorkbook.FullName)
        .Filters.Add "CSV", "*.csv"
        .AllowMultiSelect = False
        .Title = "対象のCSVファイルを選択して下さい"
        If .Show Then
            GetTargetCsvPath = .SelectedItems(1)
        End If
    End With
End Function


