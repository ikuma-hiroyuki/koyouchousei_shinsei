Attribute VB_Name = "Main"
'@folder("control")
Option Explicit

Sub ��ی��җp()
    If MsgBox("��ی��җp�̈ꗗ�\���쐬���܂����H", vbInformation + vbYesNo) = vbYes Then
        Dim youshiki As ICreateYoushiki: Set youshiki = New CreateYoushiki9
        ���шꗗ�\�쐬 youshiki, True
    End If
End Sub

Sub ��ی��҈ȊO�p()
    If MsgBox("��ی��җp�u�ȊO�v�̈ꗗ�\���쐬���܂����H", vbInformation + vbYesNo) = vbYes Then
        Dim youshiki As ICreateYoushiki: Set youshiki = New CreateYoushiki2
        ���шꗗ�\�쐬 youshiki, False
    End If
End Sub

Sub ���шꗗ�\�쐬(youshiki As ICreateYoushiki, ByVal Is��ی��� As Boolean)
    Dim workers_ As Workers: Set workers_ = New Workers
    workers_.targetCsvPath = GetTargetCsvPath
    If workers_.targetCsvPath <> vbNullString Then
        workers_.GetWorker Is��ی���
        
        If workers_.�x��_����P���Ώێҍ��v > 0 Then
            Dim c As Company:    Set c = New Company
            c.GetCompanyInfo
            youshiki.�ꗗ�쐬 c, workers_
        Else
            Dim str As String
            str = "�Ώێ҂���l�����܂���B" & vbNewLine & _
                  "CSV�t�@�C���̗�̕��т��������Ȃ���������܂���B" & vbNewLine & _
                  "���邢�́A��ی��җp�̃f�[�^�Ŕ�ی��ی��ҁu�ȊO�v�{�^����������(or ���̋t)�̂�������܂���B"
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
        .Title = "�Ώۂ�CSV�t�@�C����I�����ĉ�����"
        If .Show Then
            GetTargetCsvPath = .SelectedItems(1)
        End If
    End With
End Function


