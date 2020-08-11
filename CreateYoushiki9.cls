VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateYoushiki9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@folder("view")
Option Explicit

Implements ICreateYoushiki

'�Ώۃt�@�C���֌W
Private Const YOUSHIKI_EXCEL_NAME As String = "000639678.xls"
Private Const TARGET_BASE_SHEET_NAME = "�l������X���x�ƈꗗ�\�i�V�^�R���i�E�C���X�����Ǌ֌W�j"

Private m_kyugyo_workers As Workers '�x�ƑΏۘJ���҂��i�[����

Private m_wb���шꗗ�\ As Workbook
Private m_ws�ꗗ�\ As Worksheet
Private m_Company As Company

Private m_currentTargetWs As Worksheet '���ݓ��͑ΏۂɂȂ�V�[�g
Private m_createSheetCnt As Long '�쐬�����V�[�g������ێ�
Private m_mustCreateSheetNum As Long

'�o�͐�̍s�Ɨ�ԍ�
Private Enum col
    '���ӕ���
    �ʂ��ԍ� = 1
    �J�n��_�N = 3
    �J�n��_�� = 5
    �J�n��_�� = 7
    �I����_�N = 11
    �I����_�� = 13
    �I����_�� = 15
    �쐬��_�N = 2
    �쐬��_�� = 4
    �쐬��_�� = 6
    ��Ж� = 3
    ���Ə��ԍ�_�� = 6
    ���Ə��ԍ�_�� = 9
    ���Ə��ԍ�_�E = 12
    ���Ǝ厁�� = 3
    �J���ґ�\���� = 3
    ���� = 17
    ���� = 19
    '���Ώێҕ���
    �ΏێҖ� = 2
    �ی��ԍ�_�� = 8
    �ی��ԍ�_�� = 12
    �ی��ԍ�_�E = 17
    ��������� = 18
    ��\�I�ȏ���J������ = 18
    �S���x�� = 19
    �Z���ԋx�� = 20
    ����P�� = 21
    �Ώێґ��v = 21
End Enum

Private Enum row
    '���ӕ���
    ������� = 6
    ���v = 44
    ��\�I�ȏ���J������ = 46
    �쐬�� = 48
    ��Ж� = 49
    ���Ə��ԍ� = 50
    ���Ǝ厁�� = 51
    �J���ґ�\���� = 53
    ���� = 53
    �x��_����Ώێҍ��v = 49
    �x�Ƃ̂ݑΏێҍ��v = 50
    ����̂ݑΏێҍ��v = 51
    '���Ώێҕ���
    �ΏێҖ�_�J�n = 11
End Enum

Private Const �Ώێҍs�Ԋu = 3


Private Sub Class_Initialize()
    Set m_Company = New Company
    Set m_kyugyo_workers = New Workers
End Sub

Private Sub ICreateYoushiki_�����G�N�Z�����J��()
    Dim youshiki9 As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .AllowMultiSelect = False
        .Title = "�l��9����I�����ĉ�����"
        .InitialFileName = YOUSHIKI_EXCEL_NAME
        .Filters.Add "Excel", "*.xls?"
        If .Show Then
            youshiki9 = .SelectedItems(1)
        Else
            End
        End If
    End With
    
    '���炩���߈���͈͂�ݒ肵�AG:3��G:4�ɂ��遠�Ƀ`�F�b�N�����Ă����܂��傤�B
    Set m_wb���шꗗ�\ = Workbooks.Open(youshiki9)
    Set m_ws�ꗗ�\ = m_wb���шꗗ�\.Worksheets(TARGET_BASE_SHEET_NAME)
    ICreateYoushiki_�����t�@�C��������
End Sub

Private Sub ICreateYoushiki_�����t�@�C��������()
    If m_wb���шꗗ�\.Worksheets.Count > 1 Then
        Dim msg As String: msg = "�����ȊO�̃V�[�g�͑S�č폜���Ă���ꗗ�\���쐬���܂��B��낵���ł����H"
        If MsgBox(msg, vbExclamation + vbYesNo) = vbYes Then
            Application.DisplayAlerts = False
            Dim ws As Worksheet
            For Each ws In m_wb���шꗗ�\.Worksheets
                If ws.name <> TARGET_BASE_SHEET_NAME Then ws.Delete
            Next
            Application.DisplayAlerts = True
        Else
            End
        End If
    End If
    m_ws�ꗗ�\.Range("A11:A40").ClearContents
End Sub

Public Sub ICreateYoushiki_�ꗗ�쐬(company_ As Company, workers_ As Workers)
    Set m_Company = company_
    Set m_kyugyo_workers = workers_
    
    m_mustCreateSheetNum = WorksheetFunction.RoundUp(m_kyugyo_workers.�x��_����P���Ώێҍ��v, -1) / 10
    
    ICreateYoushiki_�����G�N�Z�����J��
    ICreateYoushiki_���V�[�g�R�s�[
    
    Dim outputRowOffset As Long
    Dim worker_cnt As Long
    For worker_cnt = 1 To m_kyugyo_workers.�x��_����P���Ώێҍ��v
        ICreateYoushiki_�Ώێҏo�� worker_cnt, outputRowOffset
        outputRowOffset = outputRowOffset + �Ώێҍs�Ԋu
        
        'kyugyo_workers�����[�v����10�l�P�ʂŃV�[�g�𑝂₷
        If outputRowOffset Mod 10 = 0 Then
            outputRowOffset = 0
            ICreateYoushiki_���V�[�g�R�s�[
            DoEvents
        End If
        
        If worker_cnt = m_kyugyo_workers.�x��_����P���Ώێҍ��v Then ICreateYoushiki_�ŏI�y�[�W�쐬
    Next
    MsgBox "���шꗗ�\�̍쐬���������܂����B", vbInformation
End Sub

Private Sub ICreateYoushiki_���V�[�g�R�s�[()
    m_createSheetCnt = m_wb���шꗗ�\.Sheets.Count
    
    m_ws�ꗗ�\.Copy After:=m_wb���шꗗ�\.Sheets(m_wb���шꗗ�\.Sheets.Count)
    Set m_currentTargetWs = ActiveSheet
    m_currentTargetWs.name = m_createSheetCnt
    
    ICreateYoushiki_�Ӎ쐬
End Sub

Private Sub ICreateYoushiki_�Ӎ쐬()
    With m_currentTargetWs
        .Cells(row.�������, col.�J�n��_�N) = Format(m_Company.�������_�J�n��, "e")
        .Cells(row.�������, col.�J�n��_��) = Format(m_Company.�������_�J�n��, "m")
        .Cells(row.�������, col.�J�n��_��) = Format(m_Company.�������_�J�n��, "d")
        .Cells(row.�������, col.�I����_�N) = Format(m_Company.�������_�I����, "e")
        .Cells(row.�������, col.�I����_��) = Format(m_Company.�������_�I����, "m")
        .Cells(row.�������, col.�I����_��) = Format(m_Company.�������_�I����, "d")
    End With
End Sub

Private Sub ICreateYoushiki_�ŏI�y�[�W�쐬()
    With m_currentTargetWs
        .Cells(row.���v, col.���������) = m_kyugyo_workers.���ԏ���J���������v
        .Cells(row.���v, col.�S���x��) = m_kyugyo_workers.�S���x�Ɠ������v
        .Cells(row.���v, col.�Z���ԋx��) = m_kyugyo_workers.�Z���ԋx�Ǝ��ԍ��v
        .Cells(row.���v, col.����P��) = m_kyugyo_workers.����P���������v
        .Cells(row.��\�I�ȏ���J������, col.��\�I�ȏ���J������) = m_Company.����J������
        
        .Cells(row.�쐬��, col.�쐬��_�N) = Format(m_Company.�쐬��, "e")
        .Cells(row.�쐬��, col.�쐬��_��) = Format(m_Company.�쐬��, "m")
        .Cells(row.�쐬��, col.�쐬��_��) = Format(m_Company.�쐬��, "d")
                
        .Cells(row.��Ж�, col.��Ж�) = "�i���́j " & m_Company.��Ж�
        .Cells(row.���Ə��ԍ�, col.���Ə��ԍ�_��) = m_Company.���Ə��ԍ�_��
        .Cells(row.���Ə��ԍ�, col.���Ə��ԍ�_��) = m_Company.���Ə��ԍ�_��
        .Cells(row.���Ə��ԍ�, col.���Ə��ԍ�_�E) = m_Company.���Ə��ԍ�_�E
        .Cells(row.���Ǝ厁��, col.���Ǝ厁��) = "�i�����j " & m_Company.���ƎҎ���
        .Cells(row.�J���ґ�\����, col.�J���ґ�\����) = "�i�����j " & m_Company.�J���ґ�\����
        
        .Cells(row.�x��_����Ώێҍ��v, col.�Ώێґ��v) = m_kyugyo_workers.�x��_����P���Ώێҍ��v
        .Cells(row.�x�Ƃ̂ݑΏێҍ��v, col.�Ώێґ��v) = m_kyugyo_workers.�x�Ƃ̂ݑΏێҍ��v
        .Cells(row.����̂ݑΏێҍ��v, col.�Ώێґ��v) = m_kyugyo_workers.����P���̂ݑΏێҍ��v
    End With
End Sub

Private Sub ICreateYoushiki_�Ώێҏo��(ByVal worker_index As Long, ByVal offset As Long)
    With m_currentTargetWs
        .Cells(row.�ΏێҖ�_�J�n + offset, col.�ʂ��ԍ�) = worker_index
        .Cells(row.�ΏێҖ�_�J�n + offset, col.�ΏێҖ�) = m_kyugyo_workers.����(worker_index)
        .Cells(row.�ΏێҖ�_�J�n + offset, col.�ΏێҖ�).Select
        .Cells(row.�ΏێҖ�_�J�n + offset, col.�ی��ԍ�_��) = m_kyugyo_workers.�ی��ԍ�_��(worker_index)
        .Cells(row.�ΏێҖ�_�J�n + offset, col.�ی��ԍ�_��) = m_kyugyo_workers.�ی��ԍ�_��(worker_index)
        .Cells(row.�ΏێҖ�_�J�n + offset, col.�ی��ԍ�_�E) = m_kyugyo_workers.�ی��ԍ�_�E(worker_index)
        .Cells(row.�ΏێҖ�_�J�n + offset, col.���������) = m_kyugyo_workers.�������(worker_index)
        .Cells(row.�ΏێҖ�_�J�n + offset, col.�S���x��) = m_kyugyo_workers.�S���x��(worker_index)
        .Cells(row.�ΏێҖ�_�J�n + offset, col.�Z���ԋx��) = m_kyugyo_workers.�Z���ԋx��(worker_index)
        .Cells(row.�ΏێҖ�_�J�n + offset, col.����P��) = m_kyugyo_workers.����P��(worker_index)
        
        '�����ڂ��o��
        .Cells(row.����, col.����) = m_createSheetCnt
        .Cells(row.����, col.����) = m_mustCreateSheetNum
    End With
End Sub