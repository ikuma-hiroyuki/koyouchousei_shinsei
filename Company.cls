VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@folder("model")
Option Explicit

Private m_�������_�J�n�� As Date
Private m_�������_�I���� As Date
Private m_����J������ As Double
Private m_�쐬�� As Date
Private m_��Ж� As String
Private m_���Ə��ԍ� As String
Private m_���Ə��ԍ�_�� As String
Private m_���Ə��ԍ�_�� As String
Private m_���Ə��ԍ�_�E As String
Private m_���ƎҎ��� As String
Private m_�J���ґ�\���� As String

Private m_ws As Worksheet

Private Sub Class_Initialize()
    Set m_ws = ThisWorkbook.Worksheets("�S�̐ݒ�")
End Sub

Public Sub GetCompanyInfo()
    m_�������_�J�n�� = InputCheck(m_ws.Range("B2"))
    m_�������_�I���� = InputCheck(m_ws.Range("B3"))
    m_����J������ = WorksheetFunction.RoundDown(InputCheck(m_ws.Range("B5")), 0)
    m_�쐬�� = InputCheck(m_ws.Range("B8"))
    m_��Ж� = InputCheck(m_ws.Range("B10"))
    Me.���Ə��ԍ� = m_ws.Range("B11")
    m_���ƎҎ��� = InputCheck(m_ws.Range("B12"))
    m_�J���ґ�\���� = InputCheck(m_ws.Range("B13"))
End Sub

Public Property Get �������_�J�n��() As Date
    �������_�J�n�� = m_�������_�J�n��
End Property

Public Property Get �������_�I����() As Date
    �������_�I���� = m_�������_�I����
End Property

Public Property Get ����J������() As Double
    ����J������ = m_����J������
End Property

Public Property Get �쐬��() As Date
    �쐬�� = m_�쐬��
End Property

Public Property Get ��Ж�() As String
    ��Ж� = m_��Ж�
End Property

Public Property Get ���Ə��ԍ�() As String
    ���Ə��ԍ� = m_���Ə��ԍ�
End Property

Public Property Get ���Ə��ԍ�_��() As String
    ���Ə��ԍ�_�� = m_���Ə��ԍ�_��
End Property

Public Property Get ���Ə��ԍ�_��() As String
    ���Ə��ԍ�_�� = m_���Ə��ԍ�_��
End Property

Public Property Get ���Ə��ԍ�_�E() As String
    ���Ə��ԍ�_�E = m_���Ə��ԍ�_�E
End Property

Public Property Let ���Ə��ԍ�(ByVal value As String)
    value = StrConv(value, vbNarrow)
    If ���Ə�_�ی��Ҕԍ��`�F�b�N(value) Then
        m_���Ə��ԍ� = value
        �ی�_���Ə��ԍ����� m_���Ə��ԍ�, m_���Ə��ԍ�_��, m_���Ə��ԍ�_��, m_���Ə��ԍ�_�E
    Else
        MsgBox "�����Ҕԍ�������������܂���B(��������:1234-123456-1)" & vbNewLine & _
            "�C�����čēx���s���ĉ�����", vbCritical, "���̓~�X�H"
        End
    End If
End Property

Public Property Get ���ƎҎ���() As String
    ���ƎҎ��� = m_���ƎҎ���
End Property

Public Property Get �J���ґ�\����() As String
    �J���ґ�\���� = m_�J���ґ�\����
End Property

Private Function InputCheck(ByRef val As Variant) As Variant
    '�����͂�������G���[��\������B���͍ς������炻�̂܂ܒl��Ԃ��B
    If val = 0 Or val = vbNullString Then
        MsgBox "�S�̐ݒ�ɖ����͍��ڂ�����܂��B" & vbNewLine & "���͂��čēx���s���ĉ������B", vbCritical
        End
    Else
        InputCheck = val
    End If
End Function