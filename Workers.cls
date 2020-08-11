VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Workers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@folder("model")
Option Explicit
Option Base 1

Private Type kyugo_workers
    ���� As String
    �ی��ԍ� As String
    �ی��ԍ�_�� As String
    �ی��ԍ�_�� As String
    �ی��ԍ�_�E As String
    ������� As Long
    �S���x�� As Long
    �Z���ԋx�� As Double
    ����P�� As Double
End Type

Private m_workers() As kyugo_workers
Private worker_cnt As Long '�L�x�ƂƋ���P���̑ΏێҌv

Private m_���ԏ���J���������v As Long '�B
Private m_�S���x�Ɠ������v As Long '�C
Private m_�Z���ԋx�Ǝ��ԍ��v As Double '�D
Private m_����P���������v As Double '�E

Private m_�x�ƑΏێҍ��v As Long '�M�x�ƑΏێҌv
Private m_����P���Ώێҍ��v As Long '�N����P���ΏێҌv

Public targetCsvPath As String

Public m_��ی��Ώێ� As Boolean

Private Sub Class_Initialize()
    worker_cnt = 1
End Sub

Public Property Get �x��_����P���Ώێҍ��v() As Long '�L�x�ƂƋ���P���̑ΏێҌv
    �x��_����P���Ώێҍ��v = worker_cnt - 1
End Property

Public Property Get �x�Ƃ̂ݑΏێҍ��v() As Long '�M�x�ƑΏێҌv
    �x�Ƃ̂ݑΏێҍ��v = m_�x�ƑΏێҍ��v
End Property

Public Property Get �ی��ԍ�(ByVal index As Long) As String
    �ی��ԍ� = m_workers(index).�ی��ԍ�
End Property

Public Property Let �ی��ԍ�(ByVal index As Long, ByVal value As String)
    value = StrConv(value, vbNarrow)
    
    If m_��ی��Ώێ� Then
        If ���Ə�_�ی��Ҕԍ��`�F�b�N(value) Then
            m_workers(index).�ی��ԍ� = value
            BangoDivision.�ی�_���Ə��ԍ����� _
                m_workers(index).�ی��ԍ�, _
                m_workers(index).�ی��ԍ�_��, _
                m_workers(index).�ی��ԍ�_��, _
                m_workers(index).�ی��ԍ�_�E
        Else
            MsgBox m_workers(index).���� & "����̔�ی��Ҕԍ�������������܂���B" & vbNewLine & _
                "(��������: 1234-123456-1)" & _
                "(4��-6��-1��)" & _
                vbNewLine & "�C�����čēx���s���ĉ������B", vbCritical, "���̓G���["
            End
        End If
    
    Else
        m_workers(worker_cnt).�ی��ԍ� = vbNullString
    End If

End Property

Public Property Get �ی��ԍ�_��(ByVal index As Long) As String
    �ی��ԍ�_�� = m_workers(index).�ی��ԍ�_��
End Property

Public Property Get �ی��ԍ�_�E(ByVal index As Long) As String
    �ی��ԍ�_�E = m_workers(index).�ی��ԍ�_�E
End Property

Public Property Get �ی��ԍ�_��(ByVal index As Long) As String
    �ی��ԍ�_�� = m_workers(index).�ی��ԍ�_��
End Property

Public Property Get �S���x��(ByVal index As Long) As Long
    �S���x�� = m_workers(index).�S���x��
End Property

Public Property Get �S���x�Ɠ������v() As Long '�C
    �S���x�Ɠ������v = m_�S���x�Ɠ������v
End Property

Public Property Get �������(ByVal index As Long) As Long
    ������� = m_workers(index).�������
End Property

Public Property Get ����P��(ByVal index As Long) As Double
    ����P�� = m_workers(index).����P��
End Property

Public Property Get ����P���̂ݑΏێҍ��v() As Long '�N����P���ΏێҌv
    ����P���̂ݑΏێҍ��v = m_����P���Ώێҍ��v
End Property

Public Property Get ����P���������v() As Long '�E
    ����P���������v = WorksheetFunction.RoundUp(m_����P���������v, 0)
End Property

Public Property Get ���ԏ���J���������v() As Long '�B
    ���ԏ���J���������v = m_���ԏ���J���������v
End Property

Public Property Get ����(ByVal index As Long) As String
    ���� = m_workers(index).����
End Property

Public Property Let ����(ByVal index As Long, ByVal name As String)
    If name = vbNullString Then
        MsgBox "�������󔒂̐l�����܂��B", vbCritical
        End
    Else
        m_workers(worker_cnt).���� = name
    End If
End Property

Public Property Get �Z���ԋx��(ByVal index As Long) As Double
    �Z���ԋx�� = m_workers(index).�Z���ԋx��
End Property

Public Property Get �Z���ԋx�Ǝ��ԍ��v() As Long '�D
    �Z���ԋx�Ǝ��ԍ��v = WorksheetFunction.RoundUp(m_�Z���ԋx�Ǝ��ԍ��v, 0)
End Property

Public Property Let �Z���ԋx��(ByVal index As Long, value As Double)
    If value < 1 Then '1���Ԗ�����0����
        m_workers(index).�Z���ԋx�� = 0
    Else
        m_workers(index).�Z���ԋx�� = value
    End If
End Property

Public Sub GetWorker(ByVal ��ی��� As Boolean)
    If targetCsvPath <> vbNullString Then
        m_��ی��Ώێ� = ��ی���
        
        Dim fileHandle As Integer: fileHandle = FreeFile
        Dim textLine As String
        
        Open targetCsvPath For Input As #fileHandle
        
        On Error GoTo TypeErr
        Do While Not EOF(fileHandle)
            Line Input #fileHandle, textLine
            Dim ary As Variant: ary = Split(textLine, ",")
            
            If Is�ΏۘJ����(ary(1)) Then 'ary(1)=�ی��ԍ�
                ReDim Preserve m_workers(worker_cnt)
                
                Me.����(worker_cnt) = ary(0)
                Me.�ی��ԍ�(worker_cnt) = ary(1)
                m_workers(worker_cnt).������� = WorksheetFunction.RoundDown(ary(2), 0)
                m_workers(worker_cnt).�S���x�� = WorksheetFunction.RoundDown(ary(3), 0)
                Me.�Z���ԋx��(worker_cnt) = WorksheetFunction.Floor(ary(4), 0.5)
                m_workers(worker_cnt).����P�� = WorksheetFunction.Floor(ary(5), 0.5)
                
                '�B
                m_���ԏ���J���������v = m_���ԏ���J���������v + m_workers(worker_cnt).�������
                '�C
                m_�S���x�Ɠ������v = m_�S���x�Ɠ������v + m_workers(worker_cnt).�S���x��
                '�D
                m_�Z���ԋx�Ǝ��ԍ��v = m_�Z���ԋx�Ǝ��ԍ��v + m_workers(worker_cnt).�Z���ԋx��
                '�E
                m_����P���������v = m_����P���������v + m_workers(worker_cnt).����P��
                
                '�M�x�ƑΏێҌv�̃J�E���g
                If m_workers(worker_cnt).�S���x�� > 0 Or m_workers(worker_cnt).�Z���ԋx�� > 0 Then m_�x�ƑΏێҍ��v = m_�x�ƑΏێҍ��v + 1
                '�N����P���ΏێҌv�̃J�E���g
                If m_workers(worker_cnt).����P�� > 0 Then m_����P���Ώێҍ��v = m_����P���Ώێҍ��v + 1
                
                worker_cnt = worker_cnt + 1
            End If
            
        Loop
        Close #fileHandle
    End If
    
    Exit Sub
    
TypeErr:
    If Err.Number <> 0 Then
        MsgBox m_workers(worker_cnt).���� & "����̓��͂���Ă���f�[�^�^���Ⴂ�܂��B" & vbNewLine & _
            vbNewLine & "�Ⴆ�ΐ��l�ł���ׂ��ӏ��ɕ��������͂���Ă��܂��B" & _
            vbNewLine & "�f�[�^���C�����čēx���s���ĉ������B", _
            vbCritical, "���̓G���["
        Close #fileHandle
        End
    End If
End Sub

Private Function Is�ΏۘJ����(ByVal �ی��ԍ� As String) As Boolean
    If m_��ی��Ώێ� Then
        '��ی��җp
        If IsNumeric(Left(�ی��ԍ�, 4)) Then
            Is�ΏۘJ���� = True
        End If
    Else
        '��ی��҈ȊO�p
        If �ی��ԍ� = vbNullString Then
            Is�ΏۘJ���� = True
        End If
    End If
End Function

Public Sub �C�~�f�B�G�C�g�Ŋm�F()
    Dim i As Long
    For i = 1 To Me.�x��_����P���Ώێҍ��v
        Debug.Print Me.����(i), Me.�ی��ԍ�(i), Me.�������(i), Me.�S���x��(i), Me.�Z���ԋx��(i), Me.����P��(i)
    Next
    
    '�L�M�N�̊m�F
    Debug.Print "�L"; Me.�x��_����P���Ώێҍ��v, "�M"; Me.�x�Ƃ̂ݑΏێҍ��v, "�N"; Me.����P���̂ݑΏێҍ��v
    
    '�F�G�H�I�̊m�F
    Debug.Print "�F"; Me.���ԏ���J���������v, "�G"; Me.�S���x�Ɠ������v, "�H"; Me.�Z���ԋx�Ǝ��ԍ��v, "�I"; Me.����P���������v

End Sub
