Attribute VB_Name = "BangoDivision"
'@folder("utility")
Option Explicit

Private Const �ԍ��`�� As String = "^\d{4}-\d{6}-\d{1}$"

Public Function ���Ə�_�ی��Ҕԍ��`�F�b�N(ByVal bango As String) As Boolean
    '���Ə��ԍ��Ɣ�ی��Ҕԍ��̌`���͓���

    Dim re As RegExp: Set re = New RegExp
    re.Pattern = �ԍ��`��
    re.Global = False
    If re.Test(bango) Then
        ���Ə�_�ی��Ҕԍ��`�F�b�N = True
    End If
End Function


'�l���ɏo�͂��邽�߂ɕی��ԍ��E���Ə��ԍ����n�C�t���ŕ�������(�ی��ԍ��E���Ə��ԍ��͓���`��)
Public Sub �ی�_���Ə��ԍ�����(ByVal numAll As String, _
                               ByRef numLeft As String, _
                               ByRef numMid As String, _
                               ByRef numRight As String)
                               
    Dim firstHyphen As Long, secondHyphen As Long
    firstHyphen = InStr(numAll, "-")
    secondHyphen = InStr(firstHyphen + 1, numAll, "-")
    numLeft = Left(numAll, firstHyphen - 1)
    numMid = Mid(numAll, firstHyphen + 1, secondHyphen - firstHyphen - 1)
    numRight = Mid(numAll, secondHyphen + 1)
End Sub

