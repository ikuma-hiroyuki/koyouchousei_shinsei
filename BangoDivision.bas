Attribute VB_Name = "BangoDivision"
'@folder("utility")
Option Explicit

Private Const 番号形式 As String = "^\d{4}-\d{6}-\d{1}$"

Public Function 事業所_保険者番号チェック(ByVal bango As String) As Boolean
    '事業所番号と被保険者番号の形式は同一

    Dim re As RegExp: Set re = New RegExp
    re.Pattern = 番号形式
    re.Global = False
    If re.Test(bango) Then
        事業所_保険者番号チェック = True
    End If
End Function


'様式に出力するために保険番号・事業所番号をハイフンで分割する(保険番号・事業所番号は同一形式)
Public Sub 保険_事業所番号分割(ByVal numAll As String, _
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

