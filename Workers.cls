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
    氏名 As String
    保険番号 As String
    保険番号_左 As String
    保険番号_中 As String
    保険番号_右 As String
    所定日数 As Long
    全日休業 As Long
    短時間休業 As Double
    教育訓練 As Double
End Type

Private m_workers() As kyugo_workers
Private worker_cnt As Long '�L休業と教育訓練の対象者計

Private m_月間所定労働日数合計 As Long '�B
Private m_全日休業日数合計 As Long '�C
Private m_短時間休業時間合計 As Double '�D
Private m_教育訓練日数合計 As Double '�E

Private m_休業対象者合計 As Long '�M休業対象者計
Private m_教育訓練対象者合計 As Long '�N教育訓練対象者計

Public targetCsvPath As String

Public m_被保険対象者 As Boolean

Private Sub Class_Initialize()
    worker_cnt = 1
End Sub

Public Property Get 休業_教育訓練対象者合計() As Long '�L休業と教育訓練の対象者計
    休業_教育訓練対象者合計 = worker_cnt - 1
End Property

Public Property Get 休業のみ対象者合計() As Long '�M休業対象者計
    休業のみ対象者合計 = m_休業対象者合計
End Property

Public Property Get 保険番号(ByVal index As Long) As String
    保険番号 = m_workers(index).保険番号
End Property

Public Property Let 保険番号(ByVal index As Long, ByVal value As String)
    value = StrConv(value, vbNarrow)
    
    If m_被保険対象者 Then
        If 事業所_保険者番号チェック(value) Then
            m_workers(index).保険番号 = value
            BangoDivision.保険_事業所番号分割 _
                m_workers(index).保険番号, _
                m_workers(index).保険番号_左, _
                m_workers(index).保険番号_中, _
                m_workers(index).保険番号_右
        Else
            MsgBox m_workers(index).氏名 & "さんの被保険者番号が正しくありません。" & vbNewLine & _
                "(正しい例: 1234-123456-1)" & _
                "(4桁-6桁-1桁)" & _
                vbNewLine & "修正して再度実行して下さい。", vbCritical, "入力エラー"
            End
        End If
    
    Else
        m_workers(worker_cnt).保険番号 = vbNullString
    End If

End Property

Public Property Get 保険番号_中(ByVal index As Long) As String
    保険番号_中 = m_workers(index).保険番号_中
End Property

Public Property Get 保険番号_右(ByVal index As Long) As String
    保険番号_右 = m_workers(index).保険番号_右
End Property

Public Property Get 保険番号_左(ByVal index As Long) As String
    保険番号_左 = m_workers(index).保険番号_左
End Property

Public Property Get 全日休業(ByVal index As Long) As Long
    全日休業 = m_workers(index).全日休業
End Property

Public Property Get 全日休業日数合計() As Long '�C
    全日休業日数合計 = m_全日休業日数合計
End Property

Public Property Get 所定日数(ByVal index As Long) As Long
    所定日数 = m_workers(index).所定日数
End Property

Public Property Get 教育訓練(ByVal index As Long) As Double
    教育訓練 = m_workers(index).教育訓練
End Property

Public Property Get 教育訓練のみ対象者合計() As Long '�N教育訓練対象者計
    教育訓練のみ対象者合計 = m_教育訓練対象者合計
End Property

Public Property Get 教育訓練日数合計() As Long '�E
    教育訓練日数合計 = WorksheetFunction.RoundUp(m_教育訓練日数合計, 0)
End Property

Public Property Get 月間所定労働日数合計() As Long '�B
    月間所定労働日数合計 = m_月間所定労働日数合計
End Property

Public Property Get 氏名(ByVal index As Long) As String
    氏名 = m_workers(index).氏名
End Property

Public Property Let 氏名(ByVal index As Long, ByVal name As String)
    If name = vbNullString Then
        MsgBox "氏名が空白の人がいます。", vbCritical
        End
    Else
        m_workers(worker_cnt).氏名 = name
    End If
End Property

Public Property Get 短時間休業(ByVal index As Long) As Double
    短時間休業 = m_workers(index).短時間休業
End Property

Public Property Get 短時間休業時間合計() As Long '�D
    短時間休業時間合計 = WorksheetFunction.RoundUp(m_短時間休業時間合計, 0)
End Property

Public Property Let 短時間休業(ByVal index As Long, value As Double)
    If value < 1 Then '1時間未満は0扱い
        m_workers(index).短時間休業 = 0
    Else
        m_workers(index).短時間休業 = value
    End If
End Property

Public Sub GetWorker(ByVal 被保険者 As Boolean)
    If targetCsvPath <> vbNullString Then
        m_被保険対象者 = 被保険者
        
        Dim fileHandle As Integer: fileHandle = FreeFile
        Dim textLine As String
        
        Open targetCsvPath For Input As #fileHandle
        
        On Error GoTo TypeErr
        Do While Not EOF(fileHandle)
            Line Input #fileHandle, textLine
            Dim ary As Variant: ary = Split(textLine, ",")
            
            If Is対象労働者(ary(1)) Then 'ary(1)=保険番号
                ReDim Preserve m_workers(worker_cnt)
                
                Me.氏名(worker_cnt) = ary(0)
                Me.保険番号(worker_cnt) = ary(1)
                m_workers(worker_cnt).所定日数 = WorksheetFunction.RoundDown(ary(2), 0)
                m_workers(worker_cnt).全日休業 = WorksheetFunction.RoundDown(ary(3), 0)
                Me.短時間休業(worker_cnt) = WorksheetFunction.Floor(ary(4), 0.5)
                m_workers(worker_cnt).教育訓練 = WorksheetFunction.Floor(ary(5), 0.5)
                
                '�B
                m_月間所定労働日数合計 = m_月間所定労働日数合計 + m_workers(worker_cnt).所定日数
                '�C
                m_全日休業日数合計 = m_全日休業日数合計 + m_workers(worker_cnt).全日休業
                '�D
                m_短時間休業時間合計 = m_短時間休業時間合計 + m_workers(worker_cnt).短時間休業
                '�E
                m_教育訓練日数合計 = m_教育訓練日数合計 + m_workers(worker_cnt).教育訓練
                
                '�M休業対象者計のカウント
                If m_workers(worker_cnt).全日休業 > 0 Or m_workers(worker_cnt).短時間休業 > 0 Then m_休業対象者合計 = m_休業対象者合計 + 1
                '�N教育訓練対象者計のカウント
                If m_workers(worker_cnt).教育訓練 > 0 Then m_教育訓練対象者合計 = m_教育訓練対象者合計 + 1
                
                worker_cnt = worker_cnt + 1
            End If
            
        Loop
        Close #fileHandle
    End If
    
    Exit Sub
    
TypeErr:
    If Err.Number <> 0 Then
        MsgBox m_workers(worker_cnt).氏名 & "さんの入力されているデータ型が違います。" & vbNewLine & _
            vbNewLine & "例えば数値であるべき箇所に文字が入力されています。" & _
            vbNewLine & "データを修正して再度実行して下さい。", _
            vbCritical, "入力エラー"
        Close #fileHandle
        End
    End If
End Sub

Private Function Is対象労働者(ByVal 保険番号 As String) As Boolean
    If m_被保険対象者 Then
        '被保険者用
        If IsNumeric(Left(保険番号, 4)) Then
            Is対象労働者 = True
        End If
    Else
        '被保険者以外用
        If 保険番号 = vbNullString Then
            Is対象労働者 = True
        End If
    End If
End Function

Public Sub イミディエイトで確認()
    Dim i As Long
    For i = 1 To Me.休業_教育訓練対象者合計
        Debug.Print Me.氏名(i), Me.保険番号(i), Me.所定日数(i), Me.全日休業(i), Me.短時間休業(i), Me.教育訓練(i)
    Next
    
    '�L�M�Nの確認
    Debug.Print "�L"; Me.休業_教育訓練対象者合計, "�M"; Me.休業のみ対象者合計, "�N"; Me.教育訓練のみ対象者合計
    
    '�F�G�H�Iの確認
    Debug.Print "�F"; Me.月間所定労働日数合計, "�G"; Me.全日休業日数合計, "�H"; Me.短時間休業時間合計, "�I"; Me.教育訓練日数合計

End Sub

