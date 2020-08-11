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

Private m_判定期間_開始日 As Date
Private m_判定期間_終了日 As Date
Private m_所定労働時間 As Double
Private m_作成日 As Date
Private m_会社名 As String
Private m_事業所番号 As String
Private m_事業所番号_左 As String
Private m_事業所番号_中 As String
Private m_事業所番号_右 As String
Private m_事業者氏名 As String
Private m_労働者代表氏名 As String

Private m_ws As Worksheet

Private Sub Class_Initialize()
    Set m_ws = ThisWorkbook.Worksheets("全体設定")
End Sub

Public Sub GetCompanyInfo()
    m_判定期間_開始日 = InputCheck(m_ws.Range("B2"))
    m_判定期間_終了日 = InputCheck(m_ws.Range("B3"))
    m_所定労働時間 = WorksheetFunction.RoundDown(InputCheck(m_ws.Range("B5")), 0)
    m_作成日 = InputCheck(m_ws.Range("B8"))
    m_会社名 = InputCheck(m_ws.Range("B10"))
    Me.事業所番号 = m_ws.Range("B11")
    m_事業者氏名 = InputCheck(m_ws.Range("B12"))
    m_労働者代表氏名 = InputCheck(m_ws.Range("B13"))
End Sub

Public Property Get 判定期間_開始日() As Date
    判定期間_開始日 = m_判定期間_開始日
End Property

Public Property Get 判定期間_終了日() As Date
    判定期間_終了日 = m_判定期間_終了日
End Property

Public Property Get 所定労働時間() As Double
    所定労働時間 = m_所定労働時間
End Property

Public Property Get 作成日() As Date
    作成日 = m_作成日
End Property

Public Property Get 会社名() As String
    会社名 = m_会社名
End Property

Public Property Get 事業所番号() As String
    事業所番号 = m_事業所番号
End Property

Public Property Get 事業所番号_左() As String
    事業所番号_左 = m_事業所番号_左
End Property

Public Property Get 事業所番号_中() As String
    事業所番号_中 = m_事業所番号_中
End Property

Public Property Get 事業所番号_右() As String
    事業所番号_右 = m_事業所番号_右
End Property

Public Property Let 事業所番号(ByVal value As String)
    value = StrConv(value, vbNarrow)
    If 事業所_保険者番号チェック(value) Then
        m_事業所番号 = value
        保険_事業所番号分割 m_事業所番号, m_事業所番号_左, m_事業所番号_中, m_事業所番号_右
    Else
        MsgBox "時強者番号が正しくありません。(正しい例:1234-123456-1)" & vbNewLine & _
            "修正して再度実行して下さい", vbCritical, "入力ミス？"
        End
    End If
End Property

Public Property Get 事業者氏名() As String
    事業者氏名 = m_事業者氏名
End Property

Public Property Get 労働者代表氏名() As String
    労働者代表氏名 = m_労働者代表氏名
End Property

Private Function InputCheck(ByRef val As Variant) As Variant
    '未入力だったらエラーを表示する。入力済だったらそのまま値を返す。
    If val = 0 Or val = vbNullString Then
        MsgBox "全体設定に未入力項目があります。" & vbNewLine & "入力して再度実行して下さい。", vbCritical
        End
    Else
        InputCheck = val
    End If
End Function
