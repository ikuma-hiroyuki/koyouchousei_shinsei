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
    –¼ As String
    •ÛŒ¯”Ô† As String
    •ÛŒ¯”Ô†_¶ As String
    •ÛŒ¯”Ô†_’† As String
    •ÛŒ¯”Ô†_‰E As String
    Š’è“ú” As Long
    ‘S“ú‹x‹Æ As Long
    ’ZŠÔ‹x‹Æ As Double
    ‹³ˆçŒP—û As Double
End Type

Private m_workers() As kyugo_workers
Private worker_cnt As Long '‡L‹x‹Æ‚Æ‹³ˆçŒP—û‚Ì‘ÎÛÒŒv

Private m_ŒŠÔŠ’è˜J“­“ú”‡Œv As Long '‡B
Private m_‘S“ú‹x‹Æ“ú”‡Œv As Long '‡C
Private m_’ZŠÔ‹x‹ÆŠÔ‡Œv As Double '‡D
Private m_‹³ˆçŒP—û“ú”‡Œv As Double '‡E

Private m_‹x‹Æ‘ÎÛÒ‡Œv As Long '‡M‹x‹Æ‘ÎÛÒŒv
Private m_‹³ˆçŒP—û‘ÎÛÒ‡Œv As Long '‡N‹³ˆçŒP—û‘ÎÛÒŒv

Public targetCsvPath As String

Public m_”í•ÛŒ¯‘ÎÛÒ As Boolean

Private Sub Class_Initialize()
    worker_cnt = 1
End Sub

Public Property Get ‹x‹Æ_‹³ˆçŒP—û‘ÎÛÒ‡Œv() As Long '‡L‹x‹Æ‚Æ‹³ˆçŒP—û‚Ì‘ÎÛÒŒv
    ‹x‹Æ_‹³ˆçŒP—û‘ÎÛÒ‡Œv = worker_cnt - 1
End Property

Public Property Get ‹x‹Æ‚Ì‚İ‘ÎÛÒ‡Œv() As Long '‡M‹x‹Æ‘ÎÛÒŒv
    ‹x‹Æ‚Ì‚İ‘ÎÛÒ‡Œv = m_‹x‹Æ‘ÎÛÒ‡Œv
End Property

Public Property Get •ÛŒ¯”Ô†(ByVal index As Long) As String
    •ÛŒ¯”Ô† = m_workers(index).•ÛŒ¯”Ô†
End Property

Public Property Let •ÛŒ¯”Ô†(ByVal index As Long, ByVal value As String)
    value = StrConv(value, vbNarrow)
    
    If m_”í•ÛŒ¯‘ÎÛÒ Then
        If –‹ÆŠ_•ÛŒ¯Ò”Ô†ƒ`ƒFƒbƒN(value) Then
            m_workers(index).•ÛŒ¯”Ô† = value
            BangoDivision.•ÛŒ¯_–‹ÆŠ”Ô†•ªŠ„ _
                m_workers(index).•ÛŒ¯”Ô†, _
                m_workers(index).•ÛŒ¯”Ô†_¶, _
                m_workers(index).•ÛŒ¯”Ô†_’†, _
                m_workers(index).•ÛŒ¯”Ô†_‰E
        Else
            MsgBox m_workers(index).–¼ & "‚³‚ñ‚Ì”í•ÛŒ¯Ò”Ô†‚ª³‚µ‚­‚ ‚è‚Ü‚¹‚ñB" & vbNewLine & _
                "(³‚µ‚¢—á: 1234-123456-1)" & _
                "(4Œ…-6Œ…-1Œ…)" & _
                vbNewLine & "C³‚µ‚ÄÄ“xÀs‚µ‚Ä‰º‚³‚¢B", vbCritical, "“ü—ÍƒGƒ‰["
            End
        End If
    
    Else
        m_workers(worker_cnt).•ÛŒ¯”Ô† = vbNullString
    End If

End Property

Public Property Get •ÛŒ¯”Ô†_’†(ByVal index As Long) As String
    •ÛŒ¯”Ô†_’† = m_workers(index).•ÛŒ¯”Ô†_’†
End Property

Public Property Get •ÛŒ¯”Ô†_‰E(ByVal index As Long) As String
    •ÛŒ¯”Ô†_‰E = m_workers(index).•ÛŒ¯”Ô†_‰E
End Property

Public Property Get •ÛŒ¯”Ô†_¶(ByVal index As Long) As String
    •ÛŒ¯”Ô†_¶ = m_workers(index).•ÛŒ¯”Ô†_¶
End Property

Public Property Get ‘S“ú‹x‹Æ(ByVal index As Long) As Long
    ‘S“ú‹x‹Æ = m_workers(index).‘S“ú‹x‹Æ
End Property

Public Property Get ‘S“ú‹x‹Æ“ú”‡Œv() As Long '‡C
    ‘S“ú‹x‹Æ“ú”‡Œv = m_‘S“ú‹x‹Æ“ú”‡Œv
End Property

Public Property Get Š’è“ú”(ByVal index As Long) As Long
    Š’è“ú” = m_workers(index).Š’è“ú”
End Property

Public Property Get ‹³ˆçŒP—û(ByVal index As Long) As Double
    ‹³ˆçŒP—û = m_workers(index).‹³ˆçŒP—û
End Property

Public Property Get ‹³ˆçŒP—û‚Ì‚İ‘ÎÛÒ‡Œv() As Long '‡N‹³ˆçŒP—û‘ÎÛÒŒv
    ‹³ˆçŒP—û‚Ì‚İ‘ÎÛÒ‡Œv = m_‹³ˆçŒP—û‘ÎÛÒ‡Œv
End Property

Public Property Get ‹³ˆçŒP—û“ú”‡Œv() As Long '‡E
    ‹³ˆçŒP—û“ú”‡Œv = WorksheetFunction.RoundUp(m_‹³ˆçŒP—û“ú”‡Œv, 0)
End Property

Public Property Get ŒŠÔŠ’è˜J“­“ú”‡Œv() As Long '‡B
    ŒŠÔŠ’è˜J“­“ú”‡Œv = m_ŒŠÔŠ’è˜J“­“ú”‡Œv
End Property

Public Property Get –¼(ByVal index As Long) As String
    –¼ = m_workers(index).–¼
End Property

Public Property Let –¼(ByVal index As Long, ByVal name As String)
    If name = vbNullString Then
        MsgBox "–¼‚ª‹ó”’‚Ìl‚ª‚¢‚Ü‚·B", vbCritical
        End
    Else
        m_workers(worker_cnt).–¼ = name
    End If
End Property

Public Property Get ’ZŠÔ‹x‹Æ(ByVal index As Long) As Double
    ’ZŠÔ‹x‹Æ = m_workers(index).’ZŠÔ‹x‹Æ
End Property

Public Property Get ’ZŠÔ‹x‹ÆŠÔ‡Œv() As Long '‡D
    ’ZŠÔ‹x‹ÆŠÔ‡Œv = WorksheetFunction.RoundUp(m_’ZŠÔ‹x‹ÆŠÔ‡Œv, 0)
End Property

Public Property Let ’ZŠÔ‹x‹Æ(ByVal index As Long, value As Double)
    If value < 1 Then '1ŠÔ–¢–‚Í0ˆµ‚¢
        m_workers(index).’ZŠÔ‹x‹Æ = 0
    Else
        m_workers(index).’ZŠÔ‹x‹Æ = value
    End If
End Property

Public Sub GetWorker(ByVal ”í•ÛŒ¯Ò As Boolean)
    If targetCsvPath <> vbNullString Then
        m_”í•ÛŒ¯‘ÎÛÒ = ”í•ÛŒ¯Ò
        
        Dim fileHandle As Integer: fileHandle = FreeFile
        Dim textLine As String
        
        Open targetCsvPath For Input As #fileHandle
        
        On Error GoTo TypeErr
        Do While Not EOF(fileHandle)
            Line Input #fileHandle, textLine
            Dim ary As Variant: ary = Split(textLine, ",")
            
            If Is‘ÎÛ˜J“­Ò(ary(1)) Then 'ary(1)=•ÛŒ¯”Ô†
                ReDim Preserve m_workers(worker_cnt)
                
                Me.–¼(worker_cnt) = ary(0)
                Me.•ÛŒ¯”Ô†(worker_cnt) = ary(1)
                m_workers(worker_cnt).Š’è“ú” = WorksheetFunction.RoundDown(ary(2), 0)
                m_workers(worker_cnt).‘S“ú‹x‹Æ = WorksheetFunction.RoundDown(ary(3), 0)
                Me.’ZŠÔ‹x‹Æ(worker_cnt) = WorksheetFunction.Floor(ary(4), 0.5)
                m_workers(worker_cnt).‹³ˆçŒP—û = WorksheetFunction.Floor(ary(5), 0.5)
                
                '‡B
                m_ŒŠÔŠ’è˜J“­“ú”‡Œv = m_ŒŠÔŠ’è˜J“­“ú”‡Œv + m_workers(worker_cnt).Š’è“ú”
                '‡C
                m_‘S“ú‹x‹Æ“ú”‡Œv = m_‘S“ú‹x‹Æ“ú”‡Œv + m_workers(worker_cnt).‘S“ú‹x‹Æ
                '‡D
                m_’ZŠÔ‹x‹ÆŠÔ‡Œv = m_’ZŠÔ‹x‹ÆŠÔ‡Œv + m_workers(worker_cnt).’ZŠÔ‹x‹Æ
                '‡E
                m_‹³ˆçŒP—û“ú”‡Œv = m_‹³ˆçŒP—û“ú”‡Œv + m_workers(worker_cnt).‹³ˆçŒP—û
                
                '‡M‹x‹Æ‘ÎÛÒŒv‚ÌƒJƒEƒ“ƒg
                If m_workers(worker_cnt).‘S“ú‹x‹Æ > 0 Or m_workers(worker_cnt).’ZŠÔ‹x‹Æ > 0 Then m_‹x‹Æ‘ÎÛÒ‡Œv = m_‹x‹Æ‘ÎÛÒ‡Œv + 1
                '‡N‹³ˆçŒP—û‘ÎÛÒŒv‚ÌƒJƒEƒ“ƒg
                If m_workers(worker_cnt).‹³ˆçŒP—û > 0 Then m_‹³ˆçŒP—û‘ÎÛÒ‡Œv = m_‹³ˆçŒP—û‘ÎÛÒ‡Œv + 1
                
                worker_cnt = worker_cnt + 1
            End If
            
        Loop
        Close #fileHandle
    End If
    
    Exit Sub
    
TypeErr:
    If Err.Number <> 0 Then
        MsgBox m_workers(worker_cnt).–¼ & "‚³‚ñ‚Ì“ü—Í‚³‚ê‚Ä‚¢‚éƒf[ƒ^Œ^‚ªˆá‚¢‚Ü‚·B" & vbNewLine & _
            vbNewLine & "—á‚¦‚Î”’l‚Å‚ ‚é‚×‚«‰ÓŠ‚É•¶š‚ª“ü—Í‚³‚ê‚Ä‚¢‚Ü‚·B" & _
            vbNewLine & "ƒf[ƒ^‚ğC³‚µ‚ÄÄ“xÀs‚µ‚Ä‰º‚³‚¢B", _
            vbCritical, "“ü—ÍƒGƒ‰["
        Close #fileHandle
        End
    End If
End Sub

Private Function Is‘ÎÛ˜J“­Ò(ByVal •ÛŒ¯”Ô† As String) As Boolean
    If m_”í•ÛŒ¯‘ÎÛÒ Then
        '”í•ÛŒ¯Ò—p
        If IsNumeric(Left(•ÛŒ¯”Ô†, 4)) Then
            Is‘ÎÛ˜J“­Ò = True
        End If
    Else
        '”í•ÛŒ¯ÒˆÈŠO—p
        If •ÛŒ¯”Ô† = vbNullString Then
            Is‘ÎÛ˜J“­Ò = True
        End If
    End If
End Function

Public Sub ƒCƒ~ƒfƒBƒGƒCƒg‚ÅŠm”F()
    Dim i As Long
    For i = 1 To Me.‹x‹Æ_‹³ˆçŒP—û‘ÎÛÒ‡Œv
        Debug.Print Me.–¼(i), Me.•ÛŒ¯”Ô†(i), Me.Š’è“ú”(i), Me.‘S“ú‹x‹Æ(i), Me.’ZŠÔ‹x‹Æ(i), Me.‹³ˆçŒP—û(i)
    Next
    
    '‡L‡M‡N‚ÌŠm”F
    Debug.Print "‡L"; Me.‹x‹Æ_‹³ˆçŒP—û‘ÎÛÒ‡Œv, "‡M"; Me.‹x‹Æ‚Ì‚İ‘ÎÛÒ‡Œv, "‡N"; Me.‹³ˆçŒP—û‚Ì‚İ‘ÎÛÒ‡Œv
    
    '‡F‡G‡H‡I‚ÌŠm”F
    Debug.Print "‡F"; Me.ŒŠÔŠ’è˜J“­“ú”‡Œv, "‡G"; Me.‘S“ú‹x‹Æ“ú”‡Œv, "‡H"; Me.’ZŠÔ‹x‹ÆŠÔ‡Œv, "‡I"; Me.‹³ˆçŒP—û“ú”‡Œv

End Sub

