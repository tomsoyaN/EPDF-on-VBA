VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private token As token '''Token
Private fn As Integer ''' FileNUmber
Private cur As Byte ''Current Byte
Private lookahead As Byte ''Next Byte
Private index As Long ''index of reader
Private symtable As Object
Private Escape As Boolean
Public Sub ReadFile(sym As t_symtalbe)
    Set symtable = sym.symtable
    Dim buf As String
    fn = FreeFile
    Open "C:\Users\nonaka\Desktop\sample.pdf" For Binary As #fn
    Get #fn, 1, cur
    index = Seek(fn)
    Prefetch
End Sub
Sub lexan(ByRef tok As t_token) ''字句解析
    Dim buf As String, sym As Integer
    Escape = False
    Do While (cur <> EOF(fn) And IsNewLine = False And IsSpace = False)
        If (Not Escape) Then ''エスケープ文字処理
            sym = IsSymbol
        Else
            sym = 0
            Escape = False
        End If
        If (sym <> 0) Then
            If (buf <> "") Then Exit Do
            buf = buf + Chr(cur)
            If (sym = 1) Then
                ''Debug.Print buf
                ReadByte
                Skip
                tok.Clone (symtable(buf))
                Exit Sub
            ElseIf (sym = 2) Then
                ReadByte
                buf = buf + Chr(cur)
                ''Debug.Print buf
                ReadByte
                Skip
                tok.Clone (symtable(buf))
                Exit Sub
            End If
        End If
        buf = buf + Chr(cur)
        ReadByte
    Loop
    If symtable.Exists(buf) Then
        tok.Clone (symtable(buf))
    ElseIf buf = "%%EOF" Then
        tok.literal = "EOF"
    ElseIf IsNumeric(buf) Then
        tok.num = CLng(buf)
        tok.literal = "VALUE"
    Else
        tok.literal = buf
    End If
    ''Debug.Print buf
End Sub
Sub Skip()
    sl = True
    ss = True
    While (Not (sl = False And ss = False))
        sl = IsNewLine
        ss = IsSpace
    Wend
End Sub
Sub ReadByte() ''読み出し
    cur = lookahead
    index = index + 1
    Seek fn, index
    Prefetch
End Sub
Sub Prefetch() ''先読み
    Get #fn, , lookahead
    Seek fn, index
End Sub

Function IsNewLine()  ''改行文字判定
    If (Chr(cur) = vbCr Or Chr(cur) = vbLf) Then
        If (Chr(cur) = vbCr And Chr(lookahead) = vbLf) Then
            ReadByte
        End If
        ReadByte
        IsNewLine = True
    Else
        IsNewLine = False
    End If
End Function

Function IsSpace() ''スペース文字判定
    If (Chr(cur) = " ") Then
        While (Chr(cur) = " ")
            ReadByte
        Wend
        IsSpace = True
    Else
        IsSpace = False
    End If
End Function
Function IsSymbol()
    n = 0
    If Chr(cur) = "<" Or Chr(cur) = "(" Or Chr(cur) = "[" Then
        n = 1
        If (Chr(lookahead) = "<" Or Chr(lookahead) = "?") Then
            n = 2
        End If
    ElseIf Chr(cur) = ">" Or Chr(cur) = ")" Or Chr(cur) = "]" Then
        n = 1
        If (Chr(lookahead) = ">") Then
            n = 2
        End If
    ElseIf Chr(cur) = "?" And Chr(lookahead) = ">" Then
        n = 2
    ElseIf Chr(cur) = "/" Then
        n = 1
    End If
    If Chr(cur) = "\" Then
        Escape = True
    Else
        Escape = False
    End If
    IsSymbol = n
End Function
Sub Dispose()
    Close #n
End Sub
