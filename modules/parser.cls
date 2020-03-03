VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lexer As New lexer
Private token As New t_token
Public symtable As t_symtalbe
Private m_pdf As pdf
Private parse_type As String
Public Sub Initialize(sym As t_symtalbe)
    Set symtable = sym
    Set m_pdf = New pdf
    m_pdf.Initialize
    lexer.ReadFile symtable
End Sub
Sub Parse()
    Dim a As Variant
    token.literal = ""
    match token.literal
    While (token.literal <> "EOF")
        ParseSyntax
        Debug.Print token.literal
    Wend
    lexer.Dispose
End Sub
Sub ParseSyntax()
    If token.literal = "VALUE" Then
        Dim val As Integer, gen As Integer
        val = token.num
        match "VALUE"
        If token.literal = "VALUE" Then
            gen = token.num
            match "VALUE"
            If token.literal = "OBJ_S" Then
                Dim obj As NamedObject
                Set obj = New NamedObject
                obj.Initialize val, gen
                match "OBJ_S"
                While (Not token.literal = "OBJ_E")
                    obj.objs.Add ParseObject
                Wend
                match "OBJ_E"
                m_pdf.body.Add obj
            End If
        End If
    Else
        match token.literal
    End If
    
End Sub
Function ParseObject() As Variant
    If token.literal = "DIC_S" Then
        parse_type = "DICTIONARY"
        match token.literal
        Dim dic As Collection
        Set dic = New Collection
        Dim key As Variant, val As Variant
        While (Not token.literal = "DIC_E")
            Set key = ParseObject()
            Set val = ParseObject()
            dic.Add val, key.str
        Wend
        Set ParseObject = dic
        match "DIC_E"
    ElseIf token.literal = "ARRAY_S" Then
        parse_type = "ARRAY"
        match token.literal
        Dim ary As Collection
        Set ary = New Collection
        Dim elem As Variant
        While (Not token.literal = "ARRAY_E")
            Set elem = ParseObject()
            ary.Add elem
        Wend
        Set ParseObject = ary
        match "ARRAY_E"
    ElseIf token.literal = "BIN_S" Then
        Dim basebin As BaseObject
        Set basebin = New BaseObject
        match token.literal
        basebin = token.literal
        Set ParseObject = basebin
        match token.literal
        match "BIN_E"
    ElseIf token.literal = "SLASH" Then
        Dim varstr As BaseObject
        Set varstr = New BaseObject
        match token.literal
        varstr.str = token.literal
        Set ParseObject = varstr
        match token.literal
    ElseIf token.literal = "STRING_S" Then
        match token.literal
        Dim basestr As BaseObject
        Set basestr = New BaseObject
        While (Not token.literal = "STRING_E")
            basestr.str = basestr.str + token.literal
            match token.literal
        Wend
        Set ParseObject = basestr
        match "STRING_E"
    ElseIf token.literal = "STREAM_S" Then
        match token.literal
        Dim basestm As BaseObject
        Set basestm = New BaseObject
        While (Not token.literal = "STREAM_E")
            basestm.str = basestm.str + token.literal
            match token.literal
        Wend
        match "STREAM_E"
        Set ParseObject = basestm
    ElseIf token.literal = "VALUE" Then
        Dim vals As BaseObject
        Set vals = New BaseObject
        If parse_type = "DICTIONARY" Then
            While (Not token.literal = "SLASH")
                If token.literal = "VALUE" Then
                    vals.str = vals.str + CStr(token.num)
                Else
                    vals.str = vals.str + token.literal
                End If
                match token.literal
            Wend
        Else
            vals.str = CStr(token.num)
            match token.literal
        End If
        Set ParseObject = vals
    Else
        Dim baseconst As BaseObject
        Set baseconst = New BaseObject
        baseconst.str = token.literal
        match token.literal
        Set ParseObject = baseconst
    End If
End Function
Sub ParseNamedObject()

End Sub
Sub ParseXref()

End Sub

Function match(lit As String) As Boolean
    If (lit = token.literal) Then
        match = True
        lexer.lexan token
    Else
        match = False
    End If
End Function
