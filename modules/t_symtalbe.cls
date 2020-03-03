VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "t_symtalbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public symtable As Object
Public Sub BuildSymTable()
    Dim tok As t_token
    Set symtable = CreateObject("Scripting.Dictionary")
    Set tok = CreateToken(0, "DIC_S")
    symtable.Add "<<", tok
    Set tok = CreateToken(0, "DIC_E")
    symtable.Add ">>", tok
    Set tok = CreateToken(0, "BIN_S")
    symtable.Add "<", tok
    Set tok = CreateToken(0, "BIN_E")
    symtable.Add ">", tok
    Set tok = CreateToken(0, "")
    symtable.Add "<?", tok
    Set tok = CreateToken(0, "")
    symtable.Add "?>", tok
    Set tok = CreateToken(0, "SLASH")
    symtable.Add "/", tok
    Set tok = CreateToken(0, "ARRAY_S")
    symtable.Add "[", tok
    Set tok = CreateToken(0, "ARRAY_E")
    symtable.Add "]", tok
    Set tok = CreateToken(0, "STRING_S")
    symtable.Add "(", tok
    Set tok = CreateToken(0, "STRING_E")
    symtable.Add ")", tok
    Set tok = CreateToken(0, "OBJ_S")
    symtable.Add "obj", tok
    Set tok = CreateToken(0, "OBJ_E")
    symtable.Add "endobj", tok
    Set tok = CreateToken(0, "STREAM_S")
    symtable.Add "stream", tok
    Set tok = CreateToken(0, "STREAM_E")
    symtable.Add "endstream", tok
End Sub
Public Function CreateToken(num As Integer, lit As String) As t_token
    Dim res As New t_token
    res.num = num
    res.literal = lit
    Set CreateToken = res
End Function

