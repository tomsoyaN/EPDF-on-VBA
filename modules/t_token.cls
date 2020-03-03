VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "t_token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public num As Long
Public literal As String
Private buf() As Byte

Public Sub Clone(ByVal token As t_token)
    num = token.num
    literal = token.literal
End Sub
