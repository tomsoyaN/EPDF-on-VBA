VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NamedObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public num As Integer
Public generation As Integer
Private stream() As Byte
Public length As Integer
Public objs As Collection


Public Sub Initialize(n As Integer, g As Integer)
    Set objs = New Collection
    num = n
    generation = g
End Sub
