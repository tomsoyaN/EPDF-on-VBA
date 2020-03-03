Attribute VB_Name = "Module1"
Public sym As New t_symtalbe
Sub execute()
    Dim obj As parser
    Set obj = New parser
    sym.BuildSymTable
    obj.Initialize sym
    obj.Parse
End Sub
