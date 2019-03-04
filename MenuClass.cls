VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MenuClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ID_Number As Integer

Public Property Get Text(ByVal LineNumber As Integer) As String
On Error GoTo ErrHan
  Text = MenuX(ID_Number).Text(LineNumber).Text
Exit Property

ErrHan:
  Debuger.DWrite Err.Number & " (MenuClass.TextGet) : " & Err.Source & " : " & Err.Description
  Err.Clear
End Property

Public Property Let Text(ByVal LineNumber As Integer, ByVal NewText As String)
On Error GoTo ErrHan
  MenuX(ID_Number).Text(LineNumber).Text = NewText
Exit Property

ErrHan:
  Debuger.DWrite Err.Number & " (MenuClass.TextLet) : " & Err.Source & " : " & Err.Description
  Err.Clear
End Property

Public Property Get TextColor(ByVal LineNumber As Integer) As Long
On Error GoTo ErrHan
  TextColor = MenuX(ID_Number).Text(LineNumber).Colour
Exit Property

ErrHan:
  Debuger.DWrite Err.Number & " (MenuClass.TextColorGet) : " & Err.Source & " : " & Err.Description
  Err.Clear
End Property

Public Property Let TextColor(ByVal LineNumber As Integer, ByVal NewColour As Long)
On Error GoTo ErrHan
  MenuX(ID_Number).Text(LineNumber).Colour = NewColour
Exit Property

ErrHan:
  Debuger.DWrite Err.Number & " (MenuClass.TextColorLet) : " & Err.Source & " : " & Err.Description
  Err.Clear
End Property

Public Sub Update()
On Error GoTo ErrHan
  Call UpdateMenuSystem
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (MenuClass.Update) : " & Err.Source & " : " & Err.Description
  Err.Clear
End Sub

Public Property Get Visible() As Boolean
On Error GoTo ErrHan
  Visible = MenuX(ID_Number).Visible
Exit Property

ErrHan:
  Debuger.DWrite Err.Number & " (MenuClass.VisibleGet) : " & Err.Source & " : " & Err.Description
  Err.Clear
End Property

Public Property Let Visible(ByVal Vis As Boolean)
On Error GoTo ErrHan
  MenuX(ID_Number).Visible = Vis
Exit Property

ErrHan:
  Debuger.DWrite Err.Number & " (MenuClass.VisibleLet) : " & Err.Source & " : " & Err.Description
  Err.Clear
End Property
