VERSION 5.00
Begin VB.Form Debuger 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debug Window"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Debuger.frx":0000
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "Debuger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DWrite(ByVal NewText As String)
  Text1.Text = Text1.Text & NewText & Chr(13) & Chr(10)
  Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Form_Load()
  Me.Top = frmMain.Height
  Me.Left = 0
  Me.Width = frmMain.Width
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = True
  Me.Visible = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
