VERSION 5.00
Begin VB.UserControl MidiPlayer 
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   ScaleHeight     =   5700
   ScaleWidth      =   7260
   Begin VB.CommandButton MidiPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton MidiStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label LabelName 
      Caption         =   "File Name: "
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "MidiPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub MidiPlay_Click()
  Dim Ret As Integer
  
  Ret = mciSendString("seek DarkainMidi to start", 0&, 0, 0)
  Ret = mciSendString("play DarkainMidi", 0&, 0, 0)
End Sub

Private Sub MidiStop_Click()
  Dim Ret As Integer
  
  Ret = mciSendString("stop DarkainMidi", 0&, 0, 0)
End Sub

Public Property Let FileName(ByVal TheName As String)
  Ret = mciSendString("open " & TheName & " type sequencer alias DarkainMidi", 0&, 0, 0)
  LabelName.Caption = "File Name: " & TheName
End Property

Private Sub UserControl_Terminate()
  Ret = mciSendString("close DarkainMidi", 0&, 0, 0)
End Sub
