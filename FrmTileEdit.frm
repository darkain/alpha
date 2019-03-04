VERSION 5.00
Begin VB.Form FrmTileEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Darkain Map Editor: Tile Editor"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BackPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   35
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   195
      Left            =   3960
      TabIndex        =   34
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   195
      Left            =   3960
      TabIndex        =   33
      Top             =   4920
      Width           =   855
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   31
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   32
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   30
      Left            =   3360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   31
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   29
      Left            =   3120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   30
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   28
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   29
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   27
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   28
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   26
      Left            =   2400
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   25
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   26
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   24
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   25
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   23
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   24
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   22
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   21
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   22
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   20
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   19
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   18
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   17
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   16
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   15
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   14
      Left            =   3360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   13
      Left            =   3120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   12
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   11
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   10
      Left            =   2400
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Colour 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      ClipControls    =   0   'False
      Height          =   4840
      Left            =   0
      ScaleHeight     =   319
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   0
      Top             =   0
      Width           =   4840
      Begin VB.Line Line1 
         Index           =   63
         X1              =   0
         X2              =   320
         Y1              =   259
         Y2              =   259
      End
      Begin VB.Line Line1 
         Index           =   62
         X1              =   0
         X2              =   320
         Y1              =   269
         Y2              =   269
      End
      Begin VB.Line Line1 
         Index           =   61
         X1              =   0
         X2              =   320
         Y1              =   279
         Y2              =   279
      End
      Begin VB.Line Line1 
         Index           =   60
         X1              =   0
         X2              =   320
         Y1              =   289
         Y2              =   289
      End
      Begin VB.Line Line1 
         Index           =   59
         X1              =   0
         X2              =   320
         Y1              =   299
         Y2              =   299
      End
      Begin VB.Line Line1 
         Index           =   58
         X1              =   0
         X2              =   320
         Y1              =   249
         Y2              =   249
      End
      Begin VB.Line Line1 
         Index           =   57
         X1              =   0
         X2              =   320
         Y1              =   309
         Y2              =   309
      End
      Begin VB.Line Line1 
         Index           =   55
         X1              =   0
         X2              =   320
         Y1              =   179
         Y2              =   179
      End
      Begin VB.Line Line1 
         Index           =   54
         X1              =   0
         X2              =   320
         Y1              =   189
         Y2              =   189
      End
      Begin VB.Line Line1 
         Index           =   53
         X1              =   0
         X2              =   320
         Y1              =   199
         Y2              =   199
      End
      Begin VB.Line Line1 
         Index           =   52
         X1              =   0
         X2              =   320
         Y1              =   209
         Y2              =   209
      End
      Begin VB.Line Line1 
         Index           =   51
         X1              =   0
         X2              =   320
         Y1              =   219
         Y2              =   219
      End
      Begin VB.Line Line1 
         Index           =   50
         X1              =   0
         X2              =   320
         Y1              =   169
         Y2              =   169
      End
      Begin VB.Line Line1 
         Index           =   49
         X1              =   0
         X2              =   320
         Y1              =   229
         Y2              =   229
      End
      Begin VB.Line Line1 
         Index           =   48
         X1              =   0
         X2              =   320
         Y1              =   239
         Y2              =   239
      End
      Begin VB.Line Line1 
         Index           =   47
         X1              =   0
         X2              =   320
         Y1              =   159
         Y2              =   159
      End
      Begin VB.Line Line1 
         Index           =   46
         X1              =   0
         X2              =   320
         Y1              =   149
         Y2              =   149
      End
      Begin VB.Line Line1 
         Index           =   45
         X1              =   0
         X2              =   320
         Y1              =   89
         Y2              =   89
      End
      Begin VB.Line Line1 
         Index           =   44
         X1              =   0
         X2              =   320
         Y1              =   139
         Y2              =   139
      End
      Begin VB.Line Line1 
         Index           =   43
         X1              =   0
         X2              =   320
         Y1              =   129
         Y2              =   129
      End
      Begin VB.Line Line1 
         Index           =   42
         X1              =   0
         X2              =   320
         Y1              =   119
         Y2              =   119
      End
      Begin VB.Line Line1 
         Index           =   41
         X1              =   0
         X2              =   320
         Y1              =   109
         Y2              =   109
      End
      Begin VB.Line Line1 
         Index           =   40
         X1              =   0
         X2              =   320
         Y1              =   99
         Y2              =   99
      End
      Begin VB.Line Line1 
         Index           =   38
         X1              =   0
         X2              =   320
         Y1              =   79
         Y2              =   79
      End
      Begin VB.Line Line1 
         Index           =   33
         X1              =   0
         X2              =   320
         Y1              =   69
         Y2              =   69
      End
      Begin VB.Line Line1 
         Index           =   39
         X1              =   0
         X2              =   320
         Y1              =   9
         Y2              =   9
      End
      Begin VB.Line Line1 
         Index           =   37
         X1              =   0
         X2              =   320
         Y1              =   59
         Y2              =   59
      End
      Begin VB.Line Line1 
         Index           =   36
         X1              =   0
         X2              =   320
         Y1              =   49
         Y2              =   49
      End
      Begin VB.Line Line1 
         Index           =   35
         X1              =   0
         X2              =   320
         Y1              =   39
         Y2              =   39
      End
      Begin VB.Line Line1 
         Index           =   30
         X1              =   0
         X2              =   320
         Y1              =   29
         Y2              =   29
      End
      Begin VB.Line Line1 
         Index           =   34
         X1              =   0
         X2              =   320
         Y1              =   19
         Y2              =   19
      End
      Begin VB.Line Line1 
         Index           =   32
         X1              =   299
         X2              =   299
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   31
         X1              =   309
         X2              =   309
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   29
         X1              =   259
         X2              =   259
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   28
         X1              =   269
         X2              =   269
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   27
         X1              =   279
         X2              =   279
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   26
         X1              =   289
         X2              =   289
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   25
         X1              =   249
         X2              =   249
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   24
         X1              =   69
         X2              =   69
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   23
         X1              =   209
         X2              =   209
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   22
         X1              =   219
         X2              =   219
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   21
         X1              =   229
         X2              =   229
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   20
         X1              =   239
         X2              =   239
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   19
         X1              =   169
         X2              =   169
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   18
         X1              =   179
         X2              =   179
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   189
         X2              =   189
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   199
         X2              =   199
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   129
         X2              =   129
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   139
         X2              =   139
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   149
         X2              =   149
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   159
         X2              =   159
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   89
         X2              =   89
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   99
         X2              =   99
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   109
         X2              =   109
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   119
         X2              =   119
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   49
         X2              =   49
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   59
         X2              =   59
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   79
         X2              =   79
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   39
         X2              =   39
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   29
         X2              =   29
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   19
         X2              =   19
         Y1              =   0
         Y2              =   320
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   9
         X2              =   9
         Y1              =   0
         Y2              =   320
      End
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   4800
   End
End
Attribute VB_Name = "FrmTileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MouseDown As Boolean
Dim CurrentColour As ColorConstants
Dim Pic(32, 32) As ColorConstants

Sub DrawTile(X As Single, Y As Single)
Dim TmpX As Integer, TmpY As Integer
TmpX = Int(X / 10)
TmpY = Int(Y / 10)
BackPic.AutoRedraw = True
BackPic.PSet (TmpX, TmpY), CurrentColour
BackPic.AutoRedraw = False
Picture1.Line (TmpX * 10, TmpY * 10)-(TmpX * 10 + 8, TmpY * 10 + 8), CurrentColour, BF
'Call BitBlt(Picture1.hDC, TmpX * 10, TmpY * 10, 9, 9, Colour(5).hDC, 0, 0, SRCCOPY)
End Sub

Private Sub Colour_Click(Index As Integer)
CurrentColour = Colour(Index).BackColor
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 15
  Colour(i).BackColor = QBColor(i)
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DrawTile(X, Y)
MouseDown = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 0 Or Y < 0 Or X > Picture1.ScaleWidth Or Y > Picture1.ScaleHeight Then Exit Sub
If MouseDown Then Call DrawTile(X, Y)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown = False
End Sub

Private Sub Picture1_Paint()
BackPic.AutoRedraw = True
Call StretchBlt(Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, BackPic.hDC, 0, 0, 32, 32, SRCCOPY)
BackPic.AutoRedraw = False
End Sub
