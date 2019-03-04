VERSION 5.00
Begin VB.UserControl DAN_Info 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   ScaleHeight     =   5760
   ScaleWidth      =   7350
   Begin VB.TextBox TxHead 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TxHead 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox TxHead 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox TxHead 
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TxHead 
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TxHead 
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      X1              =   2280
      X2              =   2280
      Y1              =   2760
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   2760
      Y2              =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "-=- Header -=-"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label LbHead 
      Alignment       =   1  'Right Justify
      Caption         =   "Title"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Label LbHead 
      Alignment       =   1  'Right Justify
      Caption         =   "Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   120
      MouseIcon       =   "DAN_Info.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.Label LbHead 
      Alignment       =   1  'Right Justify
      Caption         =   "Script"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   120
      MouseIcon       =   "DAN_Info.ctx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label LbHead 
      Alignment       =   1  'Right Justify
      Caption         =   "Tileset"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label LbHead 
      Alignment       =   1  'Right Justify
      Caption         =   "Fontset"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label LbHead 
      Alignment       =   1  'Right Justify
      Caption         =   "Pallet"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
End
Attribute VB_Name = "DAN_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim TheFileName As String

Private Sub LoadMap()
  Dim a As String
  Dim Section As Integer
  Dim StrLoc As Integer
  Dim i As Integer
  On Error GoTo ErrHan
  
  Open TheFileName For Input As #1
    Cls
    Do
      Input #1, a
      If InStr(a, ";") Then
        a = Left(a, InStr(a, ";") - 1)
      End If
            
      a = Trim$(a)
      If Left(a, 1) = "[" And Right(a, 1) = "]" Then
        a = UCase(Mid(a, 2, Len(a) - 2))
        Select Case a
          Case "HEADER"
            Section = 1
          Case "TELEPORT"
            Section = 2
            MapInfo.Telepoters = MapInfo.Telepoters + 1
            ReDim Preserve MapInfo.Teleport(MapInfo.Telepoters)
        End Select
      
      Else
        StrLoc = InStr(a, "=")
        Select Case Section
          Case 1
            If StrLoc Then
              Select Case Trim(UCase(Left(a, StrLoc - 1)))
                Case "MAP"
                  MapInfo.MapHeader.Map = Mid(a, StrLoc + 1)
                Case "TITLE"
                  MapInfo.MapHeader.Title = Mid(a, StrLoc + 1)
                Case "CODESCRIPT"
                  MapInfo.MapHeader.CodeScript = Mid(a, StrLoc + 1)
                Case "TILESET"
                  MapInfo.MapHeader.TileSet = Mid(a, StrLoc + 1)
                Case "FONTSET"
                  MapInfo.MapHeader.FontSet = Mid(a, StrLoc + 1)
                Case "PALLET"
                  MapInfo.MapHeader.Pallet = Mid(a, StrLoc + 1)
              End Select
            End If
          
          Case 2
            If StrLoc Then
              With MapInfo.Teleport(MapInfo.Telepoters)
                Select Case Trim(UCase(Left(a, StrLoc - 1)))
                  Case "SRCX1"
                    .SrcX1 = CInt(Mid(a, StrLoc + 1))
                  Case "SRCY1"
                    .SrcY1 = CInt(Mid(a, StrLoc + 1))
                  Case "SRCX2"
                    .SrcX2 = CInt(Mid(a, StrLoc + 1))
                  Case "SRCY2"
                    .SrcY2 = CInt(Mid(a, StrLoc + 1))
                  Case "DESTX"
                    .DestX = CInt(Mid(a, StrLoc + 1))
                  Case "DESTY"
                    .DestY = CInt(Mid(a, StrLoc + 1))
                  Case "DESTDIR"
                    .DestDir = CInt(Mid(a, StrLoc + 1))
                  Case "DESTLAYER"
                    .DestLayer = CInt(Mid(a, StrLoc + 1))
                  Case "DESTMAP"
                    .DestMap = Mid(a, StrLoc + 1)
                  Case "TRANSITION"
                    .Transition = CInt(Mid(a, StrLoc + 1))
                  Case "SRCLAYERMIN"
                    .srcLayerMin = CInt(Mid(a, StrLoc + 1))
                  Case "SRCLAYERMAX"
                    .srcLayerMax = CInt(Mid(a, StrLoc + 1))
                End Select
              End With
            End If
        End Select
      End If
    Loop Until EOF(1)
  
  Close #1
  FileHasChanged = False
  
Exit Sub

ErrHan:
  MsgBox Err.Description
  Err.Clear
End Sub

Public Property Let FileName(ByVal TheFile As String)
  TheFileName = TheFile
  ChangeDirectory (TheFile)
  LoadMap
  
  TxHead(0).Text = MapInfo.MapHeader.Title
  TxHead(1).Text = MapInfo.MapHeader.Map
  TxHead(2).Text = MapInfo.MapHeader.CodeScript
  TxHead(3).Text = MapInfo.MapHeader.TileSet
  TxHead(4).Text = MapInfo.MapHeader.FontSet
  TxHead(5).Text = MapInfo.MapHeader.Pallet
End Property

Private Sub LbHead_Click(Index As Integer)
  Select Case Index
    Case 1
      Load_DAM (GameMain.Paths.Maps & MapInfo.MapHeader.Map)
      
    Case 2
      Load_DAI (GameMain.Paths.Script & MapInfo.MapHeader.CodeScript)
  
  End Select
End Sub
