VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMap 
   Caption         =   "Darkain Map Editor"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11145
   FillStyle       =   0  'Solid
   Icon            =   "Editor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Map2 
      DrawMode        =   1  'Blackness
      Height          =   5115
      Left            =   1800
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   0
      Top             =   360
      Width           =   5595
      Begin VB.HScrollBar MapScrollH 
         Height          =   255
         LargeChange     =   9
         Left            =   0
         Max             =   20
         TabIndex        =   2
         Top             =   4800
         Width           =   5295
      End
      Begin VB.VScrollBar MapScrollV 
         Height          =   4815
         LargeChange     =   8
         Left            =   5280
         Max             =   20
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Useless 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5280
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   16
         Top             =   4800
         Width           =   375
      End
      Begin VB.Line TileSelectLine 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         DrawMode        =   14  'Copy Pen
         Index           =   7
         Visible         =   0   'False
         X1              =   8
         X2              =   8
         Y1              =   8
         Y2              =   37
      End
      Begin VB.Line TileSelectLine 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         DrawMode        =   14  'Copy Pen
         Index           =   6
         Visible         =   0   'False
         X1              =   37
         X2              =   37
         Y1              =   8
         Y2              =   37
      End
      Begin VB.Line TileSelectLine 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         DrawMode        =   14  'Copy Pen
         Index           =   5
         Visible         =   0   'False
         X1              =   8
         X2              =   37
         Y1              =   8
         Y2              =   8
      End
      Begin VB.Line TileSelectLine 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         DrawMode        =   14  'Copy Pen
         Index           =   4
         Visible         =   0   'False
         X1              =   8
         X2              =   37
         Y1              =   37
         Y2              =   37
      End
   End
   Begin MSComctlLib.StatusBar Bar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   6060
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Object.ToolTipText     =   "Current X and Y Location"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Currently Selected Tile/Obstruction"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "TBarPics"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Map"
            ImageIndex      =   18
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Picture Map (DAM)"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Data Map (DAN)"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Information (DAI)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Map"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Map"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Layer1"
            Object.ToolTipText     =   "Enable/Disable Layer 1"
            ImageIndex      =   8
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Layer2"
            Object.ToolTipText     =   "Enable/Disable Layer 2"
            ImageIndex      =   9
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Layer3"
            Object.ToolTipText     =   "Enable/Disable Layer 3"
            ImageIndex      =   10
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Layer4"
            Object.ToolTipText     =   "Enable/Disable Layer 4"
            ImageIndex      =   11
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grid"
            Object.ToolTipText     =   "Enable/Disable Grid Lines"
            ImageIndex      =   12
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Obst"
            Object.ToolTipText     =   "Enable/Disable Unbstruction Layer"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tiles"
            Object.ToolTipText     =   "Enable/Disable Tile Viewer"
            ImageIndex      =   14
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anim"
            Object.ToolTipText     =   "Enable/Disable Animation"
            ImageIndex      =   17
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.Frame ToolsFrm 
      Caption         =   "Tools"
      Height          =   3135
      Left            =   9960
      TabIndex        =   28
      Top             =   2280
      Width           =   1095
      Begin VB.PictureBox ToolsPic1 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   0
         ScaleHeight     =   2775
         ScaleWidth      =   975
         TabIndex        =   29
         Top             =   240
         Width           =   975
         Begin VB.VScrollBar ToolsScroll 
            Height          =   2655
            LargeChange     =   2
            Left            =   720
            Max             =   13
            Min             =   1
            TabIndex        =   30
            Top             =   0
            Value           =   1
            Width           =   255
         End
         Begin MSComctlLib.ImageList ToolsImgLst 
            Left            =   0
            Top             =   2160
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   8388736
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   13
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":0CCA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":10BD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":1487
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":1868
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":1C1F
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":1CCA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":209F
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":2488
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":285A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":2C29
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":2FD2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":33C8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Editor.frx":37BD
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar ToolsBar 
            Height          =   7410
            Left            =   120
            TabIndex        =   31
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   13070
            ButtonWidth     =   1984
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ToolsImgLst"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   13
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Draw"
                  Object.ToolTipText     =   "Draw Tiles"
                  ImageIndex      =   1
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "FloodFill"
                  Object.ToolTipText     =   "Flood Fill within selection"
                  ImageIndex      =   6
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   2
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "a"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "b"
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Cut"
                  Object.ToolTipText     =   "Cut tile(s)"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Copy"
                  Object.ToolTipText     =   "Copy tile(s)"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Paste"
                  Object.ToolTipText     =   "Paste tile(s)"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Delete"
                  Object.ToolTipText     =   "Delete tile(s)"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Insert"
                  Object.ToolTipText     =   "Insert ???"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Move"
                  Object.ToolTipText     =   "Move selected tile(s)"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Pick Up"
                  Object.ToolTipText     =   "Pick Up tile(s) for dawing"
                  ImageIndex      =   9
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Select"
                  Object.ToolTipText     =   "Select tile(s) on map"
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Zoom11"
                  Object.ToolTipText     =   "Zoom 1:1 (Normal)"
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Zoom21"
                  Object.ToolTipText     =   "Zoom 2:1 (In)"
                  ImageIndex      =   12
               EndProperty
               BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "Zoom12"
                  Object.ToolTipText     =   "Zoom 1:2 (Out)"
                  ImageIndex      =   13
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList ImgToolz 
      Left            =   9480
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   67
      ImageHeight     =   207
      MaskColor       =   8388736
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":3BBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList NodePics2 
      Left            =   0
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":596A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":65BC
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":6ADE
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":7000
            Key             =   "open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":7522
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":7B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":81E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":91BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList NodePics 
      Left            =   0
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":A18E
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":A6B0
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":ABD2
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":B0F4
            Key             =   "open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":B616
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":BC78
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":C2DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":D2AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame MainFrame 
      Caption         =   "Midi File"
      Height          =   4335
      Left            =   1800
      TabIndex        =   22
      Top             =   360
      Width           =   5655
      Begin VB.TextBox EditINI 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   26
         Top             =   3000
         Width           =   1215
      End
      Begin VB.PictureBox PicView 
         AutoSize        =   -1  'True
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1755
         ScaleWidth      =   2955
         TabIndex        =   25
         Top             =   2400
         Width           =   3015
      End
      Begin Darkain_Map_Editor.DAN_Info DAN_Info 
         Height          =   3015
         Left            =   3720
         TabIndex        =   24
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5318
      End
      Begin Darkain_Map_Editor.MidiPlayer MidiPlayer 
         Height          =   2175
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3836
      End
   End
   Begin VB.Frame LayerFrm 
      Caption         =   "Edit Layer"
      Height          =   1815
      Left            =   9960
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.Label LbLayer 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label LbLayer 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LbLayer 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.Label LbLayer 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox TileSet2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Shade 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   3840
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame TileFrm 
      Caption         =   "Tiles"
      Height          =   5115
      Left            =   7440
      TabIndex        =   3
      Top             =   360
      Width           =   2415
      Begin VB.VScrollBar TileScroll 
         Height          =   4335
         LargeChange     =   7
         Left            =   2040
         Max             =   248
         TabIndex        =   5
         Top             =   665
         Width           =   255
      End
      Begin VB.ComboBox TileCat 
         Height          =   315
         ItemData        =   "Editor.frx":E282
         Left            =   120
         List            =   "Editor.frx":E284
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
      Begin VB.PictureBox TileView 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   120
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   4
         Top             =   665
         Width           =   1920
         Begin VB.Line TileSelectLine 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   3
            X1              =   1
            X2              =   30
            Y1              =   30
            Y2              =   30
         End
         Begin VB.Line TileSelectLine 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   2
            X1              =   1
            X2              =   30
            Y1              =   1
            Y2              =   1
         End
         Begin VB.Line TileSelectLine 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   0
            X1              =   30
            X2              =   30
            Y1              =   1
            Y2              =   30
         End
         Begin VB.Line TileSelectLine 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   1
            X1              =   1
            X2              =   1
            Y1              =   1
            Y2              =   30
         End
      End
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   7200
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList TBarPics 
      Left            =   9720
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":E286
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":E39A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":E4AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":E5C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":E6D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":E7EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":E8FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":EA12
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":EB26
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":EC3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":ED4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":EE62
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":EF7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":F09A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":F1B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":F2CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":F3DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":F53E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":F69E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":F7FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox AnimSet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox AnimMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer AnimTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   360
   End
   Begin MSComctlLib.TreeView FileList 
      Height          =   4215
      Left            =   0
      TabIndex        =   21
      Top             =   360
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   7435
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "NodePics"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox TileMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   7080
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox DarkainLogo 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      Picture         =   "Editor.frx":F912
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   27
      Top             =   4680
      Width           =   615
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu MnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu MnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu MnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu MnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu MnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu MnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu MnuEditDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu MnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu MnuGrid 
         Caption         =   "&Grid Lines"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuObst 
         Caption         =   "&Obstruction"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuTiles 
         Caption         =   "&TIles"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu MnuAnim 
         Caption         =   "&Animation"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MnuLayer 
      Caption         =   "&Layers"
      Begin VB.Menu MnuLayers 
         Caption         =   "&1"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuLayers 
         Caption         =   "&2"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuLayers 
         Caption         =   "&3"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuLayers 
         Caption         =   "&4"
         Checked         =   -1  'True
         Index           =   3
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu MnuTile 
      Caption         =   "&Tile"
      Visible         =   0   'False
      Begin VB.Menu MnuTileZoom 
         Caption         =   "&Zoom Preview"
         Begin VB.Menu MnuTileZoom21 
            Caption         =   "2/1"
         End
         Begin VB.Menu MnuTileZoom11 
            Caption         =   "1/1"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuTileZoom12 
            Caption         =   "1/2"
         End
      End
   End
   Begin VB.Menu MnuRightMouse 
      Caption         =   "RightMouse"
      Visible         =   0   'False
      Begin VB.Menu MnuDAN 
         Caption         =   "DAN"
         Begin VB.Menu MnuDAN_Open 
            Caption         =   "Open"
         End
         Begin VB.Menu MnuDAN_Text 
            Caption         =   "Text Edit"
         End
      End
   End
End
Attribute VB_Name = "FrmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseDown           As Integer
Dim CurrentShade        As Integer
Dim CurrentLayer        As Integer
Dim AnimX               As Integer


Public Sub MakeVis(ByVal TheNumber As Integer)
  ActiveWin = TheNumber
  Call Form_Resize
  
  TBar.Buttons(TBar_Button_Save).Enabled = False
  TBar.Buttons(TBar_Button_Layer1).Visible = False
  TBar.Buttons(TBar_Button_Layer2).Visible = False
  TBar.Buttons(TBar_Button_Layer3).Visible = False
  TBar.Buttons(TBar_Button_Layer4).Visible = False
  TBar.Buttons(TBar_Button_Grid).Visible = False
  TBar.Buttons(TBar_Button_Obst).Visible = False
  TBar.Buttons(TBar_Button_Tile).Visible = False
  TBar.Buttons(TBar_Button_Anim).Visible = False
  TBar.Buttons(TBar_Button_Sep2).Visible = False
  
  FileHasChanged = False
  
  MainFrame.Visible = False
  
  MnuOptions.Visible = False
  MnuEdit.Visible = False
  MnuLayer.Visible = False
  
  DAN_Info.Visible = False
  MidiPlayer.Visible = False
  
  Map2.Visible = False
  TileFrm.Visible = False
  LayerFrm.Visible = False
  ToolsFrm.Visible = False
'  AnimFrm.Visible = False
  
  PicView.Visible = False
  EditINI.Visible = False
  
  DarkainLogo.Visible = False
  
  Select Case TheNumber
    Case Win_DAN
      TBar.Buttons(TBar_Button_Save).Enabled = True
      
      MnuEdit.Visible = True
      MainFrame.Visible = True
      DAN_Info.Visible = True
      MainFrame.Caption = "DAN Map File Information"
      DAN_Info.SetFocus
    
    Case Win_MapEdit
      TBar.Buttons(TBar_Button_Save).Enabled = True
      TBar.Buttons(TBar_Button_Layer1).Visible = True
      TBar.Buttons(TBar_Button_Layer2).Visible = True
      TBar.Buttons(TBar_Button_Layer3).Visible = True
      TBar.Buttons(TBar_Button_Layer4).Visible = True
      TBar.Buttons(TBar_Button_Grid).Visible = True
      TBar.Buttons(TBar_Button_Obst).Visible = True
      TBar.Buttons(TBar_Button_Tile).Visible = True
      TBar.Buttons(TBar_Button_Anim).Visible = True
      TBar.Buttons(TBar_Button_Sep2).Visible = True
      
      MnuOptions.Visible = True
      MnuLayer.Visible = True
      Map2.Visible = True
      MapScrollH.Value = -(Map2.ScaleWidth \ 64)
      MapScrollV.Value = -(Map2.ScaleHeight \ 64)
      TileFrm.Visible = True
      LayerFrm.Visible = True
      TileScroll_Change

      Loading = True
      DrawTool = Tools_Draw
      ToolsFrm.Visible = True
      ToolsScroll.Value = 1
      Call ToolsScroll_Change
      ToolsBar.Refresh
      Loading = False
'      AnimFrm.Visible = True
      Map2.SetFocus
    
    Case Win_MenuEdit
      TBar.Buttons(TBar_Button_Save).Enabled = True
      TBar.Buttons(TBar_Button_Grid).Visible = True
      Map2.Visible = True
      TileFrm.Visible = True
      Map2.SetFocus
    
    
    Case Win_MIDI
      MainFrame.Visible = True
      MidiPlayer.Visible = True
      MainFrame.Caption = "MIDI Music Player"
      MidiPlayer.SetFocus
      
    Case Win_PicView
      MainFrame.Visible = True
      PicView.Visible = True
      PicView.SetFocus
      
    Case Win_TextEdit
      TBar.Buttons(TBar_Button_Save).Enabled = True
      MnuEdit.Visible = True
      
      MainFrame.Visible = True
      EditINI.Visible = True
      EditINI.SetFocus
      
    Case Win_Logo
      DarkainLogo.Visible = True
      
  End Select
End Sub

Private Sub FindFilez(MainDir As String, Key As String)
On Error GoTo ErrHan
  
  Dim a As String
  Dim i As Integer
  Dim Folderz(100) As String
  Dim Filez(100) As String
  Dim CurFolder As Integer
  Dim CurFile As Integer
  CurFolder = -1
  CurFile = -1
  
  a = Dir(MainDir, vbDirectory)
  Do Until Len(a) = 0
    If a <> "." And a <> ".." Then
      If (GetAttr(MainDir & a) And vbDirectory) = vbDirectory Then
        CurFolder = CurFolder + 1
        Folderz(CurFolder) = a
      Else
        CurFile = CurFile + 1
        Filez(CurFile) = a
      End If
    End If
    a = Dir
  Loop

  If CurFolder <> -1 Then
    For i = 0 To CurFolder
      If Key = "" Then
        FileList.Nodes.Add , , Key & Folderz(i), Folderz(i), 1
      Else
        FileList.Nodes.Add Key, tvwChild, Key & Folderz(i), Folderz(i), 1
      End If
      Call FindFilez(MainDir & Folderz(i) & "\", Key & Folderz(i))
    Next i
  End If
  
  If CurFile <> -1 Then
    For i = 0 To CurFile
      Select Case UCase(Right(Filez(i), 4))
        Case ".DAM"
          FileList.Nodes.Add Key, tvwChild, , Filez(i), 7
        Case ".DAN"
          FileList.Nodes.Add Key, tvwChild, , Filez(i), 8
        Case Else
          FileList.Nodes.Add Key, tvwChild, , Filez(i), 3
      End Select
    Next i
  End If
Exit Sub

ErrHan:
End Sub

Public Sub SetRedraw(DrawMode As Boolean)
  Dim i As Integer
  If Not DrawMode Then Map2.Cls
  Map2.AutoRedraw = DrawMode
  Map2_Paint
  TileScroll_Change
End Sub

Private Sub AnimTimer_Timer()
  AnimX = AnimX + 1
  If AnimX = 32 Then AnimX = 0
  Map2_Paint
End Sub

Private Sub EditINI_Change()
  If Loading Then Exit Sub
  FileHasChanged = True
End Sub

Private Sub FileList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  LastButton = Button
End Sub

Private Sub FileList_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim TheReturn As Integer
  Dim Ret As Integer
  
  NodeString = Node.FullPath
  If LastButton = 2 Then
    Select Case UCase(Right(Node.Text, 4))
      Case ".DAN"
        PopupMenu MnuDAN
    End Select
    Exit Sub
  End If
  
  
  
  If Node.Children Then
    Node.Expanded = Not Node.Expanded
  Else
    If FileHasChanged Then
      TheReturn = MsgBox("Save before changing files?", 35, "Darkain's Alhpa Editor")
    
      Select Case TheReturn
        
        Case 6 'Yes
          Call MnuFileSave_Click
        
        Case 7 'No
        
        Case Else 'Other
          Node.Selected = False
          Exit Sub
      End Select
    End If
    
    Ret = mciSendString("close DarkainMidi", 0&, 0, 0)
  
    Select Case UCase(Right(Node.Text, 4))
      Case ".DAN" 'Text Map Files
        Load_DAN (GameMain.Paths.Root & Node.FullPath)
      
      Case ".DAM" 'Picture Map Files
        Load_DAM (GameMain.Paths.Root & Node.FullPath)
      
      Case ".MID" 'MIDI Music Audio
        MidiPlayer.FileName = GameMain.Paths.Root & Node.FullPath
        Caption = FrmCaption & "  --  " & Node.FullPath
        MakeVis (Win_MIDI)
      
      Case ".BMP", ".JPG", ".GIF"
        PicView.Picture = LoadPicture(GameMain.Paths.Root & Node.FullPath)
        MainFrame.Caption = GameMain.Paths.Root & Node.FullPath
        Caption = FrmCaption & "  --  " & Node.FullPath
        MakeVis (Win_PicView)
      
      Case ".DAI", ".DAC", ".DAY"
        Load_DAI (GameMain.Paths.Root & Node.FullPath)
        
      Case Else
        Caption = FrmCaption & "  --  " & Node.FullPath
        MakeVis (Win_Logo)
    End Select
  End If
End Sub

Private Sub Form_Load()
  Dim DestDC As Long
  Dim DestBMP As Long
  Dim DestX As Long
  
  Dim i As Integer
  
  'start loading
  Loading = True
  
  'create a new file
  New_File
  
  'locate the primary INI file, and load it
  If FileExist(App.Path & "\Darkain.DAI") Then
    GameMain.Paths.Main = App.Path & "\"
    Call LoadINI(App.Path & "\Darkain.DAI")
  ElseIf FileExist(App.Path & "\System\Darkain.DAI") Then
    GameMain.Paths.Main = App.Path & "\System\"
    Call LoadINI(App.Path & "\System\Darkain.DAI")
  Else
    MsgBox "File Not Found - Darkain.DAI"
    End
  End If

  'set the COMDLG default dir to the map's path
  Dlg.InitDir = GameMain.Paths.Maps
  
  'set the left property to where it should be
  Map2.Left = FileList.Width + 45
  
  'everything relating to ANIMSET removed because of the addition of
  'multiple tilesets accessable with tile selecter.
  
  'Load tiles into picturboxes
  TileSet2.Picture = LoadPicture(GameMain.Paths.Grafix & "tileset.bmp")
'  AnimSet.Picture = LoadPicture(GameMain.Paths.Grafix & "WaterAnim.bmp")
  
  'Creates mask images from tiles
  Call CreateMaskImage(TileSet2, TileMask)
'  Call CreateMaskImage(AnimSet, AnimMask)
  
  'Removes pink boxes around tiles
  DestDC = CreateCompatibleDC(TileSet2.hDc)
  DestBMP = CreateCompatibleBitmap(TileSet2.hDc, TileSet2.ScaleWidth, TileSet2.ScaleHeight)
  DestX = SelectObject(DestDC, DestBMP)
  BitBlt DestDC, 0, 0, TileSet2.ScaleWidth, TileSet2.ScaleHeight, TileMask.hDc, 0, 0, NOTSRCCOPY
  BitBlt TileSet2.hDc, 0, 0, TileSet2.ScaleWidth, TileSet2.ScaleHeight, DestDC, 0, 0, SRCAND
  DestX = SelectObject(DestDC, DestBMP)
  DestX = DeleteDC(DestDC)
  DestX = DeleteObject(DestBMP)
  
  'load "shaded" used with the obstruction editor
  Shade.Picture = LoadResPicture("Shade", vbResBitmap)
  
  CurrentLayer = 0
  
  'set values for scroll bars
  MapScrollH.Max = MaxWH - (MapScrollH.LargeChange + 2)
  MapScrollV.Max = MaxWH - (MapScrollV.LargeChange + 2)
  MapScrollH.Min = -MaxWH
  MapScrollV.Min = -MaxWH
  MapScrollH.Value = -5
  MapScrollV.Value = -5
  
  'set minumim size for the form
  gHW = Me.hwnd
  Hook
  
  'load the tileset names
  Load_DAT (GameMain.Paths.System & "Main.DAT")
  TileCat.ListIndex = 0
  Call TileScroll_Change
  
  'add version information to form caption
  Caption = Caption & " - Version: " & App.Major & "." & App.Minor & "." & App.Revision
  FrmCaption = Caption
  
  'file the fileselector with files from the current project directory
  Call FindFilez(GameMain.Paths.Root, "")
  
  'move different objects on editor so they are all in same spot
  MidiPlayer.Move 120, 240
  DAN_Info.Move 120, 240
  PicView.Move 120, 240
  EditINI.Move 120, 240
  
  'reset all menu information
  ReDim MenuX(0)
  ReDim MenuX(0).Text(0)
  
  'set current EDIT to the darkain logo
  MakeVis (Win_Logo)
  
  'loading is finished
  Loading = False
End Sub

Private Sub Form_Resize()
  If WindowState = 1 Then Exit Sub
  On Error Resume Next
  Loading = True

  Dim BorderWidth As Integer

  BorderWidth = FrmMap.Width - FrmMap.ScaleWidth * Screen.TwipsPerPixelX
  
  FileList.Move FileList.Left, 390, FileList.Width, ScaleHeight - 650
  MainFrame.Move Map2.Left, 390, ScaleWidth - FileList.Width - 95, FileList.Height
  
  Select Case ActiveWin
    
    Case Win_MapEdit

      LayerFrm.Move Me.ScaleWidth - LayerFrm.Width - 30, 390
      TileFrm.Move LayerFrm.Left - TileFrm.Width, 390, TileFrm.Width, ScaleHeight - 650
      TileScroll.Move TileScroll.Left, TileScroll.Top, TileScroll.Width, ((TileFrm.Height - 465 - TileCat.Height) \ 480) * 480
      TileView.Move TileView.Left, TileView.Top, TileView.Width, TileScroll.Height
      TileScroll.Max = 64 - (TileView.ScaleHeight \ 32)
      ToolsFrm.Move LayerFrm.Left, LayerFrm.Height + LayerFrm.Top, LayerFrm.Width, TileFrm.Height - LayerFrm.Height
      
      Map2.Move Map2.Left, 390, (((ScaleWidth - 3380 - Map2.Left) \ 480) * 480 - 165), ((ScaleHeight - 480) \ 480) * 480 - 165
      MapScrollH.Move 0, Map2.ScaleHeight - MapScrollH.Height, Map2.ScaleWidth - 17
      MapScrollV.Move Map2.ScaleWidth - MapScrollV.Width, 0, MapScrollV.Width, Map2.ScaleHeight - 17
      Useless.Move MapScrollH.Width, MapScrollV.Height

      ToolsScroll.Height = ((ToolsFrm.Height - 350) \ ToolsBar.ButtonHeight) * ToolsBar.ButtonHeight
      ToolsScroll.Max = 13 - (ToolsScroll.Height \ ToolsBar.ButtonHeight)
      ToolsPic1.Height = ToolsScroll.Height
      'AnimFrm.Move LayerFrm.Left, LayerFrm.Top + LayerFrm.Height + 60, AnimFrm.Width, Map2.Height - 1880
      Call ToolsScroll_Change
      Call TileScroll_Change
  
  
    Case Win_TextEdit
      EditINI.Move EditINI.Left, EditINI.Top, MainFrame.Width - 250, MainFrame.Height - 350
      
      
    Case Win_DAN
      DAN_Info.Move DAN_Info.Left, DAN_Info.Top, MainFrame.Width - 220, MainFrame.Height - 350
      
      
    Case Win_Logo
      DarkainLogo.Move FileList.Width + 45, 390, MainFrame.Width, MainFrame.Height
  End Select
  
  Loading = False
  Map2_Paint
  TileScroll_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim TheReturn As Integer
  Dim Ret As Integer
  
  If FileHasChanged Then
    TheReturn = MsgBox("Save before exiting?", 35, "Darkain's Alhpa Editor")
    
    Select Case TheReturn
      Case 6
        Unhook
        Call MnuFileSave_Click
        Ret = mciSendString("close DarkainMidi", 0&, 0, 0)
      Case 7
        Unhook
        Ret = mciSendString("close DarkainMidi", 0&, 0, 0)
      Case Else
        Cancel = True
    End Select
    
  Else
    Unhook
  End If
End Sub

Private Sub LbLayer_Click(Index As Integer)
  Dim i As Integer
  For i = 0 To 3
    LbLayer(i).BackColor = &H80000004
    LbLayer(i).ForeColor = &H80000007
  Next i
  LbLayer(Index).BackColor = &H8000000D
  LbLayer(Index).ForeColor = &H8000000E
  CurrentLayer = Index
  If MnuObst.Checked Then Map2_Paint
End Sub

Private Sub Map2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  Select Case KeyCode
    Case vbKeyLeft
      MapScrollH.Value = MapScrollH.Value - 1
    Case vbKeyRight
      MapScrollH.Value = MapScrollH.Value + 1
    Case vbKeyUp
      MapScrollV.Value = MapScrollV.Value - 1
    Case vbKeyDown
      MapScrollV.Value = MapScrollV.Value + 1
  End Select
End Sub

Private Sub Map2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim XX As Single
  Dim YY As Single
  
  If CurrentLayer = 255 Then Exit Sub
  If X < 1 Or Y < 1 Or X >= Map2.ScaleWidth Or Y >= Map2.ScaleHeight Then Exit Sub
  If Button = 2 Then
    PopupMenu MnuTile
  
  Else
    Select Case DrawTool
      Case Tools_Draw
        Call DrawTile(X, Y)
        MouseDown = 1
      
      Case Tools_Select
        Loading = True
        TileSelectLine(4).Visible = False
        TileSelectLine(5).Visible = False
        TileSelectLine(6).Visible = False
        TileSelectLine(7).Visible = False
        Orig_X2 = (X \ 32) * 32 + 1
        Orig_Y2 = (Y \ 32) * 32 + 1
        Orig_Z2 = 1
  
        XX = (X \ 32 + 1) * 32 - 2
        YY = (Y \ 32 + 1) * 32 - 2
    
        If XX < 2 Then XX = 2
        If XX > Map2.Width - 2 Then XX = Map2.Width - 2
  
        If YY < 2 Then YY = 2
        If YY > Map2.Height - 2 Then YY = Map2.Height - 2
  
        TileSelectLine(4).X1 = Orig_X2
        TileSelectLine(4).Y1 = YY
        TileSelectLine(4).X2 = XX
        TileSelectLine(4).Y2 = YY
  
        TileSelectLine(5).X1 = XX
        TileSelectLine(5).Y1 = Orig_Y2
        TileSelectLine(5).X2 = XX
        TileSelectLine(5).Y2 = YY
  
        TileSelectLine(6).X1 = Orig_X2
        TileSelectLine(6).Y1 = Orig_Y2
        TileSelectLine(6).X2 = Orig_X2
        TileSelectLine(6).Y2 = YY
  
        TileSelectLine(7).X1 = Orig_X2
        TileSelectLine(7).Y1 = Orig_Y2
        TileSelectLine(7).X2 = XX
        TileSelectLine(7).Y2 = Orig_Y2

        TileSelectLine(4).Visible = True
        TileSelectLine(5).Visible = True
        TileSelectLine(6).Visible = True
        TileSelectLine(7).Visible = True
        Loading = False
    End Select
  End If
End Sub

Private Sub Map2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Bar.Panels(1).Text = "X " & (X \ 32) + MapScrollH.Value & "  Y " & (Y \ 32) + MapScrollV.Value
  
  Select Case DrawTool
    Case Tools_Draw
      If MouseDown = 0 Or X < 1 Or Y < 1 Or X >= Map2.ScaleWidth - MapScrollV.Width Or Y >= Map2.ScaleHeight - MapScrollH.Height Then Exit Sub
      Call DrawTile(X, Y)
    
    Case Tools_Select
      If Orig_Z2 = 1 Then
  
        Loading = True
        TileSelectLine(4).Visible = False
        TileSelectLine(5).Visible = False
        TileSelectLine(6).Visible = False
        TileSelectLine(7).Visible = False
        
        Dim XX As Single
        Dim YY As Single
  
        If X < Orig_X2 Then
          XX = (Orig_X2 \ 32 + 1) * 32 - 2
        ElseIf X > Map2.ScaleWidth - 1 Then
          XX = (Map2.ScaleWidth \ 32) * 32 - 2
        Else
          XX = (X \ 32 + 1) * 32 - 2
        End If
    
        If Y < Orig_Y2 Then
          YY = (Orig_Y2 \ 32 + 1) * 32 - 2
        ElseIf Y > Map2.ScaleHeight - 1 Then
          YY = (Map2.ScaleHeight \ 32) * 32 - 2
        Else
          YY = (Y \ 32 + 1) * 32 - 2
        End If
    
        If XX < 2 Then XX = 2
        If XX > Map2.Width - 2 Then XX = Map2.Width - 2
      
        If YY < 2 Then YY = 2
        If YY > Map2.Height - 2 Then YY = Map2.Height - 2
    
        TileSelectLine(4).X1 = Orig_X2
        TileSelectLine(4).Y1 = YY
        TileSelectLine(4).X2 = XX
        TileSelectLine(4).Y2 = YY
  
        TileSelectLine(5).X1 = XX
        TileSelectLine(5).Y1 = Orig_Y2
        TileSelectLine(5).X2 = XX
        TileSelectLine(5).Y2 = YY
  
        TileSelectLine(6).X1 = Orig_X2
        TileSelectLine(6).Y1 = Orig_Y2
        TileSelectLine(6).X2 = Orig_X2
        TileSelectLine(6).Y2 = YY
  
        TileSelectLine(7).X1 = Orig_X2
        TileSelectLine(7).Y1 = Orig_Y2
        TileSelectLine(7).X2 = XX
        TileSelectLine(7).Y2 = Orig_Y2
  
        TileSelectLine(4).Visible = True
        TileSelectLine(5).Visible = True
        TileSelectLine(6).Visible = True
        TileSelectLine(7).Visible = True
        Loading = False
      End If

  End Select
End Sub

Private Sub Map2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseDown = 0
  
  Select Case DrawTool
    Case Tools_Select
      Orig_Z2 = 0
    
      Dim XX As Single
      Dim YY As Single
    
      Dim X_Dist As Integer
      Dim Y_Dist As Integer
    
      Dim X_Start As Integer
      Dim Y_Start As Integer
    
      Dim i1 As Integer
      Dim i2 As Integer
  
      If X < Orig_X2 Then
        XX = (Orig_X2 \ 32 + 1) * 32
      ElseIf X > Map2.ScaleWidth - 1 Then
        XX = (Map2.ScaleWidth \ 32) * 32
      Else
        XX = (X \ 32 + 1) * 32
      End If
    
      If Y < Orig_Y2 Then
        YY = (Orig_Y2 \ 32 + 1) * 32
      ElseIf Y > Map2.ScaleHeight - 1 Then
        YY = (Map2.ScaleHeight \ 32) * 32
      Else
        YY = (Y \ 32 + 1) * 32
      End If
    
      X_Start = (Orig_X2 - 1) \ 32
      Y_Start = (Orig_Y2 - 1) \ 32
      X_Dist = (XX \ 32) - 1 - X_Start
      Y_Dist = (YY \ 32) - 1 - Y_Start
  End Select
End Sub

Sub DrawTile(X As Single, Y As Single)
  Dim DestDC As Long
  Dim DestBMP As Long
  Dim DestX As Long
  
  Dim TmpX     As Integer, TmpY    As Integer
  Dim i1       As Integer, i2      As Integer
  Dim i3       As Byte, FirstDraw  As Boolean
  Dim Tmp1     As Integer, Tmp2    As Integer
  Dim Tmp3     As Integer, Tmp4    As Integer
  Dim CurTile  As Integer
  Dim CurShade As Byte
  
  FileHasChanged = True
  
  DestDC = CreateCompatibleDC(Map2.hDc)
  DestBMP = CreateCompatibleBitmap(Map2.hDc, (SelTiles.Width + 1) * 32, (SelTiles.Height + 1) * 32)
  DestX = SelectObject(DestDC, DestBMP)

  
  TmpX = X \ 32
  TmpY = Y \ 32
  Tmp3 = TmpX * 32
  Tmp4 = TmpY * 32
  
  If TmpX + MapScrollH.Value >= MapInfo.MapData.Right Then
    MapInfo.MapData.Right = TmpX + MapScrollH.Value + 1
  End If
  If TmpX + MapScrollH.Value <= MapInfo.MapData.Left Then
    MapInfo.MapData.Left = TmpX + MapScrollH.Value
  End If
  If TmpY + MapScrollV.Value >= MapInfo.MapData.Bottom Then
    MapInfo.MapData.Bottom = TmpY + MapScrollV.Value + 1
  End If
  If TmpY + MapScrollV.Value <= MapInfo.MapData.Top Then
    MapInfo.MapData.Top = TmpY + MapScrollV.Value
  End If


  If MnuObst.Checked Then
    MapInfo.MapData.Walk(TmpX + MapScrollH.Value + 1, TmpY + MapScrollV.Value + 1, CurrentLayer) = CurrentShade
  Else
    'MapInfo.MapData.Tile(TmpX + MapScrollH.Value + 1, TmpY + MapScrollV.Value + 1, CurrentLayer) = CurrentTile
    For i1 = 0 To SelTiles.Width
      For i2 = 0 To SelTiles.Height
        MapInfo.MapData.Tile(TmpX + MapScrollH.Value + 1 + i1, TmpY + MapScrollV.Value + 1 + i2, CurrentLayer) = SelTiles.Tiles(i1, i2)
      Next i2
    Next i1
  End If

  If MnuTiles.Checked Then
    For i1 = 0 To SelTiles.Width
      For i2 = 0 To SelTiles.Height
    
        For i3 = 0 To 3
          If MnuLayers(i3).Checked Then
            CurTile = MapInfo.MapData.Tile(TmpX + MapScrollH.Value + 1 + i1, TmpY + MapScrollV.Value + 1 + i2, i3)
            If CurTile > 1023 Then
              Tmp1 = AnimX
              Tmp2 = (CurTile - 1024) * 32
              If i3 = 0 Then
                Call BitBlt(DestDC, i1 * 32, i2 * 32, 32, 32, PicBlank.hDc, 0, 0, SRCCOPY)
              End If
'              Call BitBlt(DestDC, i1 * 32, i2 * 32, 32, 32, AnimMask.hDc, Tmp1, Tmp2, SRCAND)
'              Call BitBlt(DestDC, i1 * 32, i2 * 32, 32, 32, AnimSet.hDc, Tmp1, Tmp2, SRCPAINT)
            Else
              Tmp1 = (CurTile Mod (TileSet2.ScaleWidth \ 32)) * 32     ' *0.03125
              Tmp2 = (CurTile \ (TileSet2.ScaleWidth \ 32)) * 32
              If i3 = 0 Then
                Call BitBlt(DestDC, i1 * 32, i2 * 32, 32, 32, PicBlank.hDc, 0, 0, SRCCOPY)
              End If
              Call BitBlt(DestDC, i1 * 32, i2 * 32, 32, 32, TileMask.hDc, Tmp1, Tmp2, SRCAND)
              Call BitBlt(DestDC, i1 * 32, i2 * 32, 32, 32, TileSet2.hDc, Tmp1, Tmp2, SRCPAINT)
            End If
          End If
        Next i3
    
      Next i2
    Next i1
  End If

  If MnuObst.Checked Then
    For i1 = 0 To SelTiles.Width
      For i2 = 0 To SelTiles.Height
        CurShade = MapInfo.MapData.Walk(TmpX + MapScrollH.Value + 1 + i1, TmpY + MapScrollV.Value + 1 + i2, CurrentLayer)
        Tmp1 = (CurShade Mod (Shade.ScaleWidth \ 32)) * 32
        Tmp2 = (CurShade \ (Shade.ScaleWidth \ 32)) * 32
        Call BitBlt(DestDC, i1 * 32, i2 * 32, 32, 32, Shade.hDc, Tmp1, Tmp2, SRCAND)
      Next i2
    Next i1
  End If

  If MnuGrid.Checked Then
    For i1 = 0 To SelTiles.Height
      If TmpY + MapScrollV.Value + 1 + i1 = 0 Then
        Call Lines(DestDC, 0, (i1 + 1) * 32 - 1, (SelTiles.Width + 1) * 32, (i1 + 1) * 32 - 1, RGB(255, 0, 0))
      ElseIf (TmpY + MapScrollV.Value + 1 + i1) Mod 25 = 0 Then
        Call Lines(DestDC, 0, (i1 + 1) * 32 - 1, (SelTiles.Width + 1) * 32, (i1 + 1) * 32 - 1, RGB(0, 255, 0))
      ElseIf (TmpY + MapScrollV.Value + 1 + i1) Mod 5 = 0 Then
        Call Lines(DestDC, 0, (i1 + 1) * 32 - 1, (SelTiles.Width + 1) * 32, (i1 + 1) * 32 - 1, RGB(0, 0, 255))
      Else
        Call Lines(DestDC, 0, (i1 + 1) * 32 - 1, (SelTiles.Width + 1) * 32, (i1 + 1) * 32 - 1, RGB(0, 0, 0))
      End If
    Next i1
    
    For i1 = 0 To SelTiles.Width
      If TmpX + MapScrollH.Value + 1 + i1 = 0 Then
        Call Lines(DestDC, (i1 + 1) * 32 - 1, 0, (i1 + 1) * 32 - 1, (SelTiles.Height + 1) * 32, RGB(255, 0, 0))
      ElseIf (TmpX + MapScrollH.Value + 1 + i1) Mod 25 = 0 Then
        Call Lines(DestDC, (i1 + 1) * 32 - 1, 0, (i1 + 1) * 32 - 1, (SelTiles.Height + 1) * 32, RGB(0, 255, 0))
      ElseIf (TmpX + MapScrollH.Value + 1 + i1) Mod 5 = 0 Then
        Call Lines(DestDC, (i1 + 1) * 32 - 1, 0, (i1 + 1) * 32 - 1, (SelTiles.Height + 1) * 32, RGB(0, 0, 255))
      Else
        Call Lines(DestDC, (i1 + 1) * 32 - 1, 0, (i1 + 1) * 32 - 1, (SelTiles.Height + 1) * 32, RGB(0, 0, 0))
      End If
    Next i1
  End If
  
  Call BitBlt(Map2.hDc, Tmp3, Tmp4, (SelTiles.Width + 1) * 32, (SelTiles.Height + 1) * 32, DestDC, 0, 0, SRCCOPY)
  
  DestX = SelectObject(DestDC, DestBMP)
  DestX = DeleteDC(DestDC)
  DestX = DeleteObject(DestBMP)
End Sub

Private Sub Map2_Paint()
  On Error GoTo ErrHan
  
  If Loading Then Exit Sub
  If ActiveWin <> Win_MapEdit Then Exit Sub
  
  Dim i1 As Integer, i2 As Integer, i3 As Byte
  Dim Tmp1 As Integer, Tmp2 As Integer
  Dim Tmp3 As Integer, Tmp4 As Integer
  Dim FirstDraw As Boolean
  Dim CurTile As Integer
  Dim CurShade As Byte
  
  Dim DestDC As Long
  Dim DestBMP As Long
  Dim DestX As Long

  DestDC = CreateCompatibleDC(Map2.hDc)
  DestBMP = CreateCompatibleBitmap(Map2.hDc, Map2.ScaleWidth, Map2.ScaleHeight)
  DestX = SelectObject(DestDC, DestBMP)

  For i1 = MapScrollH.Value To MapScrollH.Value + (Map2.ScaleWidth \ 32)  '11
    For i2 = MapScrollV.Value To MapScrollV.Value + (Map2.ScaleHeight \ 32) '10
      Tmp3 = (i1 - 1 - MapScrollH.Value) * 32
      Tmp4 = (i2 - 1 - MapScrollV.Value) * 32
      If MnuTiles.Checked Then
        For i3 = 0 To 3
          If MnuLayers(i3).Checked Then
            CurTile = MapInfo.MapData.Tile(i1, i2, i3)
            
            If CurTile > 1023 Then
              Tmp1 = AnimX
              Tmp2 = (CurTile - 1024) * 32
              If Not FirstDraw Then
                Call BitBlt(DestDC, Tmp3, Tmp4, 32, 32, PicBlank.hDc, 0, 0, SRCCOPY)
                FirstDraw = True
              End If
              If CurTile > 0 Then
'                Call BitBlt(DestDC, Tmp3, Tmp4, 32, 32, AnimMask.hDc, Tmp1, Tmp2, SRCAND)
'                Call BitBlt(DestDC, Tmp3, Tmp4, 32, 32, AnimSet.hDc, Tmp1, Tmp2, SRCPAINT)
              End If
            
            Else
              Tmp1 = (CurTile Mod (TileSet2.ScaleWidth / 32)) * 32
              Tmp2 = Int(CurTile / (TileSet2.ScaleWidth / 32)) * 32
              If Not FirstDraw Then
                Call BitBlt(DestDC, Tmp3, Tmp4, 32, 32, PicBlank.hDc, 0, 0, SRCCOPY)
                FirstDraw = True
              End If
              If CurTile > 0 Then
                Call BitBlt(DestDC, Tmp3, Tmp4, 32, 32, TileMask.hDc, Tmp1, Tmp2, SRCAND)
                Call BitBlt(DestDC, Tmp3, Tmp4, 32, 32, TileSet2.hDc, Tmp1, Tmp2, SRCPAINT)
              End If
            End If
          End If
        Next i3
      End If
      If MnuObst.Checked Then
        CurShade = MapInfo.MapData.Walk(i1, i2, CurrentLayer)
        Tmp1 = (CurShade Mod (Shade.ScaleWidth / 32)) * 32
        Tmp2 = Int(CurShade / (Shade.ScaleWidth / 32)) * 32
        Call BitBlt(DestDC, Tmp3, Tmp4, 32, 32, Shade.hDc, Tmp1, Tmp2, SRCAND)
      End If
      FirstDraw = False
    Next i2
  Next i1

  If MnuGrid.Checked Then
    Tmp1 = (Map2.ScaleWidth \ 32) + 1
    Tmp2 = (Map2.ScaleHeight \ 32) + 1
    For i1 = 1 To Tmp1
      If MapScrollH.Value + i1 = 0 Then
        Call Lines(DestDC, i1 * 32 - 1, 0, i1 * 32 - 1, Tmp2 * 32, RGB(255, 0, 0))
      ElseIf (MapScrollH.Value + i1) Mod 25 = 0 Then
        Call Lines(DestDC, i1 * 32 - 1, 0, i1 * 32 - 1, Tmp2 * 32, RGB(0, 255, 0))
      ElseIf (MapScrollH.Value + i1) Mod 5 = 0 Then
        Call Lines(DestDC, i1 * 32 - 1, 0, i1 * 32 - 1, Tmp2 * 32, RGB(0, 0, 255))
      Else
        Call Lines(DestDC, i1 * 32 - 1, 0, i1 * 32 - 1, Tmp2 * 32, RGB(0, 0, 0))
      End If
    Next i1
    For i1 = 1 To Tmp2
      If MapScrollV.Value + i1 = 0 Then
        Call Lines(DestDC, 0, i1 * 32 - 1, Tmp1 * 32, i1 * 32 - 1, RGB(255, 0, 0))
      ElseIf (MapScrollV.Value + i1) Mod 25 = 0 Then
        Call Lines(DestDC, 0, i1 * 32 - 1, Tmp1 * 32, i1 * 32 - 1, RGB(0, 255, 0))
      ElseIf (MapScrollV.Value + i1) Mod 5 = 0 Then
        Call Lines(DestDC, 0, i1 * 32 - 1, Tmp1 * 32, i1 * 32 - 1, RGB(0, 0, 255))
      Else
        Call Lines(DestDC, 0, i1 * 32 - 1, Tmp1 * 32, i1 * 32 - 1, RGB(0, 0, 0))
      End If
    Next i1
  End If

  Call BitBlt(Map2.hDc, 0, 0, Map2.ScaleWidth, Map2.ScaleHeight, DestDC, 0, 0, SRCCOPY)
  
  DestX = SelectObject(DestDC, DestBMP)
  DestX = DeleteDC(DestDC)
  DestX = DeleteObject(DestBMP)
Exit Sub

ErrHan:
  MsgBox "Map2_Paint: " & Err.Description
End Sub

Private Sub MapScrollH_Change()
  Map2_Paint
End Sub

Private Sub MapScrollH_GotFocus()
  Map2.SetFocus
End Sub

Private Sub MapScrollH_Scroll()
  Map2_Paint
End Sub

Private Sub MapScrollV_Change()
  Map2_Paint
End Sub

Private Sub MapScrollV_GotFocus()
  Map2.SetFocus
End Sub

Private Sub MapScrollV_Scroll()
  Map2_Paint
End Sub


Private Sub MnuAnim_Click()
  Dim i As Integer
  MnuAnim.Checked = Not MnuAnim.Checked
  AnimTimer.Enabled = MnuAnim.Checked
  For i = 1 To TBar.Buttons.Count
    If TBar.Buttons(i).Key = "Anim" Then
      TBar.Buttons(i).Value = -MnuAnim.Checked
    End If
  Next i
End Sub

Public Sub MnuDAN_Open_Click()
  DAN_Info.FileName = GameMain.Paths.Root & NodeString
  Bar.Panels(2).Text = NodeString
  MakeVis (Win_DAN)
End Sub

Public Sub MnuDAN_Text_Click()
  Load_DAI (GameMain.Paths.Root & NodeString)
End Sub

Private Sub MnuFileExit_Click()
  Unload Me
End Sub

Private Sub MnuFileNew_Click()
  Call New_File
End Sub

Private Sub MnuFileOpen_Click()
'  Dlg.CancelError = True
'  On Error GoTo ErrHan
'  SetRedraw (True)
'  Dlg.Filter = "Map Files (*.Map)|*.Map|All Files (*.*)|*.*"
'  Dlg.FilterIndex = 1
'  Dlg.DialogTitle = "Open Map File"
'  Dlg.Flags = &H1000
'  Dlg.ShowOpen
'  Call LoadMap(Dlg.FileName)
'
'ErrHan:
'  SetRedraw (False)
'  Err.Clear

'LoadPictureMap ("")

MsgBox "Temporarily disabled"
End Sub

Private Sub MnuFileSave_Click()
  If Len(OpenedFile) > 0 Then
    Select Case ActiveWin
    
      Case Win_MapEdit
        Save_DAM OpenedFile
      
      Case Win_DAN
        Save_DAN OpenedFile
    
      Case Win_TextEdit
        Save_DAI OpenedFile
      
    End Select
  Else
    Call MnuFileSaveAs_Click
  End If
End Sub

Private Sub MnuFileSaveAs_Click()
  Dlg.CancelError = True
  On Error GoTo ErrHan
  SetRedraw (True)
    
  Select Case ActiveWin
    
    Case Win_MapEdit
      Dlg.InitDir = GameMain.Paths.Maps
      Dlg.Filter = "Map Files (*.DAM)|*.DAM|All Files (*.*)|*.*"
      Dlg.FilterIndex = 1
      Dlg.DialogTitle = "Save Map File As"
      Dlg.Flags = &H802
      Dlg.ShowSave
      Save_DAM Dlg.FileName
     
    Case Win_DAN
      Dlg.InitDir = GameMain.Paths.Maps
      Dlg.Filter = "Map Info Files (*.DAN)|*.DAN|All Files (*.*)|*.*"
      Dlg.FilterIndex = 1
      Dlg.DialogTitle = "Save Map Info File As"
      Dlg.Flags = &H802
      Dlg.ShowSave
      Save_DAN Dlg.FileName
    
    Case Win_TextEdit
      Dlg.InitDir = GameMain.Paths.System
      Dlg.Filter = "Information Files (*.DAI)|*.DAI|Map Info Files (*.DAN)|*.DAN|Text (*.TXT)|*.TXT|HTML (*.HTML)|*.HTML|All Files (*.*)|*.*"
      Dlg.FilterIndex = 1
      Dlg.DialogTitle = "Save Text Information File As"
      Dlg.Flags = &H802
      Dlg.ShowSave
      Save_DAI Dlg.FileName
      
  End Select
  
  SetRedraw (False)
Exit Sub
  
  
ErrHan:
  SetRedraw (False)
  Err.Clear
End Sub

Private Sub MnuGrid_Click()
  Dim i As Integer
  MnuGrid.Checked = Not MnuGrid.Checked
  Map2_Paint
End Sub

Private Sub MnuLayers_Click(Index As Integer)
  Dim i As Integer
  MnuLayers(Index).Checked = Not MnuLayers(Index).Checked
  If Not MnuLayers(0).Checked And Not MnuLayers(1).Checked And Not MnuLayers(2).Checked And Not MnuLayers(3).Checked Then
    MnuLayers(Index).Checked = True
    For i = 1 To TBar.Buttons.Count
      If TBar.Buttons(i).Key = "Layer1" Then
        TBar.Buttons(i + Index).Value = -MnuLayers(Index).Checked
        Exit For
      End If
    Next i
    Exit Sub
  End If
  LbLayer(Index).Enabled = MnuLayers(Index).Checked
  If MnuLayers(Index).Checked = False And CurrentLayer = Index Then
    LbLayer(Index).BackColor = &H80000004
    LbLayer(Index).ForeColor = &H80000007
    For i = 0 To 3
      If MnuLayers(i).Checked Then
        CurrentLayer = i
        LbLayer(i).BackColor = &H8000000D
        LbLayer(i).ForeColor = &H8000000E
        Exit For
      End If
    Next i
  End If
  For i = 1 To TBar.Buttons.Count
    If TBar.Buttons(i).Key = "Layer1" Then
      TBar.Buttons(i + Index).Value = -MnuLayers(Index).Checked
      Exit For
    End If
  Next i
  Map2.Refresh
  Map2_Paint
End Sub

Private Sub MnuObst_Click()
  Dim i As Integer
  MnuObst.Checked = Not MnuObst.Checked
  TileScroll_Change
  If MnuObst.Checked Then
    TileFrm.Caption = "Obstruction"
    TileScroll.Enabled = False
'    AnimView.Visible = False
  Else
    TileFrm.Caption = "Tiles"
    TileScroll.Enabled = True
'    AnimView.Visible = True
  End If
  For i = 1 To TBar.Buttons.Count
    If TBar.Buttons(i).Key = "Obst" Then
      TBar.Buttons(i).Value = -MnuObst.Checked
    End If
  Next i
  Map2_Paint
End Sub

Private Sub MnuTiles_Click()
  Dim i As Integer
  MnuTiles.Checked = Not MnuTiles.Checked
  TileScroll_Change
  For i = 1 To TBar.Buttons.Count
    If TBar.Buttons(i).Key = "Tiles" Then
      TBar.Buttons(i).Value = -MnuTiles.Checked
    End If
  Next i
  Map2_Paint
End Sub

Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "New"
      Call MnuFileNew_Click
    Case "Save"
      Call MnuFileSave_Click
    Case "Open"
      Call MnuFileOpen_Click
    Case "Layer1"
      Call MnuLayers_Click(0)
    Case "Layer2"
      Call MnuLayers_Click(1)
    Case "Layer3"
      Call MnuLayers_Click(2)
    Case "Layer4"
      Call MnuLayers_Click(3)
    Case "Grid"
      Call MnuGrid_Click
    Case "Obst"
      Call MnuObst_Click
    Case "Tiles"
      Call MnuTiles_Click
    Case "Anim"
      Call MnuAnim_Click
  End Select
End Sub

Private Sub TileCat_Click()
  If Loading Then Exit Sub

 TileScroll_Change
End Sub

Private Sub TileScroll_Change()
On Error GoTo ErrHan
  
  If Loading Then Exit Sub
  
  Dim DestDC As Long
  Dim DestBMP As Long
  Dim DestX As Long
  
  Dim Tmp As Integer, Tmp1 As Integer
  Dim i As Integer
  Dim CurTile As Integer
  
  DestDC = CreateCompatibleDC(TileSet2.hDc)
  DestBMP = CreateCompatibleBitmap(TileSet2.hDc, TileSet2.ScaleWidth, TileSet2.ScaleHeight)
  DestX = SelectObject(DestDC, DestBMP)
  
  If MnuObst.Checked Then
    Tmp1 = TileView.Height \ 480 ' * 4
    TileView.Cls
    'tmp = TileScroll Mod 2
    For i = 0 To 20
      Call BitBlt(DestDC, 0, i * 32, 128, 32, Shade.hDc, (i Mod 2) * 128, (i \ 2) * 32, SRCCOPY)
    Next i
    Call BitBlt(TileView.hDc, 0, 0, TileView.ScaleWidth, TileView.ScaleHeight, DestDC, 0, 0, SRCCOPY)
  
  Else
    TileView.Cls
    
    For i = 0 To (TileView.ScaleWidth \ 32) * (TileView.ScaleHeight \ 32)
      CurTile = TileList(TileCat.ListIndex, i + (TileScroll.Value * 4))
      Call BitBlt(DestDC, (i Mod 4) * 32, (i \ 4) * 32, 32, 32, TileMask.hDc, (CurTile Mod 8) * 32, (CurTile \ 8) * 32, SRCCOPY)
      Call BitBlt(DestDC, (i Mod 4) * 32, (i \ 4) * 32, 32, 32, TileSet2.hDc, (CurTile Mod 8) * 32, (CurTile \ 8) * 32, SRCPAINT)
    Next i
    
    Call BitBlt(TileView.hDc, 0, 0, TileView.ScaleWidth, TileView.ScaleHeight, DestDC, 0, 0, SRCCOPY)
  End If
  
  DestX = SelectObject(DestDC, DestBMP)
  DestX = DeleteDC(DestDC)
  DestX = DeleteObject(DestBMP)
Exit Sub

ErrHan:
End Sub

Private Sub TileScroll_GotFocus()
  TileView.SetFocus
End Sub

Private Sub TileScroll_Scroll()
  Call TileScroll_Change
End Sub

Private Sub TileView_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  Select Case KeyCode
    Case vbKeyUp
      TileScroll.Value = TileScroll.Value - 1
    Case vbKeyDown
      TileScroll.Value = TileScroll.Value + 1
  End Select
End Sub

Private Sub TileView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Tmp As Integer
  Dim XX As Single
  Dim YY As Single

  If MnuObst.Checked Then
    Tmp = (X \ 32) + ((Y \ 64) * 2) * 4
    If (Y \ 32) Mod 2 = 1 Then
      Tmp = Tmp + 4
    End If
    Bar.Panels(2).Text = "Left Button Shade: " & Tmp
    CurrentShade = Tmp
  Else
'    Tmp = (X \ 32) + ((Y \ 64) * 2 + TileScroll.Value) * 4
'    If (Y \ 32) Mod 2 = 1 Then
'      Tmp = Tmp + 4
'    End If
'    Bar.Panels(2).Text = "Left Button Tile: " & Tmp
'    CurrentTile = Tmp
  End If
  
  
  If Button = 1 Then
    Orig_X = (X \ 32) * 32 + 1
    Orig_Y = (Y \ 32) * 32 + 1
    Orig_Z = 1
  
    XX = (X \ 32 + 1) * 32 - 2
    YY = (Y \ 32 + 1) * 32 - 2
    
    If XX < 2 Then XX = 2
    If XX > TileView.Width - 2 Then XX = TileView.Width - 2
  
    If YY < 2 Then YY = 2
    If YY > TileView.Height - 2 Then YY = TileView.Height - 2
  
    TileSelectLine(0).X1 = Orig_X
    TileSelectLine(0).Y1 = YY
    TileSelectLine(0).X2 = XX
    TileSelectLine(0).Y2 = YY
  
    TileSelectLine(1).X1 = XX
    TileSelectLine(1).Y1 = Orig_Y
    TileSelectLine(1).X2 = XX
    TileSelectLine(1).Y2 = YY
  
    TileSelectLine(2).X1 = Orig_X
    TileSelectLine(2).Y1 = Orig_Y
    TileSelectLine(2).X2 = Orig_X
    TileSelectLine(2).Y2 = YY
  
    TileSelectLine(3).X1 = Orig_X
    TileSelectLine(3).Y1 = Orig_Y
    TileSelectLine(3).X2 = XX
    TileSelectLine(3).Y2 = Orig_Y
  End If
End Sub

Private Sub TileView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Orig_Z = 1 Then
  
    Dim XX As Single
    Dim YY As Single
  
    If X < Orig_X Then
      XX = (Orig_X \ 32 + 1) * 32 - 2
    ElseIf X > TileView.ScaleWidth - 1 Then
      XX = (TileView.ScaleWidth \ 32) * 32 - 2
    Else
      XX = (X \ 32 + 1) * 32 - 2
    End If
    
    If Y < Orig_Y Then
      YY = (Orig_Y \ 32 + 1) * 32 - 2
    ElseIf Y > TileView.ScaleHeight - 1 Then
      YY = (TileView.ScaleHeight \ 32) * 32 - 2
    Else
      YY = (Y \ 32 + 1) * 32 - 2
    End If
    
    If XX < 2 Then XX = 2
    If XX > TileView.Width - 2 Then XX = TileView.Width - 2
  
    If YY < 2 Then YY = 2
    If YY > TileView.Height - 2 Then YY = TileView.Height - 2
  
    TileSelectLine(0).X1 = Orig_X
    TileSelectLine(0).Y1 = YY
    TileSelectLine(0).X2 = XX
    TileSelectLine(0).Y2 = YY
  
    TileSelectLine(1).X1 = XX
    TileSelectLine(1).Y1 = Orig_Y
    TileSelectLine(1).X2 = XX
    TileSelectLine(1).Y2 = YY
  
    TileSelectLine(2).X1 = Orig_X
    TileSelectLine(2).Y1 = Orig_Y
    TileSelectLine(2).X2 = Orig_X
    TileSelectLine(2).Y2 = YY
  
    TileSelectLine(3).X1 = Orig_X
    TileSelectLine(3).Y1 = Orig_Y
    TileSelectLine(3).X2 = XX
    TileSelectLine(3).Y2 = Orig_Y
  End If
End Sub

Private Sub TileView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 1 And Orig_Z = 1 Then
    Orig_Z = 0
    
    Dim XX As Single
    Dim YY As Single
    
    Dim X_Dist As Integer
    Dim Y_Dist As Integer
    
    Dim X_Start As Integer
    Dim Y_Start As Integer
    
    Dim i1 As Integer
    Dim i2 As Integer
  
    If X < Orig_X Then
      XX = (Orig_X \ 32 + 1) * 32
    ElseIf X > TileView.ScaleWidth - 1 Then
      XX = (TileView.ScaleWidth \ 32) * 32
    Else
      XX = (X \ 32 + 1) * 32
    End If
    
    If Y < Orig_Y Then
      YY = (Orig_Y \ 32 + 1) * 32
    ElseIf Y > TileView.ScaleHeight - 1 Then
      YY = (TileView.ScaleHeight \ 32) * 32
    Else
      YY = (Y \ 32 + 1) * 32
    End If
    
    X_Start = (Orig_X - 1) \ 32
    Y_Start = (Orig_Y - 1) \ 32
    X_Dist = (XX \ 32) - 1 - X_Start
    Y_Dist = (YY \ 32) - 1 - Y_Start
    ReDim SelTiles.Tiles(X_Dist, Y_Dist)
    For i1 = 0 To X_Dist
      For i2 = 0 To Y_Dist
        SelTiles.Tiles(i1, i2) = TileList(TileCat.ListIndex, (X_Start + i1) + ((Y_Start + i2) * 4) + (TileScroll.Value * 4))
      Next i2
    Next i1
    
    SelTiles.Width = X_Dist
    SelTiles.Height = Y_Dist
  
  End If

End Sub

Private Sub TileView_Paint()
  Call TileScroll_Change
End Sub

Private Sub ToolsBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim i As Integer
  For i = 1 To ToolsBar.Buttons.Count
    ToolsBar.Buttons(i).Value = tbrUnpressed
  Next i
  
  DrawTool = Button.Index
  
  Button.Value = tbrPressed
  
  ToolsBar.Refresh
End Sub

Private Sub ToolsScroll_Change()
  Dim i As Integer
  Dim StartButton As Integer
  Dim EndButton As Integer
  

  StartButton = ToolsScroll.Value - 1
  EndButton = StartButton + ((ToolsScroll.Height - 2) \ ToolsBar.ButtonHeight) + 2
  
  If Loading Then
    StartButton = 0
    EndButton = 14
  End If
  
  For i = 1 To 13
    If i > StartButton And i < EndButton Then
      ToolsBar.Buttons(i).Visible = True
    Else
      ToolsBar.Buttons(i).Visible = False
    End If
  Next i
  
  On Local Error Resume Next
  ToolsScroll.LargeChange = EndButton - StartButton - 2
  If Err.Number Then
    ToolsScroll.LargeChange = 13
  End If
End Sub

Private Sub ToolsScroll_GotFocus()
  Map2.SetFocus
End Sub

Private Sub ToolsScroll_Scroll()
  Call ToolsScroll_Change
End Sub
