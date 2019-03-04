Attribute VB_Name = "OthersDK"
Option Explicit
Private Type MapHeaderType
  Map            As String
  Title          As String
  CodeScript     As String
  FontSet        As String
  TileSet        As String
  Pallet         As String
End Type

Private Type MapFileType
  Top            As Integer
  Bottom         As Integer
  Left           As Integer
  Right          As Integer
  Layers         As Byte
  Flag()         As Byte                   'Flag Map
  Walk()         As Byte                   'Walkable locations
  Tile()         As Integer                'Tile locations
  Visible        As Boolean
  
  'remove
  OffsetX        As Single
  OffsetY        As Single
  OSX            As Single
  OSY            As Single
  
  rTiles         As RECT
  ddsTiles       As DirectDrawSurface7     'Tile Surface
  ddsdTiles      As DDSURFACEDESC2         'Tile Info
  ddsTiles2      As DirectDrawSurface7     'Tile2 Surface
  ddsdTiles2     As DDSURFACEDESC2         'Tile2 Info
  AnimFrame      As Integer                'Current Animate Frame
  AnimFrame2     As Integer
End Type

Private Type CharacterType
  DefX           As Long
  DefY           As Long
  DefLayer       As Byte
  DefDir         As Byte
  
  X              As Single
  Y              As Single
  OX             As Single
  OY             As Single
  WX             As Single
  WY             As Integer
  Dir            As Byte
  Speed          As Integer
  Layer          As Byte
  AnimFrm        As Byte
  AnimEnabled    As Boolean
  Sprite         As DirectDrawSurface7
  SpriteDesc     As DDSURFACEDESC2
End Type

Private Type TeleportType
  Transition     As Byte
  srcLayerMin    As Byte
  srcLayerMax    As Byte
  SrcX1          As Long
  SrcY1          As Long
  SrcX2          As Long
  SrcY2          As Long
  DestX          As Long
  DestY          As Long
  DestDir        As Byte
  DestLayer      As Byte
  DestMap        As String
End Type

Private Type DAM_InfoType
  MapHeader      As MapHeaderType
  MapData        As MapFileType
  Character      As CharacterType
  Teleport()     As TeleportType
  Telepoters     As Integer
End Type
Public MapInfo   As DAM_InfoType   'Data Dec



Private Type GameMainHeaderType
  Game           As String
  Engine         As String
  Script         As String
  Controls       As String
End Type

Private Type GameMainPathsType
  System         As String
  Maps           As String
  Music          As String
  Sound          As String
  Grafix         As String
  Script         As String
  Main           As String
End Type

Private Type GameMainType
  Header         As GameMainHeaderType
  Paths          As GameMainPathsType
End Type
Public GameMain  As GameMainType



Private Type CharFacePicsType
  DDS            As DirectDrawSurface7
  DDSD           As DDSURFACEDESC2
End Type
Global CharPic(9) As CharFacePicsType


Private Type GameObjectType
  X1 As Integer
  X2 As Integer
  Y1 As Integer
  Y2 As Integer
  Layer1 As Integer
  Layer2 As Integer
  Type As String
  Property1 As String
  Property2 As String
  Property3 As String
  Property4 As String
  Property5 As String
  Property6 As String
  Property7 As String
  Property8 As String
  Property9 As String
  Property0 As String
  Activation As Integer
End Type
Global GameObject(9) As GameObjectType



Global LastError As Single
Global BreakLine As String

Global WhatToDo  As Byte                'This controls what function the render loop will do

Global Const QUITME = 0                'When to quit
Global Const Flag1 = 1
Global Const Flag2 = 2
Global Const Flag3 = 4
Global Const Flag4 = 8
Global Const Flag5 = 16
Global Const Flag6 = 32
Global Const Flag7 = 64
Global Const Flag8 = 128

Private Type EventType
  X              As Integer
  Y              As Integer
  Layer          As Byte
  Event          As Integer
  Special        As String
End Type

Global Eventz(1) As EventType

Sub LoadINI(ByRef FileName As String)
On Error GoTo ErrHan
  Dim a As String
  Dim Section As Integer
  Dim StrLoc As Integer
  Dim i As Integer
  
  Open FileName For Input As #1
    Do
      Line Input #1, a
      If InStr(a, ";") Then
        a = Left(a, InStr(a, ";") - 1)
      End If
            
      a = Trim$(a)
      If Left(a, 1) = "[" And Right(a, 1) = "]" Then
        a = UCase(Mid(a, 2, Len(a) - 2))
        Select Case a
          Case "HEADER"
            Section = 1
          Case "PATHS"
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
                Case "NAME"
                  GameMain.Header.Game = Mid(a, StrLoc + 1)
                Case "MAINSCRIPT"
                  GameMain.Header.Script = Mid(a, StrLoc + 1)
                Case "ENGINE"
                  GameMain.Header.Engine = Mid(a, StrLoc + 1)
                Case "CONTROLS"
                  GameMain.Header.Controls = Mid(a, StrLoc + 1)
              End Select
            End If
          
          Case 2
            If StrLoc Then
              With MapInfo.Teleport(MapInfo.Telepoters)
                Select Case Trim(UCase(Left(a, StrLoc - 1)))
                Case "SYSTEM"
                  ChDir (GameMain.Paths.Main)
                  ChDir (Mid(a, StrLoc + 1))
                  GameMain.Paths.System = CurDir & "\"
                Case "MAPS"
                  ChDir (GameMain.Paths.Main)
                  ChDir (Mid(a, StrLoc + 1))
                  GameMain.Paths.Maps = CurDir & "\"
                Case "SCRIPT"
                  ChDir (GameMain.Paths.Main)
                  ChDir (Mid(a, StrLoc + 1))
                  GameMain.Paths.Script = CurDir & "\"
                Case "MUSIC"
                  ChDir (GameMain.Paths.Main)
                  ChDir (Mid(a, StrLoc + 1))
                  GameMain.Paths.Music = CurDir & "\"
                Case "SOUND"
                  ChDir (GameMain.Paths.Main)
                  ChDir (Mid(a, StrLoc + 1))
                  GameMain.Paths.Sound = CurDir & "\"
                Case "GRAFIX"
                  ChDir (GameMain.Paths.Main)
                  ChDir (Mid(a, StrLoc + 1))
                  GameMain.Paths.Grafix = CurDir & "\"
                End Select
              End With
            End If
        End Select
      End If
    Loop Until EOF(1)
  
  Close #1
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (LoadINI) : " & Err.Description
  Err.Clear
  Close #1
End Sub

Sub LoadControls(ByRef FileName As String)
On Error GoTo ErrHan
  Dim a As String
  Dim Section As Integer
  Dim StrLoc As Integer
  Dim i As Integer
  
  Open FileName For Input As #1
    Do
      Line Input #1, a
      If InStr(a, ";") Then
        a = Left(a, InStr(a, ";") - 1)
      End If
            
      a = Trim$(a)
      If Left(a, 1) = "[" And Right(a, 1) = "]" Then
        a = UCase(Mid(a, 2, Len(a) - 2))
        Select Case a
          Case "CONTROLS"
            Section = 1
        End Select

      Else
        StrLoc = InStr(a, "=")
        Select Case Section
          Case 1
            If StrLoc Then
              With MapInfo.Teleport(MapInfo.Telepoters)
                Select Case UCase(Mid(a, StrLoc + 1))
                Case "OK"
                  ButtonOK = "&H" & (Mid(a, StrLoc - 1, 1))
                  frmMain.RunCommand ("ButtonOK=" & ButtonOK)
                Case "CANCEL"
                  ButtonCancel = "&H" & (Mid(a, StrLoc - 1, 1))
                  frmMain.RunCommand ("ButtonCancel=" & ButtonCancel)
                Case "MENU"
                  ButtonMenu = "&H" & (Mid(a, StrLoc - 1, 1))
                  frmMain.RunCommand ("ButtonMenu=" & ButtonMenu)
                Case "ACTION"
                  ButtonAction = "&H" & (Mid(a, StrLoc - 1, 1))
                  frmMain.RunCommand ("ButtonAction=" & ButtonAction)
                End Select
              End With
            End If
        End Select
      End If
    Loop Until EOF(1)
  
  Close #1
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (LoadControls) : " & Err.Description
  Err.Clear
  Close #1
End Sub

Private Sub UnloadMap()
  'Character
  With MapInfo.Character
    .AnimFrm = 0
    .DefDir = 0
    .DefLayer = 1
    .DefX = 0
    .DefY = 0
    .Dir = 0
    .Layer = 1
    .OX = 0
    .OY = 0
    .Speed = 3.5
    .WX = 0
    .WY = 0
    .X = 0
    .Y = 0
  End With
  
  'Map Data/Information
  With MapInfo.MapData
    .AnimFrame = 0
    .AnimFrame2 = 0
    .Bottom = 1
    ReDim .Flag(1, 1, 3)
    .Layers = 3
    .Left = 0
    .OffsetX = 0
    .OffsetY = 0
    .OSX = 0
    .OSY = 0
    .Right = 1
    ReDim .Tile(1, 1, 3)
    .Top = 0
    ReDim .Walk(1, 1, 3)
  End With
  
  'Map Header
  With MapInfo.MapHeader
    .CodeScript = ""
    .FontSet = ""
    .Map = ""
    .Pallet = ""
    .TileSet = ""
    .Title = ""
  End With
  
  'Teleporters
  MapInfo.Telepoters = 0
  ReDim MapInfo.Teleport(0)
End Sub

Public Sub LoadMap(ByRef FileName As String)
On Error GoTo ErrHan
  Dim a As String
  Dim Section As Integer
  Dim StrLoc As Integer
  Dim i As Integer
  Dim CurObject As Integer
  CurObject = -1
  
  UnloadMap
  Open FileName For Input As #1
    Do
      Line Input #1, a
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
          Case "OBJECT"
            Section = 3
            CurObject = CurObject + 1
          Case Else
            Section = 0
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
                  If Not FullScreen Then
                    frmMain.Caption = FrmCaption & "  --  " & MapInfo.MapHeader.Title
                  End If
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
        
          Case 3
            If StrLoc Then
              Select Case Trim(UCase(Left(a, StrLoc - 1)))
                Case "X1"
                  GameObject(CurObject).X1 = CInt(Mid(a, StrLoc + 1))
                Case "Y1"
                  GameObject(CurObject).Y1 = CInt(Mid(a, StrLoc + 1))
                Case "X2"
                  GameObject(CurObject).X2 = CInt(Mid(a, StrLoc + 1))
                Case "Y2"
                  GameObject(CurObject).Y2 = CInt(Mid(a, StrLoc + 1))
                Case "LAYER1"
                  GameObject(CurObject).Layer1 = CInt(Mid(a, StrLoc + 1))
                Case "LAYER2"
                  GameObject(CurObject).Layer2 = CInt(Mid(a, StrLoc + 1))
                Case "TYPE"
                  GameObject(CurObject).Type = Mid(a, StrLoc + 1)
                Case "PROPERTY1"
                  GameObject(CurObject).Property1 = Mid(a, StrLoc + 1)
                Case "PROPERTY2"
                  GameObject(CurObject).Property2 = Mid(a, StrLoc + 1)
                Case "PROPERTY3"
                  GameObject(CurObject).Property3 = Mid(a, StrLoc + 1)
                Case "PROPERTY4"
                  GameObject(CurObject).Property4 = Mid(a, StrLoc + 1)
                Case "PROPERTY5"
                  GameObject(CurObject).Property5 = Mid(a, StrLoc + 1)
                Case "PROPERTY6"
                  GameObject(CurObject).Property6 = Mid(a, StrLoc + 1)
                Case "PROPERTY7"
                  GameObject(CurObject).Property7 = Mid(a, StrLoc + 1)
                Case "PROPERTY8"
                  GameObject(CurObject).Property8 = Mid(a, StrLoc + 1)
                Case "PROPERTY9"
                  GameObject(CurObject).Property9 = Mid(a, StrLoc + 1)
                Case "PROPERTY0"
                  GameObject(CurObject).Property0 = Mid(a, StrLoc + 1)
                Case "ACTIVATION"
                  GameObject(CurObject).Activation = CInt(Mid(a, StrLoc + 1))
              End Select
            End If
        
        End Select
      End If
    Loop Until EOF(1)
  
  Close #1
  
  MapInfo.MapData.Visible = True
  LoadPictureMap (GameMain.Paths.Maps & MapInfo.MapHeader.Map)
  
  'LoadSystemScript (GameMain.Paths.Script & "System.DAC")
  LoadScript (GameMain.Paths.Script & MapInfo.MapHeader.CodeScript)
  
  CenterDI
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (LoadMap) : " & Err.Description
  Err.Clear
  Close 1
End Sub


Sub LoadPictureMap(ByRef FileName As String)
On Error GoTo ErrHan
  Dim i1 As Integer, i2 As Integer, i3 As Integer

  With MapInfo.MapData
    Open FileName For Binary As #1
      Get #1, , .Bottom
      Get #1, , .Left
      Get #1, , .Right
      Get #1, , .Top
      Get #1, , .Layers
      .Layers = 3
    
      ReDim .Tile(.Left To .Right, .Top To .Bottom, 4)
      ReDim .Walk(.Left To .Right, .Top To .Bottom, 4)
      ReDim .Flag(.Left To .Right, .Top To .Bottom, 4)
    
      For i1 = .Left To .Right
        For i2 = .Top To .Bottom
          For i3 = 0 To 3
            Get #1, , MapInfo.MapData.Tile(i1, i2, i3)
            Get #1, , MapInfo.MapData.Walk(i1, i2, i3)
            Get #1, , MapInfo.MapData.Flag(i1, i2, i3)
          Next i3
        Next i2
      Next i1
        
    Close #1
  End With
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (LoadPicMap) : " & Err.Description
  Err.Clear
End Sub

Public Sub LoadScript(ByRef FileName As String)
On Error GoTo ErrHan
  
  Dim ScriptText As String
  Dim TextTemp As String
  
  Open GameMain.Paths.Script & "System.DAC" For Input As #11
    ScriptText = Input(LOF(11), 11)
  Close #11
  With frmMain.SC
    Call .AddCode(ScriptText)
  End With
  
  Open FileName For Input As #10
    ScriptText = Input(LOF(10), 10)
  Close #10
  With frmMain.SC
    'Call .AddObject("Game", SharedData)
    Call .AddCode(ScriptText)
  End With
Exit Sub

ErrHan:
  Dim ErrText As String
  LastError = Err.Number
  If frmMain.SC.Error.Number Then
    ErrText = ErrText & BreakLine & "Error (LoadScript): "
    ErrText = ErrText & frmMain.SC.Error.Number & " - "
    ErrText = ErrText & frmMain.SC.Error.Line & " - "
    ErrText = ErrText & frmMain.SC.Error.Description & " - "
    ErrText = ErrText & frmMain.SC.Error.Source
  Else
    ErrText = ErrText & BreakLine & "Error (LoadScript): "
    ErrText = ErrText & Err.Number & " - "
    ErrText = ErrText & Err.Description & " - "
    ErrText = ErrText & Err.Source
  End If
  frmMain.SC.Error.Clear
  Err.Clear
  Debuger.DWrite ErrText
End Sub

Public Function AllowWalk(X As Integer, Y As Integer) As Boolean
  Select Case MapInfo.MapData.Walk(X, Y, MapInfo.Character.Layer)
    Case 0
      AllowWalk = True
    Case 1
      AllowWalk = False
    Case 2
      AllowWalk = True
    Case 3
      AllowWalk = True
    Case 4
      MapInfo.Character.Layer = 0
      AllowWalk = True
    Case 5
      MapInfo.Character.Layer = 1
      AllowWalk = True
    Case 6
      MapInfo.Character.Layer = 2
      AllowWalk = True
    Case 7
      MapInfo.Character.Layer = 3
      AllowWalk = True
  End Select
  
End Function

Public Function FileExist(FileName As String) As Boolean
  Err.Clear
  
  On Error Resume Next
  Open FileName For Input As #2
  If Err.Number Then
    Close #2
    Err.Clear
    FileExist = False
  Else
    Close #2
    FileExist = True
  End If
End Function
