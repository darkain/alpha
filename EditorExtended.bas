Attribute VB_Name = "EditorExtended"
Option Explicit

Global FileHasChanged   As Boolean
Global EdType           As Integer
Global OpenedFile       As String
Global Const MaxWH = 256


'Tool Bar Buttons
Global Const TBar_Button_New = 1
Global Const TBar_Button_Open = 2
Global Const TBar_Button_Save = 3
Global Const TBar_Button_Sep1 = 5
Global Const TBar_Button_Layer1 = 5
Global Const TBar_Button_Layer2 = 6
Global Const TBar_Button_Layer3 = 7
Global Const TBar_Button_Layer4 = 8
Global Const TBar_Button_Sep2 = 9
Global Const TBar_Button_Grid = 10
Global Const TBar_Button_Obst = 11
Global Const TBar_Button_Tile = 12
Global Const TBar_Button_Anim = 13

'Active Window
Global Const Win_DAN = 1
Global Const Win_MapEdit = 2
Global Const Win_Logo = 3
Global Const Win_MIDI = 4
Global Const Win_PicView = 5
Global Const Win_TextEdit = 6
Global Const Win_MenuEdit = 7
Global ActiveWin        As Integer

'Drawing Tools
Global Const Tools_Draw = 1
Global Const Tools_Flood = 2
Global Const Tools_Cut = 3
Global Const Tools_Copy = 4
Global Const Tools_Paste = 5
Global Const Tools_Delete = 6
Global Const Tools_Inert = 7
Global Const Tools_Move = 8
Global Const Tools_PickUp = 9
Global Const Tools_Select = 10
Global Const Tools_Zoom11 = 11
Global Const Tools_Zoom21 = 12
Global Const Tools_Zoom12 = 13
Global DrawTool         As Integer


Global LastButton       As Integer
Global NodeString       As String
Global FrmCaption       As String
Global FileListDir      As String


Global Orig_X           As Integer
Global Orig_Y           As Integer
Global Orig_Z           As Integer
Global Orig_X2          As Integer
Global Orig_Y2          As Integer
Global Orig_Z2          As Integer

Private Type SelTileType
  Width                 As Integer
  Height                As Integer
  MouseDown             As Boolean
  Tiles()               As Integer
End Type
Global SelTiles         As SelTileType


Global TileList(9, 255) As Integer
Global TileListName(9)  As String * 20

Global Loading             As Boolean


Private Type MenuXTextType
  X               As Integer
  Y               As Integer
  Text            As String
  Colour          As Long
End Type
Private Type MenuXType
  X1              As Integer
  Y1              As Integer
  X2              As Integer
  Y2              As Integer
  Tiles(19, 14)   As Integer
  Text()          As MenuXTextType
  TextSpeed       As Integer
  TextStyle       As Integer
  Font            As New StdFont
End Type
Global MenuX()    As MenuXType


'******************************
'this function is used to clear all variables
'so that a "new "file" is made within the editor
'******************************
Public Sub New_File()
  'Reset map data
  ReDim MapInfo.MapData.Walk(-MaxWH To MaxWH, -MaxWH To MaxWH, 3)
  ReDim MapInfo.MapData.Flag(-MaxWH To MaxWH, -MaxWH To MaxWH, 3)
  ReDim MapInfo.MapData.Tile(-MaxWH To MaxWH, -MaxWH To MaxWH, 3)
  ReDim MapInfo.Teleport(0)
  MapInfo.Telepoters = -1
  
  'Reset Cahracter Information
  MapInfo.Character.DefX = 0
  MapInfo.Character.DefY = 0
  MapInfo.Character.DefDir = 0
  MapInfo.Character.DefLayer = 0
  
  'Reset map deminsions
  MapInfo.MapData.Top = 0
  MapInfo.MapData.Bottom = 1
  MapInfo.MapData.Left = 0
  MapInfo.MapData.Right = 1
  MapInfo.MapData.Layers = 3
  
  'reset map header
  MapInfo.MapHeader.CodeScript = ""
  MapInfo.MapHeader.FontSet = ""
  MapInfo.MapHeader.Map = ""
  MapInfo.MapHeader.Pallet = ""
  MapInfo.MapHeader.Title = ""
  
  EdType = 0
  OpenedFile = ""
  FileHasChanged = False
  
  'Reset selected tile info
  ReDim SelTiles.Tiles(0, 0)
  SelTiles.Height = 0
  SelTiles.Width = 0
  
  FrmMap.Bar.Panels(2).Text = ""
  
  
  'clear map picture on screen
  FrmMap.Map2.Refresh
  'Map2_Paint
End Sub


'******************************
'this function is used to load *.DAN files
'(text based specifics on a map)
'******************************
Public Sub Load_DAN(FileName As String)
  
  If Not FileExist(FileName) Then
    MsgBox "File Not Found" & vbCrLf & FileName, vbCritical Or vbOKOnly, "Darkain Editor"
    Exit Sub
  End If
  
  FrmMap.DAN_Info.FileName = FileName
  FrmMap.Bar.Panels(2).Text = NodeString
  FrmMap.MakeVis (Win_DAN)
  
  OpenedFile = FileName
End Sub


'******************************
'this function is used to save *.DAN files
'******************************
Sub Save_DAN(FileName As String)
  Dim i As Integer
  
  Open FileName For Output As #1
    
    Print #1, ";Created With Darkain's Alpha Editor - Version: " & App.Major & "." & App.Minor & "." & App.Revision
    
    Print #1, "[Header]"
    Print #1, "Map=" & MapInfo.MapHeader.Map
    Print #1, "Title=" & MapInfo.MapHeader.Title
    Print #1, "CodeScript=" & MapInfo.MapHeader.CodeScript
    Print #1, "Tileset=" & MapInfo.MapHeader.TileSet
    Print #1, "FontSet=" & MapInfo.MapHeader.FontSet
    Print #1, "Pallet=" & MapInfo.MapHeader.Pallet
    Print #1, ""
    
    For i = 0 To MapInfo.Telepoters
      Print #1, "[Teleport]"
      Print #1, "Transition=" & MapInfo.Teleport(i).Transition
      Print #1, "SrcX1=" & MapInfo.Teleport(i).SrcX1
      Print #1, "SrcY1=" & MapInfo.Teleport(i).SrcY1
      Print #1, "SrcX2=" & MapInfo.Teleport(i).SrcX2
      Print #1, "SrcY2=" & MapInfo.Teleport(i).SrcY2
      Print #1, "SrcLayerMin=" & MapInfo.Teleport(i).srcLayerMin
      Print #1, "SrcLayerMax=" & MapInfo.Teleport(i).srcLayerMax
      Print #1, "DestX=" & MapInfo.Teleport(i).DestX
      Print #1, "DestY=" & MapInfo.Teleport(i).DestY
      Print #1, "DestDir=" & MapInfo.Teleport(i).DestDir
      Print #1, "DestLayer=" & MapInfo.Teleport(i).DestLayer
      Print #1, "DestMap=" & MapInfo.Teleport(i).DestMap
      Print #1, ""
    Next i
  
  Close 1
  FileHasChanged = False
  
  FrmMap.Bar.Panels(2).Text = NodeString
  OpenedFile = FileName
End Sub


'******************************
'this function is used to load *.DAM files
'(the picture of the map)
'******************************
Public Sub Load_DAM(FileName As String)
On Error GoTo ErrHan
  
  If Not FileExist(FileName) Then
    MsgBox "File Not Found" & vbCrLf & FileName, vbCritical Or vbOKOnly, "Darkain Editor"
    Exit Sub
  End If
  
  Dim i1 As Integer, i2 As Integer, i3 As Integer

  Call New_File

  Open FileName For Binary As #1
    Get #1, , MapInfo.MapData.Bottom
    Get #1, , MapInfo.MapData.Left
    Get #1, , MapInfo.MapData.Right
    Get #1, , MapInfo.MapData.Top
    Get #1, , MapInfo.MapData.Layers
    
    For i1 = MapInfo.MapData.Left To MapInfo.MapData.Right
      For i2 = MapInfo.MapData.Top To MapInfo.MapData.Bottom
        For i3 = 0 To 3
          Get #1, , MapInfo.MapData.Tile(i1, i2, i3)
          Get #1, , MapInfo.MapData.Walk(i1, i2, i3)
          Get #1, , MapInfo.MapData.Flag(i1, i2, i3)
        Next i3
      Next i2
    Next i1
  Close #1
  
  MapInfo.MapHeader.Map = FileName
  
  FrmMap.Bar.Panels(2).Text = NodeString
  FrmMap.MakeVis (Win_MapEdit)
  
  FrmMap.Map2.Refresh
  
  OpenedFile = FileName
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Load_DAM): " & Err.Number & " " & Err.Description
  Call New_File
End Sub


'******************************
'this function is used to save *.DAM files
'******************************
Sub Save_DAM(FileName As String)
On Error GoTo ErrHan
  
  Dim i1 As Integer, i2 As Integer, i3 As Integer
  Dim Tmp As Integer, Tmp1 As Integer

  If FileExist(FileName) Then Kill (FileName)
  
  Tmp1 = 32767
  For i1 = MapInfo.MapData.Left To MapInfo.MapData.Right
    For i2 = MapInfo.MapData.Top To MapInfo.MapData.Bottom
      For i3 = 0 To 3
        If MapInfo.MapData.Tile(i1, i2, i3) Or MapInfo.MapData.Walk(i1, i2, i3) Or MapInfo.MapData.Flag(i1, i2, i3) Then
          If i1 < Tmp1 Then
            Tmp1 = i1 - 1
          End If
          Tmp = i1
        End If
      Next i3
    Next i2
  Next i1
  If Tmp1 = 32767 Then
    MapInfo.MapData.Left = 1
  Else
    MapInfo.MapData.Left = Tmp1
  End If
  MapInfo.MapData.Right = Tmp

  Tmp1 = 32767
  For i1 = MapInfo.MapData.Top To MapInfo.MapData.Bottom
    For i2 = MapInfo.MapData.Left To MapInfo.MapData.Right
      For i3 = 0 To 3
        If MapInfo.MapData.Tile(i2, i1, i3) Or MapInfo.MapData.Walk(i2, i1, i3) Or MapInfo.MapData.Flag(i2, i1, i3) Then
          If i1 < Tmp1 Then
            Tmp1 = i1 - 1
          End If
          Tmp = i1
        End If
      Next i3
    Next i2
  Next i1
  If Tmp1 = 32767 Then
    MapInfo.MapData.Top = 1
  Else
    MapInfo.MapData.Top = Tmp1
  End If
  MapInfo.MapData.Bottom = Tmp

  If MapInfo.MapData.Top + MapInfo.MapData.Bottom < 15 Then
    MapInfo.MapData.Bottom = MapInfo.MapData.Top + 17 + MapInfo.MapData.Bottom
  End If
  If MapInfo.MapData.Left + MapInfo.MapData.Right < 20 Then
    MapInfo.MapData.Right = MapInfo.MapData.Left + 22 + MapInfo.MapData.Right
  End If

  Open FileName For Binary As #1
    Put #1, , MapInfo.MapData.Bottom
    Put #1, , MapInfo.MapData.Left
    Put #1, , MapInfo.MapData.Right
    Put #1, , MapInfo.MapData.Top
    Put #1, , MapInfo.MapData.Layers
    For i1 = MapInfo.MapData.Left To MapInfo.MapData.Right
      For i2 = MapInfo.MapData.Top To MapInfo.MapData.Bottom
        For i3 = 0 To 3
          Put #1, , MapInfo.MapData.Tile(i1, i2, i3)
          Put #1, , MapInfo.MapData.Walk(i1, i2, i3)
          Put #1, , MapInfo.MapData.Flag(i1, i2, i3)
        Next i3
      Next i2
    Next i1
      
  Close #1
  FileHasChanged = False
  
  FrmMap.Bar.Panels(2).Text = NodeString
  OpenedFile = FileName
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Load_DAM): " & Err.Number & " " & Err.Description
  Call New_File
End Sub

'******************************
'this function is used to load *.DAI files
'(basically INI files)
'******************************
Public Sub Load_DAI(FileName As String)
On Error GoTo ErrHan

  If Not FileExist(FileName) Then
    MsgBox "File Not Found" & vbCrLf & FileName, vbCritical Or vbOKOnly, "Darkain Editor"
    Exit Sub
  End If
  
  Loading = True
  Dim a As Long
  
  Open FileName For Input As #1
    a = LOF(1)
    FrmMap.EditINI.Text = Input(a, 1)
  Close 1

  FrmMap.MainFrame.Caption = FileName
  FrmMap.Bar.Panels(2).Text = NodeString
  FrmMap.MakeVis (Win_TextEdit)
  
  OpenedFile = FileName
  Loading = False
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Load_DAI): " & Err.Number & " " & Err.Description
  Loading = False
End Sub


'******************************
'this function is used to save *.DAI files
'******************************
Public Sub Save_DAI(FileName As String)
On Error GoTo ErrHan

  Open FileName For Output As #1
    Print #1, FrmMap.EditINI.Text
  Close 1

  FrmMap.MainFrame.Caption = FileName
  FrmMap.Bar.Panels(2).Text = NodeString
  FrmMap.MakeVis (Win_TextEdit)
  
  OpenedFile = FileName
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Load_DAI): " & Err.Number & " " & Err.Description
End Sub


'******************************
'this function is used to load *.DAT files
'(names for tile catagories)
'******************************
Public Sub Load_DAT(FileName As String)
On Error GoTo ErrHan

  If Not FileExist(FileName) Then
    MsgBox "File Not Found" & vbCrLf & FileName, vbCritical Or vbOKOnly, "Darkain Editor"
    Exit Sub
  End If
  
  Dim i1 As Integer
  Dim i2 As Integer
  
  Open FileName For Binary As 1
    For i1 = 0 To 9
      Get 1, , TileListName(i1)
      Call FrmMap.TileCat.AddItem(TileListName(i1), i1)
      
      For i2 = 0 To 255
        Get 1, , TileList(i1, i2)
      Next i2
      
    Next i1
  Close 1
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Load_DAI): " & Err.Number & " " & Err.Description
End Sub

Public Sub Save_DAT(FileName As String)
On Error GoTo ErrHan
  Dim i1 As Integer
  Dim i2 As Integer
  
  Open FileName For Binary As 1
    For i1 = 0 To 9
      Put 1, , TileListName(i1)
      
      For i2 = 0 To 255
        Put 1, , TileList(i1, i2)
      Next i2
      
    Next i1
  Close 1
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Load_DAI): " & Err.Number & " " & Err.Description
End Sub

Public Sub Load_DAY(FileName As String)
On Error GoTo ErrHan
  Dim a As String
  Dim MenuSection As Integer
  Dim TextSection As Integer
  Dim Section As Integer
  Dim StrLoc As Integer
  
  ReDim MenuX(0)
  ReDim MenuX(0).Text(0)
  
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
          Case "MENUX"
            MenuSection = 0
            TextSection = 0
            Section = 1
          Case "TEXT"
            TextSection = 0
            Section = 2
        End Select

      Else
        StrLoc = InStr(a, "=")
        Select Case Section
          Case 1
            If StrLoc Then
              Select Case Trim(UCase(Left(a, StrLoc - 1)))
                Case "MENUX_ID"
                  MenuSection = CInt(Mid(a, StrLoc + 1))
                  If MenuSection > UBound(MenuX) Then
                    ReDim Preserve MenuX(MenuSection)
                    ReDim MenuX(MenuSection).Text(0)
                  End If
                Case "FONT.BOLD"
                  MenuX(MenuSection).Font.Bold = CBool(Mid(a, StrLoc + 1))
                Case "FONT.ITALIC"
                  MenuX(MenuSection).Font.Italic = CBool(Mid(a, StrLoc + 1))
                Case "FONT.NAME"
                  MenuX(MenuSection).Font.Name = CStr(Mid(a, StrLoc + 1))
                Case "FONT.SIZE"
                  MenuX(MenuSection).Font.Size = CSng(Mid(a, StrLoc + 1))
                Case "FONT.STRIKETHROUGH"
                  MenuX(MenuSection).Font.Strikethrough = CBool(Mid(a, StrLoc + 1))
                Case "FONT.UNDERLINE"
                  MenuX(MenuSection).Font.Underline = CBool(Mid(a, StrLoc + 1))
                
                Case "X1"
                  MenuX(MenuSection).X1 = CInt(Mid(a, StrLoc + 1))
                Case "Y1"
                  MenuX(MenuSection).Y1 = CInt(Mid(a, StrLoc + 1))
                Case "X2"
                  MenuX(MenuSection).X2 = CInt(Mid(a, StrLoc + 1))
                Case "Y2"
                  MenuX(MenuSection).Y2 = CInt(Mid(a, StrLoc + 1))
              End Select
            End If
            
          Case 2
            If StrLoc Then
              Select Case Trim(UCase(Left(a, StrLoc - 1)))
                Case "TEXT_ID"
                  TextSection = CInt(Mid(a, StrLoc + 1))
                  If TextSection > UBound(MenuX(MenuSection).Text) Then
                    ReDim Preserve MenuX(MenuSection).Text(TextSection)
                  End If
                Case "TEXT"
                  MenuX(MenuSection).Text(TextSection).Text = CStr(Mid(a, StrLoc + 1))
                Case "COLOR"
                  MenuX(MenuSection).Text(TextSection).Colour = CLng(Mid(a, StrLoc + 1))
                Case "X"
                  MenuX(MenuSection).Text(TextSection).X = CInt(Mid(a, StrLoc + 1))
                Case "Y"
                  MenuX(MenuSection).Text(TextSection).Y = CInt(Mid(a, StrLoc + 1))
              End Select
            End If
        End Select
      End If
    Loop Until EOF(1)
  
  Close #1
  
  FrmMap.Bar.Panels(2).Text = NodeString
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Load_DAY): " & Err.Number & " " & Err.Description
End Sub

Public Sub Save_DAY(FileName As String)
On Error GoTo ErrHan
  
  FrmMap.Bar.Panels(2).Text = NodeString
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Save_DAY): " & Err.Number & " " & Err.Description
End Sub

Public Sub Load_DAX(FileName As String)
On Error GoTo ErrHan
  Dim i1 As Integer
  Dim i2 As Integer
  Dim i3 As Integer
  
  Open GameMain.Paths.System & "Menu.DAX" For Binary As 99
    Put 99, , UBound(MenuX)

    For i1 = 0 To UBound(MenuX)
      With MenuX(i1)
        For i2 = 0 To 14
          For i3 = 0 To 19
            Get 99, , .Tiles(i3, i2)
          Next i3
        Next i2
      End With
    Next i1
  Close 99
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Load_DAX): " & Err.Number & " " & Err.Description
End Sub

Public Sub Save_DAX(FileName As String)
On Error GoTo ErrHan
  
  FrmMap.Bar.Panels(2).Text = NodeString
Exit Sub

ErrHan:
  Close
  MsgBox "Error Occured (Save_DAX): " & Err.Number & " " & Err.Description
End Sub


Sub LoadINI(ByRef FileName As String)
On Error GoTo ErrHan
  Dim a As String
  Dim Section As Integer
  Dim StrLoc As Integer
  Dim i As Integer
  
  Open FileName For Input As #1
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
              End Select
            End If
          
          Case 2
            If StrLoc Then
              With MapInfo.Teleport(MapInfo.Telepoters)
                Select Case Trim(UCase(Left(a, StrLoc - 1)))
                Case "ROOT"
                  ChDir (GameMain.Paths.Main)
                  ChDir (Mid(a, StrLoc + 1))
                  GameMain.Paths.Root = CurDir & "\"
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
  MsgBox Err.Number & " (LoadINI) : " & Err.Description
  Err.Clear
  Close #1
End Sub
