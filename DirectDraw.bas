Attribute VB_Name = "DirectDrawDK"
Option Explicit

Global ShowFrameRate    As Boolean
Global UsePageFlip      As Boolean
Global FullScreen       As Boolean
Dim OriginalX           As Single
Dim OriginalY           As Single
Dim FirstTime           As Long
Dim SecondTime          As Long
Dim NumLoops            As Integer
Dim FPS                 As Integer


Dim DD                  As DirectDraw7            'the main DirectDraw object

Dim ddsPrimary          As DirectDrawSurface7     'the primary directdraw surface
Dim ddsdPrimary         As DDSURFACEDESC2         'the primary surface's description
Dim ddsBack             As DirectDrawSurface7     'the backbuffer surface
Dim ddsdBack            As DDSURFACEDESC2         'the backbuffer surface's description
Dim PrimDDClip          As DirectDrawClipper      'Clipper

Global KEY              As DDCOLORKEY


Dim rScreen As RECT                    'rect variable - useful to have around :P
Dim TilesWide As Integer
Dim rPrim As RECT
Dim ST As Long
Dim i1 As Integer, i2 As Integer
Dim i3 As Integer, i4 As Byte
Dim OffX As Integer, OffY As Integer
Dim TmpX As Integer, TmpY As Integer
Dim CurTile As Integer
Dim TmpHW As Integer
Dim TileRect As RECT



Dim ScreenBitDepth As Byte
Global ActiveWin As Integer
Global Const WalkWin = 1
Global Const MenuWin = 2
Global Const ItemWin = 3
Const CharactersPerLine = 34


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
  Visible         As Boolean
End Type
Global MenuX()    As MenuXType
Global DDS_MenuX  As DirectDrawSurface7
Global DDSD_MenuX As DDSURFACEDESC2
Global DDS_MenuX_Tiles  As DirectDrawSurface7
Global DDSD_MenuX_Tiles As DDSURFACEDESC2
Global MenuX_R1   As RECT
Global MenuX_R2   As RECT
Global MenuX_R3   As RECT
Global MenuXsc()  As New MenuClass



Type Offsets
  X_Pix           As Single
  Y_Pix           As Single
  X_Tile          As Integer
  Y_Tile          As Integer
  X_Dir           As Integer
  Y_Dir           As Integer
End Type
Dim Offset As Offsets


Public Sub InitDD()
On Error GoTo ErrHan
  Set DD = dx.DirectDrawCreate("")

  DD.GetCaps hwCaps, helCaps
  WhatToDo = 1
  
  Offset.X_Pix = 32
  Offset.Y_Pix = 32
  
  If FullScreen Then
    OriginalX = frmMain.Left
    OriginalY = frmMain.Top
    frmMain.Top = -1500
    Call DD.SetCooperativeLevel(frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE)
    DD.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
    ddsdPrimary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsdPrimary.lBackBufferCount = 1
    Set ddsPrimary = DD.CreateSurface(ddsdPrimary)
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set ddsBack = ddsPrimary.GetAttachedSurface(caps)
  Else
    frmMain.Caption = FrmCaption & "  --  " & MapInfo.MapHeader.Title
    frmMain.Left = OriginalX
    frmMain.Top = OriginalY
    frmMain.Width = 640 * Screen.TwipsPerPixelX + (frmMain.Width - frmMain.ScaleWidth)
    frmMain.Height = 480 * Screen.TwipsPerPixelY + (frmMain.Height - frmMain.ScaleHeight)
    'frmMain.Picture1.Width = 640 * Screen.TwipsPerPixelX
    'frmMain.Picture1.Height = 480 * Screen.TwipsPerPixelY
    Call DD.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    ddsdPrimary.lFlags = DDSD_CAPS
    ddsdPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set ddsPrimary = DD.CreateSurface(ddsdPrimary)
    ddsdBack.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsdBack.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsdBack.lWidth = 640
    ddsdBack.lHeight = 480
    Set ddsBack = DD.CreateSurface(ddsdBack)
  
    Set PrimDDClip = DD.CreateClipper(0)
    PrimDDClip.SetHWnd frmMain.hWnd
    ddsPrimary.SetClipper PrimDDClip
  End If
  InitSurfaces
  
  
'  DispText.DDS_Buffer.SetForeColor RGB(255, 255, 255)
'  DispText.Font.Name = "Terminal"
'  DispText.Font.Size = 18
'  DispText.Font.Bold = False
'  DispText.DDS_Buffer.SetFont DispText.Font
'  DispText.Speed = 1

'  DispMenu.Font.Name = "Terminal"
'  DispMenu.Font.Size = 12
'  DispMenu.DDS.SetFont DispMenu.Font

'  DispMenu2.Font.Name = "Terminal"
'  DispMenu2.Font.Size = 12
'  DispMenu2.DDS.SetFont DispMenu.Font

  rScreen.Left = 0
  rScreen.Top = 0
  rScreen.Right = 640
  rScreen.Bottom = 480
  
  If ActiveWin = 0 Then
    ActiveWin = WalkWin
    frmMain.RunCommand ("ActiveWin=" & WalkWin)
  End If
  
  TilesWide = MapInfo.MapData.ddsdTiles.lWidth \ 32
  
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (DD_Init) : " & Err.Description & " : This Is BAD!!"
  Err.Clear
End Sub

Sub InitSurfaces()
On Error GoTo ErrHan
  'init Surfaces
  Dim ddsSysColour        As DirectDrawSurface7
  Dim ddsdSysColour       As DDSURFACEDESC2
  'Dim hModule As Long

  'hModule = LoadLibrary(GameMain.Paths.System & "ResOnly")
  
  'Get Screen Colour Depth
  ddsPrimary.Lock rScreen, ddsdPrimary, DDLOCK_WAIT, 0
    ScreenBitDepth = (ddsdPrimary.lPitch \ ddsdPrimary.lWidth) * 8
  ddsPrimary.Unlock rScreen
  
  With MapInfo.MapData
    
    'Set Transparency
    Set ddsSysColour = DD.CreateSurfaceFromFile(GameMain.Paths.Grafix & "System Colourz.bmp", ddsdSysColour)
    
    If ScreenBitDepth = 32 Or ScreenBitDepth = 24 Then
      KEY.low = 16711935
    Else
'      key.low = 63519
      .rTiles.Top = 0
      .rTiles.Bottom = 1
      .rTiles.Left = 0
      .rTiles.Right = 8
      ddsSysColour.Lock .rTiles, ddsdSysColour, DDLOCK_WAIT, 0
        KEY.low = ddsSysColour.GetLockedPixel(0, 0)
      ddsSysColour.Unlock .rTiles
    End If
    KEY.high = KEY.low
    
    'Animated Tilez
    .ddsdTiles2.lFlags = DDSD_CAPS
    .ddsdTiles2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set .ddsTiles2 = DD.CreateSurfaceFromFile(GameMain.Paths.Grafix & "WaterAnim.bmp", .ddsdTiles2)
    .ddsTiles2.SetColorKey DDCKEY_SRCBLT, KEY
    
    'Character (AKA Billy)
    MapInfo.Character.SpriteDesc.lFlags = DDSD_CAPS
    MapInfo.Character.SpriteDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set MapInfo.Character.Sprite = DD.CreateSurfaceFromFile(GameMain.Paths.Grafix & "Man.Bmp", MapInfo.Character.SpriteDesc)
    MapInfo.Character.Sprite.SetColorKey DDCKEY_SRCBLT, KEY

    'Map Tiles
    .ddsdTiles.lFlags = DDSD_CAPS
    .ddsdTiles.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN 'Or DDSCAPS_NONLOCALVIDMEM
    Set .ddsTiles = DD.CreateSurfaceFromFile(GameMain.Paths.Grafix & "tileset.bmp", .ddsdTiles)
    .ddsTiles.SetColorKey DDCKEY_SRCBLT, KEY
    
    'Popup window grafix
    DDSD_MenuX_Tiles.lFlags = DDSD_CAPS
    DDSD_MenuX_Tiles.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set DDS_MenuX_Tiles = DD.CreateSurfaceFromFile(GameMain.Paths.Grafix & "Text3.bmp", DDSD_MenuX_Tiles)
    DDS_MenuX_Tiles.SetColorKey DDCKEY_SRCBLT, KEY
    
    'MenuX Sub-System
    MenuX_R1.Top = 0
    MenuX_R1.Left = 0
    MenuX_R1.Right = 640
    MenuX_R1.Bottom = 480
    DDSD_MenuX.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_MenuX.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    DDSD_MenuX.lWidth = MenuX_R1.Right
    DDSD_MenuX.lHeight = MenuX_R1.Bottom
    Set DDS_MenuX = DD.CreateSurface(DDSD_MenuX)
    DDS_MenuX.SetColorKey DDCKEY_SRCBLT, KEY
    Call DDS_MenuX.BltColorFill(MenuX_R1, KEY.low)
  
  End With
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (DD_Surfaces) : " & Err.Description
  Err.Clear
End Sub

Public Sub StartKillDD()
  DD.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
End Sub

Public Sub KillDD()
  Set PrimDDClip = Nothing
  Set ddsBack = Nothing
  Set ddsPrimary = Nothing
  Set DD = Nothing
End Sub


Private Sub DrawMenu(MenuNum As Integer)
  If MenuNum = 5 Then
    
    For i1 = 0 To 14
      For i2 = 0 To 14
        MenuX(5).Tiles(i1, i2) = MenuX(4).Tiles(i1, i2)
      Next i2
    Next i1
  
    MenuX(5).Text(0).Text = "Yo"
    'MenuX(5).Text(1).Text = Item(HeldItem(2).Number).Name
    'MenuX(5).Text(2).Text = Item(HeldItem(3).Number).Name
    'MenuX(5).Text(1).Text = Item(HeldItem(1).Number).Name
  
  End If
End Sub

Public Sub UpdateMenuSystem()
  Dim X1 As Integer
  Dim X2 As Integer
  Dim Y1 As Integer
  Dim Y2 As Integer
  X1 = 20
  X2 = -1
  Y1 = 15
  Y2 = -1
  
  
  Call DDS_MenuX.BltColorFill(MenuX_R1, KEY.low)
  
  For i3 = 1 To UBound(MenuX)
    If MenuX(i3).Visible Then
      If Not frmMain.RunCommand("DrawMenu(" & i3 & ")", True) Then
        DrawMenu (i3)
      End If
        
        With MenuX(i3)
          If .X1 < X1 Then X1 = .X1
          If .X2 > X2 Then X2 = .X2
          If .Y1 < Y1 Then Y1 = .Y1
          If .Y2 > Y2 Then Y2 = .Y2
        
          For i1 = .X1 To .X2
            For i2 = .Y1 To .Y2
              CurTile = .Tiles(i1, i2)
              If CurTile <> 0 Then
                MenuX_R2.Top = (CurTile \ 8) * 32
                MenuX_R2.Left = (CurTile Mod 8) * 32
                MenuX_R2.Bottom = MenuX_R2.Top + 32
                MenuX_R2.Right = MenuX_R2.Left + 32
                Call DDS_MenuX.BltFast(i1 * 32, i2 * 32, DDS_MenuX_Tiles, MenuX_R2, DDBLTFAST_WAIT)
              End If
        
            Next i2
          Next i1
        End With
      
        DDS_MenuX.SetFont MenuX(i3).Font
        For i1 = 0 To UBound(MenuX(i3).Text)
          With MenuX(i3).Text(i1)
            If Len(.Text) > 0 Then
              Call DDS_MenuX.SetForeColor(.Colour)
              Call DDS_MenuX.DrawText(.X, .Y, .Text, False)
            End If
          End With
        Next i1
    
    End If
  Next i3
  
  Select Case True
    Case (X1 > X2), (Y1 > Y2)
      MenuX_R3.Left = 0
      MenuX_R3.Right = 640
      MenuX_R3.Top = 0
      MenuX_R3.Bottom = 480
    Case Else
      MenuX_R3.Left = X1 * 32
      MenuX_R3.Right = (X2 + 1) * 32
      MenuX_R3.Top = Y1 * 32
      MenuX_R3.Bottom = (Y2 + 1) * 32
  End Select
  
End Sub

Public Sub InitMenuSystem()
On Error GoTo ErrHan
  Dim a As String
  Dim MenuSection As Integer
  Dim TextSection As Integer
  Dim Section As Integer
  Dim StrLoc As Integer
  Dim i As Integer
  
  ReDim MenuX(0)
  ReDim MenuX(0).Text(0)
  ReDim MenuXsc(0)
  
  Open GameMain.Paths.System & "Menu.DAY" For Input As #1
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
                    ReDim Preserve MenuXsc(MenuSection)
                    MenuXsc(MenuSection).ID_Number = MenuSection
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

  
  For i1 = 0 To UBound(MenuX)
    Call frmMain.SC.AddObject("MenuX_" & i1, MenuXsc(i1))
  Next i1
  
  UpdateMenuSystem
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (InitMenuSystem) : " & Err.Description
  Err.Clear
  Close #1
End Sub


Private Sub UpdateBilly()
  Dim CharSpeed As Single
  With MapInfo.Character

  CharSpeed = (dx.TickCount - ST) / 1000 * .Speed

  If .WX = 1 And .WY = 0 Then
    If .X = .OX Then
      If AllowWalk(Int(.X + 1), Int(.Y)) Then
        Call Char_LeaveTile(.X, .Y)
        .X = .X + 0.01
        
        If Offset.X_Tile + 20 < MapInfo.MapData.Right Then  'map offset
          If .OX - Offset.X_Tile > 9 Then
            Offset.X_Tile = Offset.X_Tile + 1
            Offset.X_Pix = 0.32
          End If
        End If

      End If
      
'      Select Case .WY  'selects animation direction
'        Case 1
'          .Dir = 4
'        Case -1
'          .Dir = 5
'        Case Else
          .Dir = 1
'      End Select
    End If
  ElseIf .WX = -1 And .WY = 0 Then
    If .X = .OX Then
      If AllowWalk(Int(.X - 1), Int(.Y)) Then
        Call Char_LeaveTile(.X, .Y)
        .X = .X - 0.01
        
       If Offset.X_Tile > MapInfo.MapData.Left Then  'map offset
          If .OX - Offset.X_Tile < 11 Then
            Offset.X_Tile = Offset.X_Tile - 1
            Offset.X_Pix = 63.68
          End If
        End If

      End If
      
'      Select Case .WY  'selects animation direction
'        Case 1
'          .Dir = 7
'        Case -1
'          .Dir = 7
'        Case Else
          .Dir = 3
'      End Select
    End If
  End If
  
  If .WY = 1 And .WX = 0 Then
    If .Y = .OY Then
      If AllowWalk(Int(.X), Int(.Y + 1)) Then
        Call Char_LeaveTile(.X, .Y)
        .Y = .Y + 0.01
        
        If Offset.Y_Tile + 15 < MapInfo.MapData.Bottom Then  'map offset
          If .OY - Offset.Y_Tile > 6 Then
            Offset.Y_Tile = Offset.Y_Tile + 1
            Offset.Y_Pix = 0.32
          End If
        End If
        
      End If
      
'      Select Case .WX  'selects animation direction
'        Case 1
'          .Dir = 4
'        Case -1
'          .Dir = 7
'        Case Else
          .Dir = 0
'      End Select
    End If
  ElseIf .WY = -1 And .WX = 0 Then
    If .Y = .OY Then
      If AllowWalk(Int(.X), Int(.Y - 1)) Then
        Call Char_LeaveTile(.X, .Y)
        .Y = .Y - 0.01
        
        If Offset.Y_Tile > MapInfo.MapData.Top Then  'map offset
          If .OY - Offset.Y_Tile < 9 Then
            Offset.Y_Tile = Offset.Y_Tile - 1
            Offset.Y_Pix = 63.68
          End If
        End If
        
      End If
      
'      Select Case .WX  'selects animation direction
'        Case 1
'          .Dir = 5
'        Case -1
'          .Dir = 6
'        Case Else
          .Dir = 2
'      End Select
    End If
  End If
  
  If .X <> .OX Or .Y <> .OY Then
    .AnimEnabled = True
  Else
    .AnimEnabled = False
    .AnimFrm = 0
  End If
  
  If .OX < .X Then
    .X = .X + CharSpeed
    If .X >= .OX + 1 Then
      Offset.X_Pix = 32
      .X = .OX + 1
      .OX = .X
      Call Char_EnterTile(.X, .Y)
    End If
    
    If Offset.X_Pix <> 32 Then  'map offset
      Offset.X_Pix = Offset.X_Pix + (CharSpeed * 32)
    End If

  ElseIf .OX > .X Then
    .X = .X - CharSpeed
    If .X <= .OX - 1 Then
      Offset.X_Pix = 32
      .X = .OX - 1
      .OX = .X
      Call Char_EnterTile(.X, .Y)
    End If
    
    If Offset.X_Pix <> 32 Then  'map offset
      Offset.X_Pix = Offset.X_Pix - (CharSpeed * 32)
    End If

  End If
  
  If .OY < .Y Then
    .Y = .Y + CharSpeed
    If .Y >= .OY + 1 Then
      Offset.Y_Pix = 32
      .Y = .OY + 1
      .OY = .Y
      Call Char_EnterTile(.X, .Y)
    End If
  
    If Offset.Y_Pix <> 32 Then  'map offset
      Offset.Y_Pix = Offset.Y_Pix + (CharSpeed * 32)
    End If
  
  ElseIf .OY > .Y Then
    .Y = .Y - CharSpeed
    If .Y <= .OY - 1 Then
      Offset.Y_Pix = 32
      .Y = .OY - 1
      .OY = .Y
      Call Char_EnterTile(.X, .Y)
    End If
    
    If Offset.Y_Pix <> 32 Then  'map offset
      Offset.Y_Pix = Offset.Y_Pix - (CharSpeed * 32)
    End If
  
  
  End If
  
  ST = dx.TickCount
  End With
End Sub

Public Sub CustomClipper()
  Dim Neg         As Integer
  Dim X_Off       As Integer
  Dim Y_Off       As Integer
  
  With MapInfo.MapData.rTiles
    If Offset.X_Pix <> 32 Then
      If Offset.X_Pix > 32 Then
        Neg = 1
        X_Off = Offset.X_Pix - 32
      Else
        Neg = 0
        X_Off = Offset.X_Pix
      End If
        
      For i1 = Offset.Y_Tile + 1 To Offset.Y_Tile + 15
        For i2 = 0 To 3
          CurTile = MapInfo.MapData.Tile(Offset.X_Tile + Neg, i1, i2)
          If CurTile > 1023 Then
            .Top = 0
            .Bottom = .Top + 32
            .Left = (MapInfo.MapData.AnimFrame * 32) + MapInfo.MapData.AnimFrame2 - (32 - X_Off)
            .Right = .Left + (32 - X_Off)
            Call ddsBack.BltFast(0, (i1 - Offset.Y_Tile - 1) * 32 - Offset.Y_Pix + 32, MapInfo.MapData.ddsTiles2, MapInfo.MapData.rTiles, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT)
        
          ElseIf CurTile > 0 Then
            .Top = (CurTile \ TilesWide) * 32
            .Bottom = .Top + 32
            .Left = ((CurTile Mod TilesWide) + 1) * 32 - (32 - X_Off)
            .Right = .Left + (32 - X_Off)
            Call ddsBack.BltFast(0, (i1 - Offset.Y_Tile - 1) * 32 - Offset.Y_Pix + 32, MapInfo.MapData.ddsTiles, MapInfo.MapData.rTiles, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT)
          End If
        
        
          CurTile = MapInfo.MapData.Tile(Offset.X_Tile + 20 + Neg, i1, i2)
          If CurTile > 1023 Then
            .Top = (CurTile - 1024) * 32
            .Bottom = .Top + 32
            .Left = (MapInfo.MapData.AnimFrame * 32) + MapInfo.MapData.AnimFrame2
            .Right = .Left + X_Off
            Call ddsBack.BltFast(640 - X_Off, (i1 - Offset.Y_Tile - 1) * 32 - Offset.Y_Pix + 32, MapInfo.MapData.ddsTiles2, MapInfo.MapData.rTiles, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT)
        
          ElseIf CurTile > 0 Then
            .Top = (CurTile \ TilesWide) * 32
            .Bottom = .Top + 32
            .Left = (CurTile Mod TilesWide) * 32
            .Right = .Left + X_Off
            Call ddsBack.BltFast(640 - X_Off, (i1 - Offset.Y_Tile - 1) * 32 - Offset.Y_Pix + 32, MapInfo.MapData.ddsTiles, MapInfo.MapData.rTiles, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT)
          End If
        Next i2
      Next i1
    End If
    
    
    
    If Offset.Y_Pix <> 32 Then
      If Offset.Y_Pix > 32 Then
        Neg = 1
        Y_Off = Offset.Y_Pix - 32
      Else
        Neg = 0
        Y_Off = Offset.Y_Pix
      End If
        
      For i1 = Offset.X_Tile + 1 To Offset.X_Tile + 20
        For i2 = 0 To 3
          CurTile = MapInfo.MapData.Tile(i1, Offset.Y_Tile + Neg, i2)
          If CurTile > 1023 Then
            .Top = Y_Off
            .Bottom = .Top + (32 - Y_Off)
            .Left = (MapInfo.MapData.AnimFrame * 32) + MapInfo.MapData.AnimFrame2
            .Right = .Left + 32
            Call ddsBack.BltFast((i1 - Offset.X_Tile - 1) * 32 - Offset.X_Pix + 32, 0, MapInfo.MapData.ddsTiles2, MapInfo.MapData.rTiles, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT)
        
          ElseIf CurTile > 0 Then
            .Top = ((CurTile \ TilesWide) + 1) * 32 - (32 - Y_Off)
            .Bottom = .Top + (32 - Y_Off)
            .Left = (CurTile Mod TilesWide) * 32
            .Right = .Left + 32
            Call ddsBack.BltFast((i1 - Offset.X_Tile - 1) * 32 - Offset.X_Pix + 32, 0, MapInfo.MapData.ddsTiles, MapInfo.MapData.rTiles, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT)
          End If
        
        
          CurTile = MapInfo.MapData.Tile(i1, Offset.Y_Tile + 15 + Neg, i2)
          If CurTile > 1023 Then
            .Top = 0
            .Bottom = .Top + Y_Off
            .Left = (MapInfo.MapData.AnimFrame * 32) + MapInfo.MapData.AnimFrame2
            .Right = .Left + 32
            Call ddsBack.BltFast((i1 - Offset.X_Tile - 1) * 32 - Offset.X_Pix + 32, 480 - Y_Off, MapInfo.MapData.ddsTiles2, MapInfo.MapData.rTiles, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT)
        
          ElseIf CurTile > 0 Then
            .Top = (CurTile \ TilesWide) * 32
            .Bottom = .Top + Y_Off
            .Left = (CurTile Mod TilesWide) * 32
            .Right = .Left + 32
            Call ddsBack.BltFast((i1 - Offset.X_Tile - 1) * 32 - Offset.X_Pix + 32, 480 - Y_Off, MapInfo.MapData.ddsTiles, MapInfo.MapData.rTiles, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT)
          End If
        
        Next i2
      Next i1
    
    End If
  End With
End Sub

Public Sub DrawTiles()
On Error GoTo ErrHan
  
  With MapInfo.MapData
    For i4 = 0 To 3
      For i1 = Offset.X_Tile + 1 To Offset.X_Tile + 20
        For i2 = Offset.Y_Tile + 1 To Offset.Y_Tile + 15
          
          CurTile = .Tile(i1, i2, i4)
          If CurTile > 1023 Then
            .rTiles.Left = (.AnimFrame * 32) + .AnimFrame2
            .rTiles.Right = .rTiles.Left + 32
            .rTiles.Top = 0
            .rTiles.Bottom = .rTiles.Top + 32
          
            TileRect.Left = (i1 - Offset.X_Tile) * 32 - Offset.X_Pix
            TileRect.Top = (i2 - Offset.Y_Tile) * 32 - Offset.Y_Pix
            TileRect.Right = TileRect.Left + 32
            TileRect.Bottom = TileRect.Top + 32
            Call ddsBack.Blt(TileRect, .ddsTiles2, .rTiles, DDBLT_KEYSRC)
          ElseIf CurTile > 0 Then
            
            
            
            .rTiles.Top = (CurTile \ TilesWide) * 32
            .rTiles.Bottom = .rTiles.Top + 32
            .rTiles.Left = (CurTile Mod TilesWide) * 32
            .rTiles.Right = .rTiles.Left + 32
          
            TileRect.Left = (i1 - Offset.X_Tile) * 32 - Offset.X_Pix
            TileRect.Top = (i2 - Offset.Y_Tile) * 32 - Offset.Y_Pix
            TileRect.Right = TileRect.Left + 32
            TileRect.Bottom = TileRect.Top + 32
            Call ddsBack.Blt(TileRect, .ddsTiles, .rTiles, DDBLT_KEYSRC)
          End If
        Next i2
      Next i1
    
      If MapInfo.Character.Layer = i4 Then
        rPrim.Top = MapInfo.Character.AnimFrm * 64
        rPrim.Left = MapInfo.Character.Dir * 32
        rPrim.Bottom = rPrim.Top + 64
        rPrim.Right = rPrim.Left + 32
        'Call ddsBack.BltFast((MapInfo.Character.X - Offset.X_Tile) * 32 - Offset.X_Pix, _
                             (MapInfo.Character.Y - Offset.Y_Tile - 1) * 32 - Offset.Y_Pix, _
                              MapInfo.Character.Sprite, rPrim, DDBLTFAST_SRCCOLORKEY)
        TileRect.Left = (MapInfo.Character.X - Offset.X_Tile) * 32 - Offset.X_Pix
        TileRect.Top = (MapInfo.Character.Y - Offset.Y_Tile - 1) * 32 - Offset.Y_Pix
        TileRect.Right = TileRect.Left + 32
        TileRect.Bottom = TileRect.Top + 64
        Call ddsBack.Blt(TileRect, MapInfo.Character.Sprite, rPrim, DDBLT_KEYSRC)
      End If
    Next i4
  End With
  
Exit Sub

ErrHan:
  'Debuger.DWrite Err.Number & " (DrawTiles): " & Err.Source & " - " & Err.Description
End Sub

Public Sub UpdateDD()
  Dim ShowMenu As Boolean
  
  DD.GetCaps hwCaps, helCaps
  
  Call ddsBack.BltColorFill(rScreen, RGB(255, 255, 255))
  
'  Select Case ActiveWin
'    Case WalkWin
  If MapInfo.MapData.Visible Then
    Call UpdateBilly
    Call CustomClipper
    Call DrawTiles
  End If
  
  For i1 = 1 To UBound(MenuX)
    If MenuX(i1).Visible Then
      'UpdateMenuSystem (i1)
      ShowMenu = True
      Exit For
    End If
  Next i1
  
  If ShowMenu Then
    Call ddsBack.Blt(MenuX_R3, DDS_MenuX, MenuX_R3, DDBLT_KEYSRC)
  End If
      'If ShowFrameRate Then ddsBack.DrawText 0, 0, "FPS: " & CStr(FPS) & " - PageFlip: " & UsePageFlip, False
    
    
    
'    Case TextWin
      
'      Call DrawTiles
'      If ShowFrameRate Then ddsBack.DrawText 0, 0, "FPS: " & CStr(FPS) & " - PageFlip: " & UsePageFlip, False
      
'      Call DrawText
'      DispText.R2.Top = 0
'      DispText.R2.Left = 0
'      DispText.R2.Bottom = 70
'      DispText.R2.Right = 70
'      Call DispText.DDS_Buffer.BltFast(14, 14, CharPic(0).DDS, DispText.R2, DDBLTFAST_SRCCOLORKEY)
'      Call ddsBack.BltFast(0, 380, DispText.DDS_Buffer, DispText.R, DDBLTFAST_SRCCOLORKEY)
    
    
    
'    Case MainMenuWin
      
'      Call DrawTiles
'      If ShowFrameRate Then ddsBack.DrawText 0, 0, "FPS: " & CStr(FPS) & " - PageFlip: " & UsePageFlip, False
      
'      Call DrawMainMenu
      
      
'    Case Menu2Win
      
'      Call DrawTiles
      
'      Call DrawMainMenu
'      Call DrawMenu2
  
'      If ShowFrameRate Then ddsBack.DrawText 0, 0, "FPS: " & CStr(FPS) & " - PageFlip: " & UsePageFlip, False
'  End Select
  
  
  
  If FullScreen Then
    If ShowFrameRate Then ddsBack.DrawText 0, 0, "FPS: " & CStr(FPS) & " - PageFlip: " & UsePageFlip, False
    If UsePageFlip Then
      ddsPrimary.Flip Nothing, DDFLIP_WAIT
    Else
      Call ddsPrimary.Blt(rScreen, ddsBack, rScreen, 0)
    End If
  Else
    If ShowFrameRate Then ddsBack.DrawText 0, 0, "FPS: " & CStr(FPS), False
    Call dx.GetWindowRect(frmMain.hWnd, rPrim)
    rPrim.Top = rPrim.Top + 22
    rPrim.Left = rPrim.Left + 3
    rPrim.Bottom = rPrim.Top + rScreen.Bottom
    rPrim.Right = rPrim.Left + rScreen.Right
    Call ddsPrimary.Blt(rPrim, ddsBack, rScreen, 0)
  End If

  NumLoops = NumLoops + 1  'FPS Calculations
  If TickCount - FirstTime > 999 Then
    FPS = NumLoops
    NumLoops = 0
    FirstTime = TickCount
  End If
End Sub

Public Sub UpdateAnim()
  With MapInfo.MapData
    
    .AnimFrame = .AnimFrame + 1
    If .AnimFrame = 6 Then .AnimFrame = 0
    
    .AnimFrame2 = .AnimFrame2 + 1
    If .AnimFrame2 = 32 Then .AnimFrame2 = 0
    
  End With
End Sub

Public Sub AnimateSprite()
  MapInfo.Character.AnimFrm = MapInfo.Character.AnimFrm + 1
  If MapInfo.Character.AnimFrm = 4 Then MapInfo.Character.AnimFrm = 0
End Sub


