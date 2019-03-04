Attribute VB_Name = "DirectInputDK"
Option Explicit

Dim DI                  As DirectInput     'the main DirectInput object

Dim diDEV_KB            As DirectInputDevice
Dim diState_KB          As DIKEYBOARDSTATE

Dim diDev               As DirectInputDevice
Dim diDevEnum           As DirectInputEnumDevices
Dim EventHandle         As Long
Dim joyCaps             As DIDEVCAPS
Dim js                  As DIJOYSTATE
Dim DiProp_Dead         As DIPROPLONG
Dim DiProp_Range        As DIPROPRANGE
Dim DiProp_Saturation   As DIPROPLONG
Dim AxisPresent(1 To 8) As Boolean
Dim ButtonDown(8)       As Integer
Dim XDown               As Integer
Dim YDown               As Integer

Global ButtonOK         As Integer
Global ButtonCancel     As Integer
Global ButtonMenu       As Integer
Global ButtonAction     As Integer
Dim Buttons             As Integer

Dim LastLeft            As Boolean
Dim LastRight           As Boolean
Dim LastUp              As Boolean
Dim LastDown            As Boolean

Public Sub CenterDI()
  LastLeft = False
  LastRight = False
  LastUp = False
  LastDown = False
  XDown = 5000
  YDown = 5000
End Sub

Public Sub KillDI()
  diDEV_KB.Unacquire
  diDev.Unacquire
End Sub

Public Sub InitDI()
On Local Error Resume Next
  Set DI = dx.DirectInputCreate()
  
  'DI-Keyboard
  Set diDEV_KB = DI.CreateDevice("GUID_SysKeyboard")
  diDEV_KB.SetCommonDataFormat DIFORMAT_KEYBOARD
  diDEV_KB.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  diDEV_KB.Acquire
  
  
  'DI-JoyStick
  Set diDevEnum = DI.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
  If diDevEnum.GetCount = 0 Then
    Exit Sub
  End If
  EventHandle = dx.CreateEvent(frmMain)
  
  Set diDev = Nothing
  Set diDev = DI.CreateDevice(diDevEnum.GetItem(1).GetGuidInstance)
  diDev.SetCommonDataFormat DIFORMAT_JOYSTICK
  diDev.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  diDev.GetCapabilities joyCaps
  'Call IdentifyAxes(diDev)
  Buttons = joyCaps.lButtons
  Call diDev.SetEventNotification(EventHandle)

  With DiProp_Dead
    .lData = 1000
    .lObj = DIJOFS_X
    .lSize = Len(DiProp_Dead)
    .lHow = DIPH_BYOFFSET
    .lObj = DIJOFS_X
    diDev.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    .lObj = DIJOFS_Y
    diDev.SetProperty "DIPROP_DEADZONE", DiProp_Dead
  End With

  With DiProp_Saturation
    .lData = 9500
    .lHow = DIPH_BYOFFSET
    .lSize = Len(DiProp_Saturation)
    .lObj = DIJOFS_X
    diDev.SetProperty "DIPROP_SATURATION", DiProp_Saturation
    .lObj = DIJOFS_Y
    diDev.SetProperty "DIPROP_SATURATION", DiProp_Saturation
  End With
    
  With DiProp_Range
    .lHow = DIPH_DEVICE
    .lSize = Len(DiProp_Range)
    .lMin = 0
    .lMax = 10000
  End With
  diDev.SetProperty "DIPROP_RANGE", DiProp_Range

  diDev.Acquire
End Sub

Public Sub UpdateDI()
On Local Error Resume Next
  Dim i1 As Integer
  
  diDEV_KB.GetDeviceStateKeyboard diState_KB
  If diState_KB.KEY(DIK_LEFT) <> 0 Then
    If Not LastLeft Then
      Call X_Left(0)
      LastLeft = True
    End If
  ElseIf diState_KB.KEY(DIK_RIGHT) <> 0 Then
    If Not LastRight Then
      Call X_Right(0)
      LastRight = True
    End If
  Else
    If LastLeft Then
      Call X_Center(0)
      LastLeft = False
    End If
    If LastRight Then
      Call X_Center(0)
      LastRight = False
    End If
  End If
  
  If diState_KB.KEY(DIK_UP) <> 0 Then
    If Not LastUp Then
      Call Y_Up(0)
      LastUp = True
    End If
  ElseIf diState_KB.KEY(DIK_DOWN) <> 0 Then
    If Not LastDown Then
      Call Y_Down(0)
      LastDown = True
    End If
  Else
    If LastUp Then
      Call Y_Center(0)
      LastUp = False
    End If
    If LastDown Then
      Call Y_Center(0)
      LastDown = False
    End If
  End If
  
  
  If diDev Is Nothing Then Exit Sub
  Dim i As Integer
  
  diDev.GetDeviceStateJoystick js
  diDev.Poll
  If Err.Number = DIERR_NOTACQUIRED Or Err.Number = DIERR_INPUTLOST Then
    diDev.Acquire
    Err.Clear
  End If

  If js.X < 7500 And js.X > 2500 Then
    If XDown > 7500 Or XDown < 2500 Then
      XDown = js.X
      X_Center (XDown)
    End If
  ElseIf js.X > 7500 Then
    If XDown < 7500 Then
      XDown = js.X
      X_Right (XDown)
    End If
  ElseIf js.X < 2500 Then
    If XDown > 2500 Then
      XDown = js.X
      X_Left (XDown)
    End If
  End If

  If js.Y < 7500 And js.Y > 2500 Then
    If YDown > 7500 Or YDown < 2500 Then
      YDown = js.Y
      Y_Center (YDown)
    End If
  ElseIf js.Y > 7500 Then
    If YDown < 7500 Then
      YDown = js.Y
      Y_Down (YDown)
    End If
  ElseIf js.Y < 2500 Then
    If YDown > 2500 Then
      YDown = js.Y
      Y_Up (YDown)
    End If
  End If

  For i = 0 To Buttons - 1
    If js.Buttons(i) <> ButtonDown(i) Then
      ButtonDown(i) = js.Buttons(i)
      If ButtonDown(i) Then
        ButtonPressed (i)
      Else
        ButtonUnPressed (i)
      End If
    End If
  Next i

End Sub

Public Sub ButtonPressed(ByVal ButtonNum As Byte)
On Error GoTo ErrHan
  
  If Not frmMain.RunCommand("ButtonPressed(" & ButtonNum & ")", True) Then
  
    Dim i As Integer
    Select Case ActiveWin
      Case WalkWin
      
        Select Case ButtonNum
          Case ButtonAction
            For i = 0 To 9
              If IsFacing(GameObject(i).X1, GameObject(i).Y1) Then
                'AddText GameObject(i).Property1
                'RedrawTextBox
              End If
            Next i
        
          Case ButtonMenu
            If MapInfo.Character.X = MapInfo.Character.OX Then
              If MapInfo.Character.Y = MapInfo.Character.OY Then
              
                ActiveWin = MenuWin
                frmMain.RunCommand ("ActiveWin=" & MenuWin)
                'CreateMenu
                'MenuX(2).Visible = True
                'MenuX(3).Visible = True
                UpdateMenuSystem
              End If
            End If
        End Select
  
    End Select
    
  End If
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (ButtonPressed) : " & Err.Description
End Sub

Public Sub ButtonUnPressed(ByVal ButtonNum As Byte)
On Error GoTo ErrHan
  
  If Not frmMain.RunCommand("ButtonUnPressed(" & ButtonNum & ")", True) Then
    Select Case ButtonNum
      'Case 0
      '  If ActiveWin = TextWin Then
      '    DispText.Speed = 1
      '  End If
    End Select
  End If
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (ButtonUnPressed) : " & Err.Description
End Sub

Public Sub X_Left(Pos As Integer)
On Error GoTo ErrHan
  
  If Not frmMain.RunCommand("DA_X_Neg(1)", True) Then
    Select Case ActiveWin
      Case WalkWin
        If MapInfo.Character.WY = 0 Then
          MapInfo.Character.WX = -1
        End If
    End Select
  End If
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (X_Neg) : " & Err.Description
End Sub

Public Sub X_Right(Pos As Integer)
On Error GoTo ErrHan
  
  If Not frmMain.RunCommand("DA_X_Pos(1)", True) Then
    Select Case ActiveWin
      Case WalkWin
        If MapInfo.Character.WY = 0 Then
          MapInfo.Character.WX = 1
        End If
    End Select
  End If
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (X_Pos) : " & Err.Description
End Sub

Public Sub X_Center(Pos As Integer)
On Error GoTo ErrHan
  
  If Not frmMain.RunCommand("DA_X_Center(1)", True) Then
    Select Case ActiveWin
      Case WalkWin
        MapInfo.Character.WX = 0
    End Select
  End If
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (X_Center) : " & Err.Description
End Sub

Public Sub Y_Up(Pos As Integer)
On Error GoTo ErrHan
  
  If Not frmMain.RunCommand("DA_Y_Neg(1)", True) Then
    Select Case ActiveWin
      Case WalkWin
        If MapInfo.Character.WX = 0 Then
          MapInfo.Character.WY = -1
        End If
      
      'Case MainMenuWin
      '  If DispMenu.CurSlot = 0 Then
      '    DispMenu.CurSlot = 15
      '  Else
      '    DispMenu.CurSlot = DispMenu.CurSlot - 1
      '  End If
    End Select
  End If
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (Y_Neg) : " & Err.Description
End Sub

Public Sub Y_Down(Pos As Integer)
On Error GoTo ErrHan
  
  If Not frmMain.RunCommand("DA_Y_Pos(1)", True) Then
    Select Case ActiveWin
      Case WalkWin
        If MapInfo.Character.WX = 0 Then
          MapInfo.Character.WY = 1
        End If
      
      'Case MainMenuWin
      '  If DispMenu.CurSlot = 15 Then
      '    DispMenu.CurSlot = 0
      '  Else
      '    DispMenu.CurSlot = DispMenu.CurSlot + 1
      '  End If
    End Select
  End If
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (Y_Pos) : " & Err.Description
End Sub

Public Sub Y_Center(Pos As Integer)
On Error GoTo ErrHan

  If Not frmMain.RunCommand("DA_Y_Center(1)", True) Then
    Select Case ActiveWin
      Case WalkWin
        MapInfo.Character.WY = 0
    End Select
  End If
  
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (Y_Center) : " & Err.Description
End Sub

Function IsFacing(ByRef X As Integer, ByRef Y As Integer) As Boolean
  IsFacing = False
  
  With MapInfo.Character
    If Int(.X) <> .X Or Int(.Y) <> .Y Then Exit Function
    
    Select Case .Dir
      Case 0
        If Int(.X) - 1 = X Then
          If Int(.Y) = Y Then
            IsFacing = True
          End If
        End If
      Case 1
        If Int(.X) = X Then
          If Int(.Y) - 1 = Y Then
            IsFacing = True
          End If
        End If
      Case 2
        If Int(.X) = X Then
          If Int(.Y) - 1 = Y Then
            IsFacing = True
          End If
        End If
      Case 3
        If Int(.X) - 2 = X Then
          If Int(.Y) - 1 = Y Then
            IsFacing = True
          End If
        End If
    End Select
  End With
End Function
