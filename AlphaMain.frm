VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Darkain's Angel - Alpha SiX 6"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "AlphaMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   120
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
      UseSafeSubset   =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub LoadSysVar()
  Dim a As String
  Open GameMain.Paths.Script & "SystemVar.DAC" For Input As #1
    a = Input(LOF(1), 1)
  Close 1
  
  SC.Reset
  SC.AddCode a
  SC.AddObject "Game", SharedData
End Sub

Public Function RunCommand(CmdLine As String, Optional NeedsReturn As Boolean = False)
On Error GoTo ErrHan
  Dim ReturnVal
  
  If NeedsReturn Then
    SC.ExecuteStatement ("DA_ReturnVal=" & CmdLine)
    RunCommand = SC.Eval("DA_ReturnVal")
  Else
    SC.ExecuteStatement (CmdLine)
  End If

Exit Function

ErrHan:
  If Err.Number = LastError Then Exit Function
  
  Dim ErrText As String
  LastError = Err.Number
  If frmMain.SC.Error.Number Then
    ErrText = ErrText & BreakLine & "Error: "
    ErrText = ErrText & CmdLine & " - "
    ErrText = ErrText & frmMain.SC.Error.Number & " - "
    ErrText = ErrText & frmMain.SC.Error.Line & " - "
    ErrText = ErrText & frmMain.SC.Error.Description & " - "
    ErrText = ErrText & frmMain.SC.Error.Source
  Else
    ErrText = ErrText & BreakLine & "Error: "
    ErrText = ErrText & CmdLine & " - "
    ErrText = ErrText & Err.Number & " - "
    ErrText = ErrText & Err.Description & " - "
    ErrText = ErrText & Err.Source
  End If
  frmMain.SC.Error.Clear
  Err.Clear
  Debuger.DWrite ErrText
End Function

Private Sub Form_Load()
  Caption = Caption & "  (" & App.Major & "." & App.Minor & "." & App.Revision & ")"
  FrmCaption = Caption
  
  MusicDisabled = True
  FullScreen = False
  BreakLine = Chr(13) & Chr(10)
  Me.Move Me.Left, Me.Top, 640 * 15 + (Me.Width - Me.ScaleWidth), 480 * 15 + (Me.Height - Me.ScaleHeight)
  
  'ShowCursor False
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
  Call LoadSysVar
  Call LoadControls(GameMain.Paths.System & GameMain.Header.Controls)
  Call LoadMap(GameMain.Paths.Maps & "MainMap_House.DAN")

  MapInfo.Character.OX = 5
  MapInfo.Character.OY = 5
  MapInfo.Character.X = 5
  MapInfo.Character.Y = 5
  MapInfo.Character.Layer = 1
  MapInfo.Character.Speed = 3.5
  
  Timers(0) = 100
  Timers(1) = 250
  Timers(2) = 333
  Timers(3) = 500
  Timers(4) = 750
  Timers(5) = 1000
  Timers(6) = 5

  
  'LoadItems
  
  
  Me.Show
  Debuger.Show , Me
  DebugStats.Show , Me
  Me.SetFocus
  
  InitDirectX
  InitMenuSystem
  Call RunCommand("DA_InitGame")
  
  RenderLoop
  'ShowCursor True
  End
End Sub

Private Sub Form_Paint()
Render
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  End
End Sub

'Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyEscape Then WhatToDo = QUITME
'  If (KeyCode = vbKey1) And (FullScreen = True) Then UsePageFlip = Not UsePageFlip
'  If KeyCode = vbKey2 Then ShowFrameRate = Not ShowFrameRate
'  If KeyCode = vbKeyReturn Then
'    If (Shift And vbAltMask) > 0 Then
'      ChangeViewMode
'    End If
'  End If
'
'  If KeyCode = vbKeyS Then
'    DebugStats.Visible = Not DebugStats.Visible
'    frmMain.SetFocus
'  End If
'  If KeyCode = vbKeyD Then
'    Debuger.Visible = Not Debuger.Visible
'    frmMain.SetFocus
'  End If
'
'
'  If KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
'    ButtonPressed (KeyCode - 96)
'  End If
'
'
'  If KeyCode = vbKeyJ Then
'    Debuger.DWrite "ACTION = " & ButtonAction
'    Debuger.DWrite "MENU = " & ButtonMenu
'    Debuger.DWrite "CANCEL = " & ButtonCancel
'    Debuger.DWrite "OK = " & ButtonOK
'
'  End If
'End Sub

'Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
'    X_Center 1
'  End If
'  If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
'    Y_Center 1
'  End If
'
'  If KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
'    ButtonUnPressed (KeyCode - 96)
'  End If
'
'End Sub
Private Sub Timer1_Timer()
  Render
End Sub
