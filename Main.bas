Attribute VB_Name = "MainDK"
Option Explicit

Global dx               As New DirectX7     'The Main DirectX Object
Global hwCaps           As DDCAPS           'Hardware Emulation
Global helCaps          As DDCAPS           'SOFTWARE Emulation
Global SharedData       As New SD

Global MenuScript(4)    As New MenuClass

Global Timers(10)       As Integer
Dim TimerLen(10)        As Long

Global TickCount        As Long

Global FrmCaption       As String

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal LibraryName As String) As Long
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)

Public Sub InitDirectX()
  Call InitDD
  Call InitDI
  Call InitDS
  Call InitDM
  
  Dim CurTime As Long
  CurTime = dx.TickCount
  TimerLen(0) = CurTime
  TimerLen(1) = CurTime
  TimerLen(2) = CurTime
  TimerLen(3) = CurTime
  TimerLen(4) = CurTime
  TimerLen(5) = CurTime
  TimerLen(6) = CurTime
End Sub

Public Sub RenderLoop()
Dim LastFrameTime As Long
LastFrameTime = dx.TickCount

On Error GoTo ErrHan
  Call LoadMusic("FUNKY.MID")
  Do
    TickCount = dx.TickCount
    
    Render
    frmMain.Timer1.Enabled = True
    
    If TickCount - LastFrameTime < 5 Then
      Do
        TickCount = dx.TickCount
      Loop Until TickCount - LastFrameTime > 4
    End If
    
    LastFrameTime = dx.TickCount
  Loop Until WhatToDo = QUITME
Exit Sub

ErrHan:
  Debuger.DWrite Err.Number & " (Engine) : " & Err.Source & " : " & Err.Description
  Err.Clear
End Sub

Public Sub Render()
  Call frmMain.RunCommand("DA_FrameUpdate(1)")
  Call UpdateDD
  Call UpdateDI
  Call UpdateDS
  Call UpdateTime
  DoEvents
End Sub

Sub UpdateTime()
  Dim i As Integer
  Dim CurTime As Long
  
  CurTime = dx.TickCount
  
  For i = 0 To 6
    If CurTime - TimerLen(i) > Timers(i) Then
      TimerLen(i) = TimerLen(i) + Timers(i)

      Select Case i
        Case 0
          Call frmMain.RunCommand("DA_Timer100(1)")
          UpdateAnim
          
          If MapInfo.Character.AnimEnabled Then
            Call AnimateSprite
          End If
        Case 1
          Call frmMain.RunCommand("DA_Timer250(1)")
        
        Case 2
          Call frmMain.RunCommand("DA_Timer333(1)")
                
        Case 3
          Call frmMain.RunCommand("DA_Timer500(1)")
        
          'MenuX(3).Text(4).Text = Hour(Time) & ":" & Format(Minute(Time), "00") & ":" & Format(Second(Time), "00")
          'UpdateMenuSystem
          
        Case 4
          Call frmMain.RunCommand("DA_Timer750(1)")
        
        Case 5
          Call frmMain.RunCommand("DA_Timer1000(1)")
                
        Case 6
          'If ActiveWin = TextWin Then
          '  UpdateText
          'End If
          'i = i - 1
      End Select
    End If
  Next i
End Sub

Sub KillDX()
  Call StartKillDD
  DoEvents
  Call KillDD
  Call KillDI
  Set dx = Nothing
End Sub

Public Sub ChangeViewMode()
  FullScreen = Not FullScreen
  KillDX
  InitDirectX
  UpdateMenuSystem
End Sub
