VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Get Tile(ByVal X As Integer, ByVal Y As Integer, ByVal Layer As Byte) As Integer
  Tile = MapInfo.MapData.Tile(X + 1, Y + 1, Layer + 1)
End Property

Property Let Tile(ByVal X As Integer, ByVal Y As Integer, ByVal Layer As Byte, Value As Integer)
  MapInfo.MapData.Tile(X + 1, Y + 1, Layer + 1) = Value
End Property

Property Get Flag(ByVal X As Integer, ByVal Y As Integer, ByVal Layer As Byte) As Byte
  Flag = MapInfo.MapData.Flag(X + 1, Y + 1, Layer + 1)
End Property

Property Let Flag(ByVal X As Integer, ByVal Y As Integer, ByVal Layer As Byte, Value As Byte)
  MapInfo.MapData.Flag(X + 1, Y + 1, Layer + 1) = Value
End Property

Property Get Walk(ByVal X As Integer, ByVal Y As Integer, ByVal Layer As Byte) As Byte
  Walk = MapInfo.MapData.Walk(X + 1, Y + 1, Layer + 1)
End Property

Property Let Walk(ByVal X As Integer, ByVal Y As Integer, ByVal Layer As Byte, Value As Byte)
  MapInfo.MapData.Walk(X + 1, Y + 1, Layer + 1) = Value
End Property


Public Sub SetActiveWin(ByVal NewWin As Integer)
  ActiveWin = NewWin
  frmMain.RunCommand ("ActiveWin=" & ActiveWin)
End Sub

Public Sub ExitEngine()
  WhatToDo = QUITME
End Sub


