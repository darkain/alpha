Attribute VB_Name = "Resource"
Option Explicit

Private Type MapHeaderType
  Map As String
  Title As String
  CodeScript As String
  FontSet As String
  TileSet As String
  Pallet As String
End Type

Private Type MapFileType
  Top As Integer
  Bottom As Integer
  Left As Integer
  Right As Integer
  Layers As Byte
  Flag()         As Byte                   'Flag Map
  Walk()         As Byte                   'Walkable locations
  Tile()         As Integer                'Tile locations
End Type

Private Type CharacterType
  DefX As Long
  DefY As Long
  DefLayer As Byte
  DefDir As Byte
End Type

Private Type TeleportType
  Transition As Byte
  srcLayerMin As Byte
  srcLayerMax As Byte
  SrcX1  As Long
  SrcY1 As Long
  SrcX2 As Long
  SrcY2 As Long
  DestX As Long
  DestY As Long
  DestDir As Byte
  DestLayer As Byte
  DestMap As String
End Type

Private Type DAM_InfoType
  MapHeader As MapHeaderType
  MapData As MapFileType
  Character As CharacterType
  Teleport() As TeleportType
  Telepoters As Integer
End Type
Public MapInfo As DAM_InfoType   'Data Dec

Global Const Flag1 = 1
Global Const Flag2 = 2
Global Const Flag3 = 4
Global Const Flag4 = 8
Global Const Flag5 = 16
Global Const Flag6 = 32
Global Const Flag7 = 64
Global Const Flag8 = 128




Private Type GameMainHeaderType
  Game           As String
  Engine         As String
  Script         As String
End Type

Private Type GameMainPathsType
  Root           As String
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




Global Const SRCAND = &H8800C6       ' (DWORD) dest = source AND dest
Global Const SRCCOPY = &HCC0020      ' (DWORD) dest = source
Global Const SRCERASE = &H440328     ' (DWORD) dest = source AND (NOT dest )
Global Const SRCINVERT = &H660046    ' (DWORD) dest = source XOR dest
Global Const SRCPAINT = &HEE0086     ' (DWORD) dest = source OR dest
Global Const NOTSRCCOPY = &H330008   ' (DWORD) dest = (NOT source)

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type
'Constant Declaration
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0

Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type

Global lpPrevWndProc As Long
Global gHW As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)

Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long


'draw a line from point X1,Y1 to X2,Y2 with specific atributes
Public Sub Lines(ByVal hDc As Long, _
                 ByVal X1 As Long, ByVal Y1 As Long, _
                 ByVal X2 As Long, ByVal Y2 As Long, _
                 Optional ByVal crColor As Long = 0, _
                 Optional ByVal PenWidth As Integer = 1, _
                 Optional ByVal PenStyle As Integer = 0)

  Dim MTPnts As POINTAPI
  Dim Pen As Long
  Dim OldPen As Long

  Pen = CreatePen(PenStyle, PenWidth, crColor)
  OldPen = SelectObject(hDc, Pen)
  MoveToEx hDc, X1, Y1, MTPnts
  LineTo hDc, X2, Y2

  SelectObject hDc, OldPen
  OldPen = DeleteObject(Pen)
End Sub

Public Function CreateMaskImage(ByRef picFrom As PictureBox, ByRef picTo As PictureBox, Optional ByVal lTransparentColor As Long = -1) As Boolean
  Dim lhDC As Long
  Dim lhBmp As Long
  Dim lhBmpOld As Long
  
  With picTo
    .BorderStyle = 0
    .Width = picFrom.Width
    .Height = picFrom.Height
    .Cls
  End With

  lhDC = CreateCompatibleDC(0)
  If (lhDC <> 0) Then
    lhBmp = CreateCompatibleBitmap(lhDC, picFrom.ScaleWidth, picFrom.ScaleHeight)
    If (lhBmp <> 0) Then
      lhBmpOld = SelectObject(lhDC, lhBmp)
      If (lTransparentColor = -1) Then lTransparentColor = picFrom.BackColor
      SetBkColor lhDC, lTransparentColor
      BitBlt lhDC, 0, 0, picFrom.ScaleWidth, picFrom.ScaleHeight, picFrom.hDc, 0, 0, SRCCOPY
      BitBlt picTo.hDc, 0, 0, picFrom.ScaleWidth, picFrom.ScaleHeight, lhDC, 0, 0, SRCAND
      
      SelectObject lhDC, lhBmpOld
      DeleteObject lhBmp
    End If
    DeleteDC lhDC
  End If
End Function

Public Function FileExist(FileName As String) As Boolean
  On Error Resume Next
  Open FileName For Input As #2
  If Err.Number > 0 Then
    Close #2
    Err.Clear
    FileExist = False
  Else
    Close #2
    FileExist = True
  End If
End Function

Sub ChangeDirectory(ByVal FileName As String)
  Dim i As Integer
  
  For i = Len(FileName) To 1 Step -1
    If Mid(FileName, i, 1) = "\" Then
      Exit For
    End If
  Next i

  FileName = Left(FileName, i)
  ChDir FileName
End Sub


Public Sub Hook()
  lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
  Dim Tmp As Long
  Tmp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim MinMax As MINMAXINFO
  If uMsg = WM_GETMINMAXINFO Then
    CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)
    MinMax.ptMinTrackSize.X = 546
    MinMax.ptMinTrackSize.Y = 310
    'MinMax.ptMaxTrackSize.x = 500
    'MinMax.ptMaxTrackSize.y = 500
    CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)
    WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
  Else
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
  End If
End Function
