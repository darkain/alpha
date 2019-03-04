Attribute VB_Name = "EventsDK"
Sub Char_EnterTile(X As Single, Y As Single)
  Dim i As Integer
  Dim MapToLoad As String
  
  If MapInfo.Telepoters > 0 Then
    For i = 1 To MapInfo.Telepoters
        
        If X >= MapInfo.Teleport(i).SrcX1 Then
          If X <= MapInfo.Teleport(i).SrcX2 Then
            If Y >= MapInfo.Teleport(i).SrcY1 Then
              If Y <= MapInfo.Teleport(i).SrcY2 Then
                
                Dim CharLayer As String, CharX As String
                Dim CharY As Single, CharDir As String
                
                CharLayer = MapInfo.Teleport(i).DestLayer
                CharX = MapInfo.Teleport(i).DestX
                CharY = MapInfo.Teleport(i).DestY
                CharDir = MapInfo.Teleport(i).DestDir
                
                LoadMap GameMain.Paths.Maps & MapInfo.Teleport(i).DestMap
                
                MapInfo.Character.Layer = CharLayer
                MapInfo.Character.X = CharX
                MapInfo.Character.Y = CharY
                MapInfo.Character.OX = CharX
                MapInfo.Character.OY = CharY
                MapInfo.Character.Dir = CharDir
                Exit For
                
              End If
            End If
          End If
        End If
      
    Next i
  End If

End Sub

Sub Char_LeaveTile(X As Single, Y As Single)

End Sub
