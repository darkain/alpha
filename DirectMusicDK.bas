Attribute VB_Name = "DirectMusicDK"
Option Explicit

Dim perf As DirectMusicPerformance
Dim seg As DirectMusicSegment
Dim segstate As DirectMusicSegmentState
Dim loader As DirectMusicLoader
Dim PlayList(5) As String
Dim CurrentSong As Integer
Dim Playing As Boolean
Global MusicDisabled As Boolean

Public Sub InitDM()
  On Error GoTo LocalErrors
  If MusicDisabled Or Playing Then Exit Sub

  Set loader = dx.DirectMusicLoaderCreate()
  Set perf = dx.DirectMusicPerformanceCreate()
  Call perf.Init(Nothing, 0)
  perf.SetPort -1, 80
  Call perf.SetMasterAutoDownload(True)
  perf.SetMasterVolume (75 * 42 - 3000)

  PlayList(0) = "FUNKY.MID"
  PlayList(1) = "_NOMIDDL.MID"
  PlayList(2) = "!LIFE.MID"
  PlayList(3) = "_SECRET.MID"
  PlayList(4) = "1.MID"
  PlayList(5) = "7EVIL.MID"
  CurrentSong = 0
LocalErrors:
End Sub

Public Sub LoadMusic(FileName As String)
  On Error GoTo LocalErrors
  If MusicDisabled Or Playing Then Exit Sub
  Set seg = loader.LoadSegment(GameMain.Paths.Music & PlayList(CurrentSong))
  seg.SetStandardMidiFile
  Set segstate = perf.PlaySegment(seg, 0, 0)
  'Call perf.Stop(seg, segstate, 0, 0)    'stop muzix
  Playing = True
LocalErrors:
End Sub

Sub RepeatMusic()
  If MusicDisabled Then Exit Sub
  If seg Is Nothing Then Exit Sub
    
  If Not perf.IsPlaying(seg, segstate) Then
    CurrentSong = CurrentSong + 1
    If CurrentSong = 5 Then
      CurrentSong = 0
    End If
    
    Set seg = loader.LoadSegment(GameMain.Paths.Music & PlayList(CurrentSong))
    Set segstate = perf.PlaySegment(seg, 0, 0)
  End If
End Sub
