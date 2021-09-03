Attribute VB_Name = "Module3"
Option Explicit

Public Sub InitializeDM()
On Error Resume Next
Set DMLoader = DirectX.DirectMusicLoaderCreate
Set DMPerformance = DirectX.DirectMusicPerformanceCreate
Call DMPerformance.Init(Nothing, 0)
Call DMPerformance.SetPort(-1, 16)

Exit Sub

If Err.Number <> 0 Then
MsgBox "Unable to start DirectMusic. Check to see that your sound card is properly installed", , "DirectMusic"
EndGame
End If
End Sub

Public Function PlayMusic(FileName As String, MusicVolume As Integer)
If Right(FileName, 4) = ".mid" Then
Set DMSeg = DMLoader.LoadSegment(FileName)
Call DMSeg.SetStandardMidiFile
Call DMPerformance.SetMasterAutoDownload(True)
Call DMSeg.Download(DMPerformance)
DMPerformance.SetMasterVolume (MusicVolume * 42 - 3000)

Set DMSegState = DMPerformance.PlaySegment(DMSeg, 0, 0)
End If
End Function

Public Function SetMusicTempo(Value As Single)
If Value <= 2 Then
DMPerformance.SetMasterTempo (Value)
End If
End Function

Public Function SetMusicVolume(Value As Integer)
If Value >= 0 Then
DMPerformance.SetMasterVolume (Value * 42 - 3000)
End If
End Function

Public Function StopMusic()
If DMPerformance.IsPlaying(DMSeg, DMSegState) = True Then
Call DMPerformance.Stop(DMSeg, DMSegState, 0, 0)
End If
End Function

Public Function PauseMusic()
If DMPerformance.IsPlaying(DMSeg, DMSegState) = True Then
mtTime = DMPerformance.GetMusicTime()
StartTime = DMSegState.GetStartTime()
Call DMPerformance.Stop(DMSeg, Nothing, 0, 0)
Else
Offset = mtTime - StartTime + Offset + 1
Call DMSeg.SetStartPoint(Offset)
Set DMSegState = DMPerformance.PlaySegment(DMSeg, 0, 0)
End If
End Function

Public Function GetMusicPMTempo() As Integer
mtTime = DMPerformance.GetMusicTime()
mtTempo = DMPerformance.GetTempo(mtTime + 2000, 0)
GetMusicPMTempo = Format(mtTempo, "00.00")
End Function
