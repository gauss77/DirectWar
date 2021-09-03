Attribute VB_Name = "Module4"
Option Explicit

Public Sub InitializeDS()

On Local Error GoTo ErrorHandler
    
Set DSound = DirectX.DirectSoundCreate("")
DSound.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL

BufferDesc.lFlags = DSBCAPS_PRIMARYBUFFER Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME

Set DSBuffer = DSound.CreateSoundBuffer(BufferDesc, WaveFormat)

BufferDesc1.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME

Set DSFire = DSound.CreateSoundBufferFromFile(App.Path & "\Data\Fire.wav", BufferDesc1, WaveFormat1)
Set DSBomb = DSound.CreateSoundBufferFromFile(App.Path & "\Data\Bomb.wav", BufferDesc1, WaveFormat1)
Set DSBomb1 = DSound.CreateSoundBufferFromFile(App.Path & "\Data\Bomb1.wav", BufferDesc1, WaveFormat1)
Set DSStart = DSound.CreateSoundBufferFromFile(App.Path & "\Data\Start.wav", BufferDesc1, WaveFormat1)
Set DSForce = DSound.CreateSoundBufferFromFile(App.Path & "\Data\Force.wav", BufferDesc1, WaveFormat1)
Set DSDestroy = DSound.CreateSoundBufferFromFile(App.Path & "\Data\Destroy.wav", BufferDesc1, WaveFormat1)

Exit Sub

ErrorHandler:

MsgBox "Unable to create Direct Sound object." & vbCrLf & "Check to ensure that a sound card is installed and working properly."
EndGame

End Sub

Public Function PlaySound(DSB As DirectSoundBuffer, SoundVolume As Long, OLoop As Boolean)

If OLoop = True Then
DSB.Play DSBPLAY_LOOPING
Else
DSB.Play DSBPLAY_DEFAULT
End If

If SoundVolume <= 0 Then
DSB.SetVolume SoundVolume
End If

End Function

Public Function PauseSound(DSB As DirectSoundBuffer)
If DSB Is Nothing Then Exit Function
DSB.Stop
End Function

Public Function StopSound(DSB As DirectSoundBuffer)
If DSB Is Nothing Then Exit Function
DSB.Stop
DSB.SetCurrentPosition 0
End Function

Public Function SetSoundVolume(DSB As DirectSoundBuffer, Value As Long)
If Value <= 0 Then
DSB.SetVolume Value
Else
DSB.SetVolume 0
End If
End Function

Public Function SetSoundPan(DSB As DirectSoundBuffer, Value As Long)
If Value >= -10000 Or Value <= 10000 Then
DSB.SetPan Value
Else
DSB.SetPan 0
End If
End Function
