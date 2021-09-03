Attribute VB_Name = "Module6"
Option Explicit

'Declare Type
Type Fire1
X As Single
Y As Single
Exist As Boolean
End Type

Type DDPosition
X As Long
Y As Long
End Type

'Declare Variable
Public MainDC As Boolean
Public ExitDC As Long
Public ExitC As Long
Public LifeX As Single
Public SCLive As Long
Public PauseN As Long
Public EnemyX As Single
Public EnemyY As Single
Public SpriteX As Single
Public SpriteY As Single
Public DestroypX As Single
Public DestroypY As Single
Public DestroyX As Single
Public DestroyY As Single
Public DestroyX1 As Single
Public DestroyY1 As Single
Public DestroyX2 As Single
Public DestroyY2 As Single
Public Binit As Boolean
Public GPlay As Boolean
Public EnemyA As Boolean
Public EnemyExist As Boolean
Public SpriteExist As Boolean
Public EnemyA1 As Boolean
Public EnemyA2 As Boolean
Public BRestore As Boolean
Public BDestroy As Boolean
Public BDestroy1 As Boolean
Public BDestroy2 As Boolean
Public ClC As Long
Public BICo As Long
Public IntC As Long
Public ClCo As Long
Public IntCo As Long
Public EnemyAC As Long
Public EnemyACo As Long
Public EnemyLifeC As Long
Public EnemyLifeO As Long
Public ScoreC As Long
Public Clcount As Long
Public BIcount As Long
Public Intcount As Long
Public Bombn As Fire1
Public BOC As Long
Public BOCo As Long
Public BOCount As Long
Public FIC As Long
Public FICo As Long
Public FICount As Long
Public FIC1 As Long
Public FICo1 As Long
Public FICount1 As Long
Public BOC1 As Long
Public BOCo1 As Long
Public BOCount1 As Long
Public Bombn1(0 To 6) As Fire1
Public Firen(0 To 4) As Fire1
Public Firem(0 To 4) As Fire1
Public Firen1(0 To 7) As Fire1
Public BIcons(0 To 2) As Fire1
Public Cloudn(0 To 6) As Fire1
Public EnemyPX As DDPosition

'Direct X Object
Public DirectX As New DirectX7

'Direct Draw
Public DirectDraw As DirectDraw7
Public FrontSurface As DirectDrawSurface7
Public BackSurface As DirectDrawSurface7
Public LoadSurface As DirectDrawSurface7
Public BackGSurface As DirectDrawSurface7
Public SpriteSurface As DirectDrawSurface7
Public LaserSurface As DirectDrawSurface7
Public LaserSurface1 As DirectDrawSurface7
Public BombSurface As DirectDrawSurface7
Public BombSurface1 As DirectDrawSurface7
Public EnemySurface As DirectDrawSurface7
Public LifeSurface As DirectDrawSurface7
Public CloudSurface As DirectDrawSurface7
Public DestroySurface As DirectDrawSurface7
Public BIconSurface As DirectDrawSurface7
Public LoverSurface As DirectDrawSurface7
Public HitSurface As DirectDrawSurface7
Public FrontDdsd As DDSURFACEDESC2
Public BackDdsd As DDSURFACEDESC2
Public ColorKey As DDCOLORKEY
Public DDCaps As DDSCAPS2
Public LifeRect As RECT
Public EmptyRect As RECT
Public SpriteRect As RECT
Public EnemyRect As RECT
Public EnemyRectO As RECT
Public BombRect As RECT
Public DestroyRect As RECT
Public DestroyRect1 As RECT
Public DestroyRect2 As RECT
Public FireRect1(0 To 7) As RECT
Public CloudRect(0 To 6) As RECT
Public FireRect(0 To 4) As RECT
Public FireRectm(0 To 4) As RECT
Public BombRect1(0 To 6) As RECT

'Direct Input
Public DInput As DirectInput
Public DIState As DIKEYBOARDSTATE
Public DIDevice As DirectInputDevice

'Direct Sound
Public DSound As DirectSound
Public BufferDesc As DSBUFFERDESC
Public BufferDesc1 As DSBUFFERDESC
Public WaveFormat As WAVEFORMATEX
Public WaveFormat1 As WAVEFORMATEX
Public DSBuffer As DirectSoundBuffer
Public DSFire As DirectSoundBuffer
Public DSBomb As DirectSoundBuffer
Public DSBomb1 As DirectSoundBuffer
Public DSStart As DirectSoundBuffer
Public DSForce As DirectSoundBuffer
Public DSDestroy As DirectSoundBuffer

'Direct Music
Public Offset As Long
Public mtTime As Long
Public mtTempo As Long
Public StartTime As Long
Public DMSeg As DirectMusicSegment
Public DMLoader As DirectMusicLoader
Public DMSegState As DirectMusicSegmentState
Public DMPerformance As DirectMusicPerformance

'Declare Functions
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function IntersectRect Lib "user32" (ByRef r As RECT, ByRef r2 As RECT, ByRef r3 As RECT) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function PlaySoundC Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'Declare Consts
Public Const SRCCOPY = &HCC0020
Public Const ScreenWidth = 600
Public Const ScreenHeight = 440
