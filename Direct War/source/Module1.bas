Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()

On Error GoTo ErrorHandler

Load Form1

InitializeStartup

Do Until Binit = False

UpdateControlFunction

If MainDC = True Then

UpdateControl
UpdateBackground
UpdateLife
UpdateEnemy
UpdateSprite
UpdateFire
CheckCollisions
End If

If EnemyExist = True Then
Else
If ExitDC = 1000 Then
Binit = False
Else
BackSurface.DrawText (ScreenWidth / 2) - 100, ScreenHeight / 2, "Congratulation you win !", False
BackSurface.DrawText (ScreenWidth / 2) - 50, (ScreenHeight / 2) + 15, "Your Score Is : " & ScoreC, False
ExitDC = ExitDC + 1
End If
End If

If SpriteExist = True Then
Else
If ExitDC = 500 Then
Binit = False
Else
BackSurface.BltFast ScreenWidth / 2 - 70, ScreenHeight / 2, LoverSurface, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
ExitDC = ExitDC + 1
End If
End If

If GPlay = True Then
Else
BackSurface.DrawText ScreenWidth / 2, ScreenHeight / 2, "Pause", False
End If

FrontSurface.Flip Nothing, 0

Loop

GPlay = False
SpriteExist = False
EnemyExist = False
EndGame

ErrorHandler:

If Err.Number = DDERR_NOEXCLUSIVEMODE Then
EndGame
ElseIf Err.Number = DDERR_SURFACELOST Then
EndGame
End If

If Err Then
EndGame
End If

End Sub

Public Sub InitializeStartup()

DoEvents

Form1.MouseIcon = LoadPicture(App.Path & "\Data\Mouse.cur")
Form1.MousePointer = 99

Set LoadSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Load.gif")
BackSurface.BltFast 0, 0, LoadSurface, EmptyRect, DDBLTFAST_WAIT
FrontSurface.Flip Nothing, 0
Set LoadSurface = Nothing

Call InitializeDI
Call InitializeDS
Call InitializeDM

Binit = True

MainDC = True

GPlay = True
SpriteExist = True
EnemyExist = True

PlayMusic App.Path & "\Data\Music1.mid", 80

PlaySound DSStart, -300, False

SpriteX = 305
SpriteY = ScreenHeight - 20
EnemyX = 190
EnemyY = 50

LifeX = 0
ExitC = 0
ExitDC = 0
SCLive = 3
EnemyACo = 5
If EnemyLifeO = 0 Then
EnemyLifeO = 1000
EnemyLifeC = EnemyLifeO
Else
EnemyLifeC = EnemyLifeO
End If

Exit Sub

End Sub

Public Sub UpdateBackground()
BackSurface.BltFast 0, 0, BackGSurface, EmptyRect, DDBLTFAST_WAIT

'Update Clouds
If ClC = 10 Then
ClC = 0
If Clcount = 6 Then
Clcount = 0
Else
Clcount = Clcount + 1
End If
Else
ClC = ClC + 1
End If

If Cloudn(Clcount).Exist = False Then
Cloudn(Clcount).X = 500 * Rnd + Clcount
Cloudn(Clcount).Y = 50 * Rnd + Clcount
Cloudn(Clcount).Exist = True
End If

Do Until ClCo = 6
ClCo = ClCo + 1
If Cloudn(ClCo).Exist = True Then
CloudRect(ClCo).Left = Cloudn(ClCo).X
CloudRect(ClCo).Right = Cloudn(ClCo).X + 124
CloudRect(ClCo).Top = Cloudn(ClCo).Y
CloudRect(ClCo).Bottom = Cloudn(ClCo).Y + 130
If DetectCollision(SpriteRect, CloudRect(ClCo)) = True Then
Cloudn(ClCo).Exist = False
PlaySound DSForce, -300, False
LifeX = LifeX - 5
End If
Cloudn(ClCo).Y = Cloudn(ClCo).Y + 0.5
BackSurface.BltFast Cloudn(ClCo).X, Cloudn(ClCo).Y, CloudSurface, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

If Cloudn(ClCo).Y > 420 Then
Cloudn(ClCo).Exist = False
End If
Loop
ClCo = 0
'End Update

'Update Firen
If FIC = 20 Then
FIC = 0
If FICount = 4 Then
FICount = 0
Else
FICount = FICount + 1
End If
Else
FIC = FIC + 1
End If

If Firen(FICount).Exist = False Then
Firen(FICount).X = EnemyX + 60
Firen(FICount).Y = EnemyY + 70
Firen(FICount).Exist = True
End If

Do Until FICo = 4
FICo = FICo + 1
If Firen(FICo).Exist = True Then
FireRect(FICo).Left = Firen(FICo).X
FireRect(FICo).Right = Firen(FICo).X + 19
FireRect(FICo).Top = Firen(FICo).Y
FireRect(FICo).Bottom = Firen(FICo).Y + 19
If DetectCollision(SpriteRect, FireRect(FICo)) = True Then
Firen(FICo).Exist = False
PlaySound DSForce, -300, False
SCLive = SCLive - 1
BDestroy1 = True
LifeX = 0
RBIcon
End If
Firen(FICo).Y = Firen(FICo).Y + 2
BackSurface.BltFast Firen(FICo).X, Firen(FICo).Y, LaserSurface, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

If Firen(FICo).Y > 420 Then
Firen(FICo).Exist = False
End If
Loop
FICo = 0
'End Firen

'Update Firem
If FIC1 = 20 Then
FIC1 = 0
If FICount1 = 4 Then
FICount1 = 0
Else
FICount1 = FICount1 + 1
End If
Else
FIC1 = FIC1 + 1
End If

If Firem(FICount1).Exist = False Then
Firem(FICount1).X = EnemyX + 190
Firem(FICount1).Y = EnemyY + 70
Firem(FICount1).Exist = True
End If

Do Until FICo1 = 4
FICo1 = FICo1 + 1
If Firem(FICo1).Exist = True Then
FireRectm(FICo1).Left = Firem(FICo1).X
FireRectm(FICo1).Right = Firem(FICo1).X + 19
FireRectm(FICo1).Top = Firem(FICo1).Y
FireRectm(FICo1).Bottom = Firem(FICo1).Y + 19
If DetectCollision(SpriteRect, FireRectm(FICo1)) = True Then
Firem(FICo1).Exist = False
PlaySound DSForce, -300, False
SCLive = SCLive - 1
BDestroy1 = True
LifeX = 0
RBIcon
End If
Firem(FICo1).Y = Firem(FICo1).Y + 2
BackSurface.BltFast Firem(FICo1).X, Firem(FICo1).Y, LaserSurface, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

If Firem(FICo1).Y > 420 Then
Firem(FICo1).Exist = False
End If
Loop
FICo1 = 0
'End Firem


End Sub

Public Sub UpdateSprite()
BackSurface.BltFast EnemyX, EnemyY, EnemySurface, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
BackSurface.BltFast SpriteX, SpriteY, SpriteSurface, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

If DMPerformance.IsPlaying(DMSeg, DMSegState) = False Then
PlayMusic App.Path & "\Data\Music1.mid", 80
End If

End Sub

Public Sub UpdateLife()
LifeRect.Left = 0
LifeRect.Right = 200 + LifeX
LifeRect.Top = 0
LifeRect.Bottom = 18
BackSurface.BltFast 420, 440, LifeSurface, LifeRect, DDBLTFAST_WAIT
BackSurface.DrawText 10, 10, "Boss Life :", False
BackSurface.DrawText 85, 10, EnemyLifeC, False
BackSurface.DrawText 520, 10, "Score :", False
BackSurface.DrawText 570, 10, ScoreC, False
BackSurface.DrawText 360, 440, "Lives :", False
BackSurface.DrawText 400, 440, SCLive, False
Do Until BICo = BIcount
If BIcons(BICo).Exist = True Then
BackSurface.BltFast BIcons(BICo).X, BIcons(BICo).Y, BIconSurface, EmptyRect, DDBLTFAST_WAIT
End If
BICo = BICo + 1
Loop
BICo = 0

If ClC = 10 Then
ClC = 0
If Clcount = 6 Then
Clcount = 0
Else
Clcount = Clcount + 1
End If
Else
ClC = ClC + 1
End If

If Cloudn(Clcount).Exist = False Then
Cloudn(Clcount).X = 500 * Rnd + Clcount
Cloudn(Clcount).Y = 50 * Rnd + Clcount
Cloudn(Clcount).Exist = True
End If

End Sub

Public Sub UpdateEnemy()
If EnemyA = True Then
If EnemyA1 = True Then
If EnemyX < 340 Then
EnemyX = EnemyX + 1
End If
If EnemyY < 90 Then
EnemyY = EnemyY + 0.3
End If
End If
If EnemyX >= 340 And EnemyY >= 90 Then
EnemyA1 = False
If EnemyAC > EnemyACo * Rnd Then
EnemyAC = 0
EnemyA = False
Else
EnemyAC = EnemyAC + 1
End If
End If
If EnemyA1 = False Then
If EnemyX > 20 Then
EnemyX = EnemyX - 1
End If
If EnemyY > 20 Then
EnemyY = EnemyY - 0.3
End If
End If
If EnemyX <= 20 And EnemyY <= 20 Then
EnemyA1 = True
End If
End If
If EnemyA = False Then
If EnemyA2 = False Then
If EnemyX < 450 Then
EnemyX = EnemyX + 1
End If
If EnemyY > -1 Then
EnemyY = EnemyY - 0.3
End If
End If
If EnemyX >= 450 And EnemyY <= -1 Then
EnemyA2 = True
End If
If EnemyA2 = True Then
If EnemyX > 20 Then
EnemyX = EnemyX - 1
End If
End If
If EnemyX <= 20 Then
EnemyA2 = False
EnemyA = True
End If
End If
End Sub

Public Sub UpdateFire()

'Update Sprite Fire
Do Until IntCo = 7
IntCo = IntCo + 1
If Firen1(IntCo).Exist = True Then
Firen1(IntCo).Y = Firen1(IntCo).Y - 7
'Detect Collision
FireRect1(IntCo).Left = Firen1(IntCo).X
FireRect1(IntCo).Right = Firen1(IntCo).X + 10
FireRect1(IntCo).Top = Firen1(IntCo).Y
FireRect1(IntCo).Bottom = Firen1(IntCo).Y + 11
If DetectCollision(EnemyRect, FireRect1(IntCo)) = True Then
If EnemyRect.Left < 0 Or EnemyRect.Left > 450 Or EnemyRect.Top < 0 Then
Else
Firen1(IntCo).Exist = False
EnemyLifeC = EnemyLifeC - 1
ScoreC = ScoreC + 20
BackSurface.BltFast Firen1(IntCo).X, Firen1(IntCo).Y - 80 * Rnd, HitSurface, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If
End If
'End Detect
BackSurface.BltFast Firen1(IntCo).X, Firen1(IntCo).Y, LaserSurface1, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

If Firen1(IntCo).Y < 0 Then
Firen1(IntCo).Exist = False
End If
Loop
IntCo = 0
'End Update

'Update Sprite Bomb
If Bombn.Exist = True Then
Bombn.Y = Bombn.Y - 1
BackSurface.BltFast Bombn.X, Bombn.Y, BombSurface, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If
If Bombn.Y < 0 Then
Bombn.Exist = False
StopSound DSBomb
End If
'End Update

If EnemyRect.Left < 0 Or EnemyRect.Left > 450 Or EnemyRect.Top < 0 Then

'Update Bombn1
If BOC1 = 20 Then
BOC1 = 0
If BOCount1 = 6 Then
BOCount1 = 0
Else
BOCount1 = BOCount1 + 1
End If
Else
BOC1 = BOC1 + 1
End If

If Bombn1(BOCount1).Exist = False Then
Bombn1(BOCount1).X = 500 * Rnd + BOCount
Bombn1(BOCount1).Y = 50 * Rnd + BOCount
Bombn1(BOCount1).Exist = True
PlaySound DSBomb1, -500, False
End If

Do Until BOCo1 = 6
BOCo1 = BOCo1 + 1
If Bombn1(BOCo1).Exist = True Then
BombRect1(BOCo1).Left = Bombn1(BOCo1).X
BombRect1(BOCo1).Right = Bombn1(BOCo1).X + 47
BombRect1(BOCo1).Top = Bombn1(BOCo1).Y
BombRect1(BOCo1).Bottom = Bombn1(BOCo1).Y + 37
If DetectCollision(SpriteRect, BombRect1(BOCo1)) = True Then
Bombn1(BOCo1).Exist = False
PlaySound DSForce, -300, False
SCLive = SCLive - 1
BDestroy1 = True
LifeX = 0
End If
Bombn1(BOCo1).Y = Bombn1(BOCo1).Y + 1
BackSurface.BltFast Bombn1(BOCo1).X, Bombn1(BOCo1).Y, BombSurface1, EmptyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

If Bombn1(BOCo1).Y > 420 Then
Bombn1(BOCo1).Exist = False
End If
Loop
BOCo1 = 0
'End Update

End If

End Sub

Public Sub CheckCollisions()
EnemyRect.Left = EnemyX
EnemyRect.Right = EnemyX + 252
EnemyRect.Top = EnemyY
EnemyRect.Bottom = EnemyY + 155

SpriteRect.Left = SpriteX
SpriteRect.Right = SpriteX + 23
SpriteRect.Top = SpriteY
SpriteRect.Bottom = SpriteY + 29

BombRect.Left = Bombn.X
BombRect.Right = Bombn.X + 15
BombRect.Top = Bombn.Y
BombRect.Bottom = Bombn.Y + 55

If DetectCollision(EnemyRect, BombRect) = True Then
If EnemyRect.Left < 0 Or EnemyRect.Left > 450 Or EnemyRect.Top < 0 Then
Else
If Bombn.Exist = True Then
EnemyLifeC = EnemyLifeC - 100
ScoreC = ScoreC + 200
BDestroy = True
PlaySound DSDestroy, -400, False
StopSound DSBomb
End If
Bombn.Exist = False
End If
End If

If DetectCollision(EnemyRect, SpriteRect) = True Then
If EnemyRect.Left < 0 Or EnemyRect.Left > 450 Or EnemyRect.Top < 0 Then
Else
LifeX = LifeX - 1
End If
End If

If LifeX < -200 Then
SCLive = SCLive - 1
BDestroy1 = True
LifeX = 0
RBIcon
End If

If SCLive <= 0 Then
SCLive = 0
BDestroy1 = True
If ExitC = 20 Then
SpriteExist = False
MainDC = False
Else
ExitC = ExitC + 1
End If
End If

If EnemyLifeC < 0 Then
BDestroy2 = True
EnemyLifeC = 0
End If

If BDestroy2 = True Then
If ExitC = 300 Then
BDestroy2 = False
EnemyExist = False
MainDC = False
Else
ExitC = ExitC + 1
End If
End If

If BDestroy = True Then
If DestroyX = 480 - 120 Then
DestroyX = 0
If DestroyY = 600 - 120 Then
DestroyY = 0
BDestroy = False
Else
DestroyY = DestroyY + 120
End If
Else
DestroyX = DestroyX + 120
End If
DestroyRect.Left = DestroyX
DestroyRect.Right = 120 + DestroyX
DestroyRect.Top = DestroyY
DestroyRect.Bottom = 120 + DestroyY
BackSurface.BltFast Bombn.X - 60, Bombn.Y - 60, DestroySurface, DestroyRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

If BDestroy1 = True Then
If DestroyX1 = 480 - 120 Then
DestroyX1 = 0
If DestroyY1 = 600 - 120 Then
DestroyY1 = 0
BDestroy1 = False
Else
DestroyY1 = DestroyY1 + 120
End If
Else
DestroyX1 = DestroyX1 + 120
End If
DestroyRect1.Left = DestroyX1
DestroyRect1.Right = 120 + DestroyX1
DestroyRect1.Top = DestroyY1
DestroyRect1.Bottom = 120 + DestroyY1
BackSurface.BltFast SpriteX - 50, SpriteY - 58, DestroySurface, DestroyRect1, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

If BDestroy2 = True Then
If DestroyX2 = 480 - 120 Then
DestroyX2 = 0
If DestroyY2 = 600 - 120 Then
DestroyY2 = 0
Else
DestroyY2 = DestroyY2 + 120
End If
Else
DestroyX2 = DestroyX2 + 120
End If
DestroypX = EnemyX + 70 * Rnd
DestroypY = EnemyY + 15 * Rnd
DestroyRect2.Left = DestroyX2
DestroyRect2.Right = 120 + DestroyX2
DestroyRect2.Top = DestroyY2
DestroyRect2.Bottom = 120 + DestroyY2
BackSurface.BltFast DestroypX, DestroypY, DestroySurface, DestroyRect2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

Public Sub UpdateControl()

GetKey

If DIState.Key(DIK_LEFT) Then
If SpriteX > 20 Then
SpriteX = SpriteX - 1
End If
End If
If DIState.Key(DIK_RIGHT) Then
If SpriteX < 600 Then
SpriteX = SpriteX + 1
End If
End If
If DIState.Key(DIK_UP) Then
If SpriteY > 20 Then
SpriteY = SpriteY - 1
End If
End If
If DIState.Key(DIK_DOWN) Then
If SpriteY < 440 Then
SpriteY = SpriteY + 1
End If
End If

If DIState.Key(DIK_A) Then
If IntC = 6 Then
IntC = 0
If Intcount = 7 Then
Intcount = 0
Else
Intcount = Intcount + 1
End If
Else
IntC = IntC + 1
End If

If Firen1(Intcount).Exist = False Then
Firen1(Intcount).X = SpriteX + 7
Firen1(Intcount).Y = SpriteY - 7
Firen1(Intcount).Exist = True
PlaySound DSFire, -600, False
If SpriteX > EnemyX Then
SetSoundPan DSFire, 10000
Else
SetSoundPan DSFire, -10000
End If
End If
End If

If DIState.Key(DIK_S) Then
If BIcount = 0 Then
Else
If Bombn.Exist = False Then
Bombn.X = SpriteX + 4
Bombn.Y = SpriteY - 40
Bombn.Exist = True
BIcount = BIcount - 1
PlaySound DSBomb, -500, False
End If
End If
End If

End Sub

Public Sub UpdateControlFunction()

GetKey

If DIState.Key(DIK_RETURN) Then
If PauseN > 8 Then
If GPlay = True Then
MainDC = False
GPlay = False
Else
MainDC = True
GPlay = True
End If
PauseN = 0
Else
PauseN = PauseN + 1
End If
End If

End Sub

Public Sub RBIcon()
For Intcount = 0 To 2
BIcons(Intcount).X = 0
BIcons(Intcount).Y = 440
BIcons(Intcount).Exist = True
Next
Intcount = 0
BIcons(0).X = 20
BIcons(1).X = 40
BIcons(2).X = 60
BIcount = 3
End Sub

Public Sub EndGame()

On Error Resume Next

DIDevice.Unacquire
Set DIDevice = Nothing
Set DInput = Nothing

Set DSFire = Nothing
Set DSBomb = Nothing
Set DSBomb1 = Nothing
Set DSStart = Nothing
Set DSForce = Nothing
Set DSDestroy = Nothing
Set DSBuffer = Nothing
Set DSound = Nothing

If Not (DMSeg Is Nothing) And Not (DMSegState Is Nothing) Then
If DMPerformance.IsPlaying(DMSeg, DMSegState) Then
Call DMPerformance.Stop(DMSeg, DMSegState, 0, 0)
End If
End If

Form1.Init

End Sub
