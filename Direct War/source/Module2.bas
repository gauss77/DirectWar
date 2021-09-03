Attribute VB_Name = "Module2"
Option Explicit

Public RWidth As Long
Public RHeight As Long

Public Sub InitializeDD()
  
On Local Error GoTo ErrorHandler
    
Set DirectDraw = DirectX.DirectDrawCreate("")
DirectDraw.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE

DirectDraw.SetDisplayMode RWidth, RHeight, 16, 0, 0

FrontDdsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
FrontDdsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_SYSTEMMEMORY                                                                                 'this surface will be the primary, flipable surface, it's a complex surface, and it will be located in system memory
FrontDdsd.lBackBufferCount = 1

Set FrontSurface = DirectDraw.CreateSurface(FrontDdsd)

DDCaps.lCaps = DDSCAPS_BACKBUFFER
Set BackSurface = FrontSurface.GetAttachedSurface(DDCaps)

ColorKey.high = 0
ColorKey.low = 0

BackSurface.SetForeColor vbRed

Set BackGSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Stage.bmp")
Set SpriteSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Sprite.gif")
SpriteSurface.SetColorKey DDCKEY_SRCBLT, ColorKey
Set EnemySurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Enemy.gif")
EnemySurface.SetColorKey DDCKEY_SRCBLT, ColorKey
Set LaserSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Fire.gif")
LaserSurface.SetColorKey DDCKEY_SRCBLT, ColorKey
Set LaserSurface1 = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Fire1.gif")
LaserSurface1.SetColorKey DDCKEY_SRCBLT, ColorKey
Set BombSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Bomb.gif")
BombSurface.SetColorKey DDCKEY_SRCBLT, ColorKey
Set BombSurface1 = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Bomb1.gif")
BombSurface1.SetColorKey DDCKEY_SRCBLT, ColorKey
Set LifeSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Life.gif")
Set DestroySurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Destroy.gif")
DestroySurface.SetColorKey DDCKEY_SRCBLT, ColorKey
Set BIconSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\BIcon.gif")
Set CloudSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Cloud.gif")
CloudSurface.SetColorKey DDCKEY_SRCBLT, ColorKey
Set HitSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Hit.gif")
HitSurface.SetColorKey DDCKEY_SRCBLT, ColorKey
Set LoverSurface = CreateDDSFromBitmap(DirectDraw, App.Path & "\Data\Lover.gif")
LoverSurface.SetColorKey DDCKEY_SRCBLT, ColorKey

Exit Sub

ErrorHandler:

If Err.Number = DDERR_NOEXCLUSIVEMODE Then
MsgBox "Unable to create Direct Draw object. Ensure that your drivers are the latest available for your video card."
EndGame
Else
EndGame
End If

End Sub

Public Function CreateDDSFromBitmap(dd As DirectDraw7, ByVal strFile As String, Optional VideoMem As Boolean) As DirectDrawSurface7
    
If Dir(strFile) = "" Then
                                                                        
End If
    
Dim ddsd As DDSURFACEDESC2
Dim dds As DirectDrawSurface7
Dim hdcPicture As Long
Dim hdcSurface As Long
Dim Picture As StdPicture
    
Set Picture = LoadPicture(strFile)
    
With ddsd
.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
If VideoMem Then
.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
Else
.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
End If
.lWidth = Screen.ActiveForm.ScaleX(Picture.Width, vbHimetric, vbPixels)
                                                                           
.lHeight = Screen.ActiveForm.ScaleY(Picture.Height, vbHimetric, vbPixels)
                                                                           
End With
Set dds = dd.CreateSurface(ddsd)
hdcPicture = CreateCompatibleDC(ByVal 0&)
SelectObject hdcPicture, Picture.Handle
dds.restore
hdcSurface = dds.GetDC
StretchBlt hdcSurface, 0, 0, ddsd.lWidth, ddsd.lHeight, hdcPicture, 0, 0, Screen.ActiveForm.ScaleX(Picture.Width, vbHimetric, vbPixels), Screen.ActiveForm.ScaleY(Picture.Height, vbHimetric, vbPixels), SRCCOPY
                                                                            
dds.ReleaseDC hdcSurface
DeleteDC hdcPicture
Set Picture = Nothing
Set CreateDDSFromBitmap = dds
    
End Function

Public Function DetectCollision(ObjectRectA As RECT, ObjectRectB As RECT) As Boolean
Dim ReturnRect As RECT
If IntersectRect(ReturnRect, ObjectRectA, ObjectRectB) Then
DetectCollision = True
End If
End Function

Public Function CScreenObject(ObjectRect As RECT, ObjectPosition As DDPosition, ObjectW As Long, ObjectH As Long, ObjectM As Single, ObjectMD As Integer, OScreenW As Long, OScreenH As Long, ObjectOP As RECT)

If ObjectMD = 1 Then
If ObjectOP.Right <= 0 Then
If ObjectPosition.X <= 0 Then
If ObjectRect.Left >= ObjectW Then
ObjectOP.Left = ObjectOP.Left + ObjectM
Else
ObjectRect.Left = ObjectRect.Left + ObjectM
End If
Else
ObjectPosition.X = ObjectPosition.X - ObjectM
If ObjectRect.Right >= ObjectW Then
Else
ObjectRect.Right = ObjectRect.Right + ObjectM
End If
End If
Else
ObjectOP.Right = ObjectOP.Right - ObjectM
End If
End If

If ObjectMD = 2 Then
If ObjectOP.Left <= 0 Then
If ObjectPosition.X >= OScreenW - ObjectW Then
If ObjectRect.Right <= 0 Then
ObjectOP.Right = ObjectOP.Right + ObjectM
Else
ObjectRect.Right = ObjectRect.Right - ObjectM
ObjectPosition.X = ObjectPosition.X + ObjectM
End If
Else
If ObjectRect.Left <= 0 Then
ObjectPosition.X = ObjectPosition.X + ObjectM
Else
ObjectRect.Left = ObjectRect.Left - ObjectM
End If
End If
Else
ObjectOP.Left = ObjectOP.Left - ObjectM
End If
End If

If ObjectMD = 4 Then
If ObjectOP.Top <= 0 Then
If ObjectPosition.Y >= OScreenH - ObjectH Then
If ObjectRect.Bottom <= 0 Then
ObjectOP.Bottom = ObjectOP.Bottom + ObjectM
Else
ObjectRect.Bottom = ObjectRect.Bottom - ObjectM
ObjectPosition.Y = ObjectPosition.Y + ObjectM
End If
Else
If ObjectRect.Top <= 0 Then
ObjectPosition.Y = ObjectPosition.Y + ObjectM
Else
ObjectRect.Top = ObjectRect.Top - ObjectM
End If
End If
Else
ObjectOP.Top = ObjectOP.Top - ObjectM
End If
End If

If ObjectMD = 3 Then
If ObjectOP.Bottom <= 0 Then
If ObjectPosition.Y <= 0 Then
If ObjectRect.Top >= ObjectH Then
ObjectOP.Top = ObjectOP.Top + ObjectM
Else
ObjectRect.Top = ObjectRect.Top + ObjectM
End If
Else
ObjectPosition.Y = ObjectPosition.Y - ObjectM
If ObjectRect.Bottom >= ObjectH Then
Else
ObjectRect.Bottom = ObjectRect.Bottom + ObjectM
End If
End If
Else
ObjectOP.Bottom = ObjectOP.Bottom - ObjectM
End If
End If

End Function

Public Function ObjectExist(ObjectRect As RECT, ObjectW As Long, ObjectH As Long) As Boolean
If ObjectRect.Right <= ObjectW And ObjectRect.Left <= ObjectW And ObjectRect.Bottom <= ObjectH And ObjectRect.Top <= ObjectH Then
ObjectExist = True
Else
ObjectExist = False
End If
End Function