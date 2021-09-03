VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Direct War - DXL"
   ClientHeight    =   7200
   ClientLeft      =   1020
   ClientTop       =   1065
   ClientWidth     =   9600
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6960
      Picture         =   "Form1.frx":0594
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   6720
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      ScaleHeight     =   345
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   6510
      Visible         =   0   'False
      Width           =   2295
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -240
         ScaleHeight     =   375
         ScaleWidth      =   135
         TabIndex        =   1
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   1080
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   3600
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   3480
      Top             =   5040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   2400
      Top             =   3960
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   3600
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3000
      Top             =   1920
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   3360
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Initb
Init
End Sub

Sub Initb()
If RWidth <> 800 Or RHeight <> 600 Then
RWidth = 800
RHeight = 600
Call InitializeDD
End If
End Sub

Sub Init()

Form1.MouseIcon = LoadPicture(App.Path & "\Data\Mouse.cur")
Form1.MousePointer = 99

Proce
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Picture1.Visible = False
Picture3.Visible = False

Form1.Cls
Form1.PaintPicture LoadPicture(App.Path & "\data\Flash.gif"), 0, 0, 640, 480

Timer1.Enabled = True

Picture1.Visible = True

End Sub

Private Sub Image1_Click()
On Error Resume Next
PlaySoundC App.Path & "\Data\Click.wav", 1, 1
Main
End Sub

Private Sub Image2_Click()
On Error Resume Next
PlaySoundC App.Path & "\Data\Click.wav", 1, 1
Picture3.Top = 136
EnemyLifeO = 1000
End Sub

Private Sub Image3_Click()
On Error Resume Next
PlaySoundC App.Path & "\Data\Click.wav", 1, 1
Picture3.Top = 200
EnemyLifeO = 2000
End Sub

Private Sub Image4_Click()
On Error Resume Next
PlaySoundC App.Path & "\Data\Click.wav", 1, 1
Picture3.Top = 272
EnemyLifeO = 3000
End Sub

Private Sub Image5_Click()
On Error Resume Next
PlaySoundC App.Path & "\Data\Click.wav", 1, 1
Picture1.Visible = False
Picture3.Visible = False
Form1.Cls
Form1.PaintPicture LoadPicture(App.Path & "\data\Ab.bmp"), 0, 0, 640, 480
End Sub

Private Sub Image6_Click()
On Error Resume Next
PlaySoundC App.Path & "\Data\Click.wav", 1, 1
CloseGame
Unload Me
End Sub

Private Sub Image7_Click()
On Error Resume Next
PlaySoundC App.Path & "\Data\Click.wav", 1, 1
Picture1.Visible = False
Picture3.Visible = True
Form1.Cls
Form1.PaintPicture LoadPicture(App.Path & "\data\Screen.bmp"), 0, 0, 640, 480
End Sub

Sub Timer1_Timer()
If Picture2.Width > 2530 Then
Picture2.Width = 135
Picture1.Visible = False
Form1.Cls
Timer1.Enabled = False
Init1
Else
Picture2.Width = Picture2.Width + 55
End If
End Sub

Sub Init1()
Image1.Visible = True
Image2.Visible = True
Image3.Visible = True
Image4.Visible = True
Image5.Visible = True
Image6.Visible = True
Image7.Visible = True
Picture1.Visible = False
Picture3.Visible = True
Form1.MouseIcon = LoadPicture(App.Path & "\Data\MouseA.ico")
Form1.MousePointer = 99
Form1.PaintPicture LoadPicture(App.Path & "\data\Screen.bmp"), 0, 0, 640, 480
End Sub

Sub CloseGame()
DirectDraw.RestoreDisplayMode
DirectDraw.SetCooperativeLevel Form1.hWnd, DDSCL_NORMAL
End
End Sub

Sub Proce()
For Intcount = 0 To 4
Firen(Intcount).X = 0
Firen(Intcount).Y = 0
Firen(Intcount).Exist = False
Next
Intcount = 0

For Intcount = 0 To 4
Firem(Intcount).X = 0
Firem(Intcount).Y = 0
Firem(Intcount).Exist = False
Next
Intcount = 0

For Intcount = 0 To 6
Bombn1(Intcount).X = 0
Bombn1(Intcount).Y = 0
Bombn1(Intcount).Exist = False
Next
Intcount = 0

For Intcount = 0 To 7
Firen1(Intcount).X = 0
Firen1(Intcount).Y = 0
Firen1(Intcount).Exist = False
Next
Intcount = 0

For Intcount = 0 To 6
Cloudn(Intcount).X = 0
Cloudn(Intcount).Y = 0
Cloudn(Intcount).Exist = False
Next
Intcount = 0

For Intcount = 0 To 6
Bombn1(Intcount).X = 0
Bombn1(Intcount).Y = 0
Bombn1(Intcount).Exist = False
Next
Intcount = 0

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

Bombn.X = 0
Bombn.Y = 0
Bombn.Exist = False
End Sub
