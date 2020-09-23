VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rocket Tennis"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2520
      Picture         =   "main.frx":0442
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picAstM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2760
      Picture         =   "main.frx":068C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picBurn2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   5520
      Picture         =   "main.frx":08D6
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picBurn2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   5280
      Picture         =   "main.frx":0930
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picBurn1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   4980
      Picture         =   "main.frx":098A
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   7
      Top             =   180
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picBurn1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   4740
      Picture         =   "main.frx":09E4
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picShip2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4140
      Picture         =   "main.frx":0A3E
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picShip2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3900
      Picture         =   "main.frx":0D28
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picShip1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3600
      Picture         =   "main.frx":1012
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picShip1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3360
      Picture         =   "main.frx":12FC
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox PicM 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   4515
      Left            =   60
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   0
      Top             =   540
      Width           =   6675
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "-------"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   6645
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Show
    Randomize
    DoEvents
    'set various things
    Const Pi = 3.14159265358979
    For a = -360 To 360
        MyCos(a) = Cos(a / 180 * Pi)
        MySin(a) = Sin(a / 180 * Pi)
    Next a
    'board dimentions
    WD = PicM.ScaleWidth
    HG = PicM.ScaleHeight
    
    Ball.X = WD / 2
    Ball.Y = HG / 2
    P(1).X = 12
    P(1).Y = HG / 2
    P(2).X = WD - 12
    P(2).Y = HG / 2
    
    MainLoop
End Sub
Sub MainLoop()
    Do
        t = GetTickCount
        'Framelimiter
        Do Until GetTickCount > TickTag 'Or MaxSpeed
        DoEvents
        Loop: TickTag = GetTickCount + 30 'TickTag = IIf(Faster, GetTickCount + 25, GetTickCount + 50)
        DoEvents
        
        DoKeys
        MovePlayers
        DoBall
        DoRockets
        
        PaintBoard
       
        'Main.Label2.Caption = "ang " & Ball.Ang & "  speed " & Ball.Speed
        lblScores.Caption = "Player I: " & P(1).Score & "                         Player II: " & P(2).Score
        'Label1.Caption = GetTickCount - t
    Loop Until FlagTermination
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
