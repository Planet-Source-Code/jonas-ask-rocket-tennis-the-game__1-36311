Attribute VB_Name = "Publics"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020

Public Type aBall
 X As Currency
 Y As Currency
 Speed As Currency
 Ang As Integer
End Type

Public Type aPlayer
 X As Integer
 Y As Currency
 Speed As Currency
 LastFire As Long
 BurnUP As Boolean
 BurnDown As Boolean
 Score As Integer
End Type

Public Type Rocket
 X As Currency
 Y As Currency
 Act As Boolean
 Dire As Integer
End Type

Public MySin(-360 To 360) As Currency
Public MyCos(-360 To 360) As Currency

Public Ball As aBall
Public P(1 To 2) As aPlayer
Public Roc(1 To 100) As Rocket

Public WD As Integer
Public HG As Integer


Public Const BallRad As Integer = 5
Public Const FireDelay As Integer = 130
Public Sub PaintBoard()
    Main.PicM.Cls
    'The astroid
     BitBlt Main.PicM.hdc, Ball.X - BallRad, Ball.Y - BallRad, 13, 13, Main.picAstM.hdc, 0, 0, SRCAND
     BitBlt Main.PicM.hdc, Ball.X - BallRad, Ball.Y - BallRad, 13, 13, Main.picAst.hdc, 0, 0, SRCPAINT
     
     'The players
     BitBlt Main.PicM.hdc, P(1).X - 10, P(1).Y - 8, 13, 17, Main.picShip2M.hdc, 0, 0, SRCAND
     BitBlt Main.PicM.hdc, P(1).X - 10, P(1).Y - 8, 13, 17, Main.picShip2.hdc, 0, 0, SRCPAINT
     BitBlt Main.PicM.hdc, P(2).X - 3, P(2).Y - 8, 13, 17, Main.picShip1M.hdc, 0, 0, SRCAND
     BitBlt Main.PicM.hdc, P(2).X - 3, P(2).Y - 8, 13, 17, Main.picShip1.hdc, 0, 0, SRCPAINT
     
     'Afterburners
     If P(1).BurnDown Then BitBlt Main.PicM.hdc, P(1).X - 9, P(1).Y - 9, 3, 2, Main.picBurn2M.hdc, 0, 0, SRCAND
     If P(1).BurnDown Then BitBlt Main.PicM.hdc, P(1).X - 9, P(1).Y - 9, 3, 2, Main.picBurn2.hdc, 0, 0, SRCPAINT
     If P(2).BurnDown Then BitBlt Main.PicM.hdc, P(2).X + 6, P(2).Y - 9, 30, 20, Main.picBurn2M.hdc, 0, 0, SRCAND
     If P(2).BurnDown Then BitBlt Main.PicM.hdc, P(2).X + 6, P(2).Y - 9, 30, 20, Main.picBurn2.hdc, 0, 0, SRCPAINT
     If P(1).BurnUP Then BitBlt Main.PicM.hdc, P(1).X - 9, P(1).Y + 8, 3, 2, Main.picBurn1M.hdc, 0, 0, SRCAND
     If P(1).BurnUP Then BitBlt Main.PicM.hdc, P(1).X - 9, P(1).Y + 8, 3, 2, Main.picBurn1.hdc, 0, 0, SRCPAINT
     If P(2).BurnUP Then BitBlt Main.PicM.hdc, P(2).X + 6, P(2).Y + 8, 3, 2, Main.picBurn1M.hdc, 0, 0, SRCAND
     If P(2).BurnUP Then BitBlt Main.PicM.hdc, P(2).X + 6, P(2).Y + 8, 3, 2, Main.picBurn1.hdc, 0, 0, SRCPAINT
     
     For a = 1 To UBound(Roc)
        If Roc(a).Act Then
            Main.PicM.PSet (Roc(a).X, Roc(a).Y)
        End If
    Next a
End Sub

Public Sub DoBall()
    With Ball
        For b = 0 To .Speed - 0.1 'step 0.1
            .X = .X + MyCos(.Ang) * 0.1
            .Y = .Y - MySin(.Ang) * 0.1
            CheckColide Ball
            
            If .X < -10 Or .Y < -10 Then Stop
            If .X > WD + 10 Or .Y > HG + 10 Then Stop
        Next b
        If Ball.Speed = 0 Then
            If Ball.X < 10 Or Ball.X > WD - 10 Then
                Ball.X = WD / 2
            End If
        End If
    End With
End Sub

Public Sub CheckColide(Pb As aBall)
    If Pb.Ang < 0 Then Pb.Ang = 360 + Pb.Ang
    If Int(Pb.X) = 0 Or Int(Pb.X) = WD Then
        If MyCos(Pb.Ang) > 0 Then
            P(1).Score = P(1).Score + 1
        Else
            P(2).Score = P(2).Score + 1
        End If
        Pb.Ang = 180 - Pb.Ang
    End If
    If Int(Pb.Y) = 0 Or Int(Pb.Y) = HG Then
        Pb.Ang = 360 - Pb.Ang
    End If
End Sub

Public Sub DoKeys()
    P(1).BurnUP = False:        P(1).BurnDown = False
    P(2).BurnUP = False:        P(2).BurnDown = False
    If GetAsyncKeyState(vbKeyUp) Then P(2).Speed = P(2).Speed - 0.1: P(2).BurnUP = True
    If GetAsyncKeyState(vbKeyDown) Then P(2).Speed = P(2).Speed + 0.1: P(2).BurnDown = True
    If GetAsyncKeyState(vbKeyNumpad0) Then Fire (2)
    
    If GetAsyncKeyState(vbKeyW) Then P(1).Speed = P(1).Speed - 0.1: P(1).BurnUP = True
    If GetAsyncKeyState(vbKeyS) Then P(1).Speed = P(1).Speed + 0.1: P(1).BurnDown = True
    If GetAsyncKeyState(vbKeySpace) Then Fire (1)

End Sub
Public Sub MovePlayers()
    For a = 1 To 2
    With P(a)
        .Speed = .Speed * 0.99
        If .Speed < -20 Then .Speed = -20
        If .Speed > 20 Then .Speed = 20
        If .Speed <> 0 Then
            For b = 0 To .Speed Step 0.1 * (.Speed / Abs(.Speed))
                .Y = .Y + 0.1 * (.Speed / Abs(.Speed))
                If .Y <= 0 Then .Speed = -.Speed
                If .Y >= HG Then .Speed = -.Speed
            Next b
        End If
    End With
    Next a
End Sub
Public Sub Fire(Ply)
    If GetTickCount < P(Ply).LastFire + FireDelay Then Exit Sub
    For a = 1 To UBound(Roc)
        If Roc(a).Act = False Then
            Roc(a).Act = True
            Roc(a).Dire = IIf(Ply = 1, 1, -1)
            Roc(a).X = P(Ply).X + IIf(Ply = 1, 2, -2)
            Roc(a).Y = P(Ply).Y
            
            P(Ply).LastFire = GetTickCount
            Exit For
        End If
    Next a
End Sub
Public Sub DoRockets()
    For a = 1 To UBound(Roc)
    With Roc(a)
        If .Act Then
            Speed = .Dire / 3
            For b = 0 To Speed Step 0.01 * (Speed / Abs(Speed))
                .X = .X + 0.1 * (Speed / Abs(Speed))
                If CheckForImpact(Roc(a)) Then Exit For
            Next b
            
            
            
            If .X < 0 Or .X > WD Then .Act = False
        End If
    End With
    Next a
End Sub
Function CheckForImpact(R As Rocket) As Boolean
Const Pi = 3.14159265358979
Dim Ang As Integer
Dim xSpd As Currency, ySpd As Currency
    'impact with astroid
    If R.X > Ball.X - BallRad And R.X < Ball.X + BallRad Then
    If R.Y > Ball.Y - BallRad And R.Y < Ball.Y + BallRad Then
        ydis = Ball.Y - R.Y
        xdis = -Ball.X + R.X
        If xdis = 0 Then Exit Function
        Ang = Atn(ydis / xdis) / Pi * 180
        If R.Dire = -1 Then Ang = Ang + 180
        
        xSpd = MyCos(Ball.Ang) * Ball.Speed
        ySpd = MySin(Ball.Ang) * Ball.Speed
        
        xSpd = xSpd + IIf(MyCos(Ball.Ang) < 0, MyCos(Ang), MyCos(Ang)) * 6
        ySpd = ySpd + MySin(Ang) * 2
        
        If xSpd = 0 Then
        Else
            Ball.Ang = Atn(ySpd / xSpd) / Pi * 180
        End If
        If xSpd < 0 Then Ball.Ang = Ball.Ang + 180
        Ball.Speed = Sqr(ySpd ^ 2 + xSpd ^ 2)
        
        
        R.Act = False
        CheckForImpact = True
    End If
    End If
    
End Function
