VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Particle System by Stéphanie Rancourt"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "quit"
      Height          =   375
      Left            =   10080
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "blur"
      Height          =   495
      Left            =   10080
      TabIndex        =   5
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "render"
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   8040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   10200
      TabIndex        =   2
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox DDPIC 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   0
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'* ParticulSystem:particle simulator                                          *
'******************************************************************************
'*Author: Stéphanie Rancourt                                                  *
'*Contact:asuka_tetsuo@hotmail.com                                            *
'*Date:20/04/2002                                                             *
'*Version: 1                                                                  *
'*Note: 16bit mode not supported completely                                   *
'*Comment from the author: You may use or alter this code for your personal   *
'*  use.If you do please credit the author.You can't sell or use it for       *
'*  commercial purpose without the approval of the author but you may         *
'*  distribute it in it's integrity.                                          *
'******************************************************************************
Dim img() As Byte
Dim pal(255) As Long
Dim running As Boolean
Private Type vector2D
    x As Single
    y As Single
End Type
Dim dd8b As New DD8BitSurf
Dim Maxparticul As Long
Dim BlurFactor As Long
Dim windenable As Boolean
Dim windpower As Single
Dim PGFrom As Long
Dim PGVaray As Long
Dim PGEnable As Boolean
Dim Yoffset As Long
Dim Yoffsetenabled As Boolean
Dim onclickPartnum As Long
Dim palletteind As Long
Dim SpeedVaray As Single
Dim SpeedFrom As Single
Dim partrate As Long
Dim partrate2 As Long
Dim PartColorDecRate As Single
Dim PartColorDecRate2 As Single
Dim ClearSurface As Boolean

Private Type Particul
    x As Single
    y As Single
    VX As Single
    VY As Single
    color As Single
    colordec As Single
End Type
Dim part() As Particul
Dim PartCount As Long
Dim partcolorVaray As Byte
Dim Partcolorfrom As Byte


Dim PgenX As Long
Dim PgenY As Long
Dim PTimecount As Long

Dim Mx As Long
Dim My As Long
Dim GravVX As Single
Dim GravVY As Single
Dim SinV(360) As Single
Dim CosV(360) As Single
Dim Wind(240) As Single
Dim WindTime As Long
Dim WindTimeex As Long
Dim Foyer As Long

Public Function AnimWind()
'Anim wind power and region
Dim tick As Long
Dim V0 As vector2D
Dim V1 As vector2D
Dim V2 As vector2D
Dim V3 As vector2D
Dim rv As vector2D

tick = GetTickCount()
If tick - WindTime >= WindTimeex Then
    WindTimeex = Int(Rnd * 1000) + 1000
    WindTime = tick

    For x = 0 To 5
        V0.x = Rnd * windpower - (windpower / 2)
        V0.y = 0
        V1.x = V0.x
        V1.y = 10
        V3.x = Rnd * windpower - (windpower / 2)
        V3.y = 40
        V2.x = V3.x
        V3.y = 30
        For t = 0 To 39
            rv = IntBezier(V0, V1, V2, V3, t / 39)
            Wind(t + x * 40) = rv.x
            
        Next t
    Next x
End If
For t = 0 To Maxparticul
    If part(t).y >= 0 And part(t).y <= 240 Then
        part(t).x = part(t).x + Wind(part(t).y)
    End If
Next t
        
End Function
Private Function IntBezier(V0 As vector2D, V1 As vector2D, V2 As vector2D, V3 As vector2D, ByVal a As Single) As vector2D
    IntBezier.x = ((1 - a) ^ 3) * V0.x + (3 * a) * ((1 - a) ^ 2) * V1.x + (3 * a * a) * (1 - a) * V2.x + (a ^ 3) * V3.x
    IntBezier.y = ((1 - a) ^ 3) * V0.y + (3 * a) * ((1 - a) ^ 2) * V1.y + (3 * a * a) * (1 - a) * V2.y + (a ^ 3) * V3.y
End Function

Public Function InitTrigo()
Dim t As Long
For t = 0 To 360
    SinV(t) = Sin(t / 180 * 3.1415)
    CosV(t) = Cos(t / 180 * 3.1415)
Next t

End Function

Public Function animPart()
'Move particle
Dim tick As Long
Dim x As Long
Dim t As Long
Dim d As Long
Dim d2 As Single
For t = 0 To Maxparticul
    part(t).VX = part(t).VX + GravVX
    part(t).VY = part(t).VY + GravVY
    part(t).x = part(t).x + part(t).VX
    part(t).y = part(t).y + part(t).VY
    If part(t).color >= 1 Then
        If part(t).colordec <= part(t).color Then
            part(t).color = part(t).color - part(t).colordec
        Else
            part(t).color = 0
        End If
    End If
Next t

If PGEnable Then

    'tick = GetTickCount()
    'If tick - PTimecount >= 0 Then
            If Foyer = 1 Or Foyer = 3 Then
                My = DDPIC.Height * 0.9
            ElseIf Foyer = 2 Then
                Mx = Int(Rnd * (DDPIC.Width - 10)) + 5
                My = Int(Rnd * (DDPIC.Width - 10)) + 5
            End If
            
            For t = 1 To partrate + (Int(Rnd * partrate2 * 2) - partrate2) Step 1
                If Foyer = 1 Then
                    Mx = Int(Rnd * (DDPIC.Width - 10)) + 5
                ElseIf Foyer = 3 Then
                    Mx = Int(Rnd * (DDPIC.Width - 50)) + 25
                
                End If

                PartCount = PartCount + 1
                If PartCount > Maxparticul Then
                    PartCount = 0
                End If
                d = Int(Rnd * PGVaray) + PGFrom
                d2 = Rnd * SpeedVaray + SpeedFrom
                part(PartCount).VX = CosV(d) * d2
                part(PartCount).VY = SinV(d) * d2
                part(PartCount).x = Mx 'Int(Rnd * 8) - 4 + Mx
                part(PartCount).y = My
                part(PartCount).color = Int(Rnd * partcolorVaray) + Partcolorfrom
                part(PartCount).colordec = PartColorDecRate + (Int(Rnd * PartColorDecRate2 * 2) - PartColorDecRate2)
            Next t
        'PTimecount = tick
    'End If
End If
Dim s As Single
Dim c As Single




If Yoffsetenabled Then
    For t = 0 To (DDPIC.Height) - Yoffset
        For x = 0 To (DDPIC.Width)
            img(x, t) = img(x, t + Yoffset)
        Next x
    Next t
End If
End Function

Public Function Initpart()
'Set particle off screen
Dim t As Long
For t = 0 To Maxparticul
    part(t).x = 1000
    part(t).y = 1000
Next t

End Function
Public Function DrawPart()
'Draw all particles on the buffer surface
Dim t As Long, x As Long, y As Long
If ClearSurface Then
    For x = 0 To DDPIC.Width - 1
        For y = 0 To DDPIC.Height - 1
            img(x, y) = 0
        Next y
    Next x
End If
For t = 0 To Maxparticul
    If part(t).x > 0 And part(t).x < DDPIC.Width And part(t).y > 0 And part(t).y < DDPIC.Height Then
        img(part(t).x, part(t).y) = part(t).color
    End If
Next t
End Function
Public Function initpal()
'Select the palette we're going to use
Dim t As Long

Select Case palletteind
Case 0
palfire:
    dd8b.Load_MPAL App.Path & "\fire.pal", pal()
    Exit Function
Case 1
palwater:
    For t = 0 To 63
        pal(t) = RGB(255, _
        64 + t, _
        0)
    Next t
    For t = 64 To 127
        pal(t) = RGB(255, _
        (t - 64) * 2.04 + 127, _
        (t - 64) * 4.04)
    Next t
    For t = 128 To 255
        pal(t) = RGB(255, _
        255 - (t - 128) * 2, _
        255 - (t - 128) * 2)
    Next t
    Exit Function
Case 2
palwater2:
    For t = 0 To 63
        pal(t) = RGB(255, _
        64 + t, _
        0)
    Next t
    For t = 64 To 191
        pal(t) = RGB(255, _
        Sin(t / 191 * 6.28) * 25 + 200, _
        Sin(t / 191 * 6.28) * 25 + 200)
    Next t
    For t = 192 To 255
        pal(t) = RGB(255, _
        255 - ((t - 192) * 2.04 + 125), _
        255 - ((t - 192) * 4.04))
    Next t
    Exit Function
allcolor:
    For t = 0 To 255
        pal(t) = RGB(Sin(t / 255 * 50) * 127 + 128, _
        Sin(t / 255 * 25) * 127 + 128, _
        Sin(t / 255 * 12) * 127 + 128)
    Next t
    Exit Function
Case 3
grayscall:
    For t = 0 To 255
        pal(t) = RGB(t, _
        t, _
        t)
    Next t
    Exit Function
Case 4
water2:
    For t = 0 To 255
        pal(t) = RGB(255, _
        t, _
        t)
    Next t
    Exit Function
Case 5
    For t = 0 To 255
        pal(t) = RGB(0, _
         Sin(t / 255 * 64) * 127 + 128, _
        255 - Sin(t / 255 * 64) * 127 + 128)
    Next t
Case Else
    End
    
End Select

End Function



Private Sub DDPIC_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift

End Sub

Private Sub DDPIC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If y > 240 Or x > 320 Then
    If x > 320 Then x = 320
    If y > 240 Then y = 240
    SetCursorPos x, y
End If
Select Case Button
Case 1
        For t = 1 To onclickPartnum
            PartCount = PartCount + 1
            If PartCount > Maxparticul Then
                PartCount = 0
            End If
            d = Int(Rnd * 360)
            part(PartCount).VX = CosV(d)
            part(PartCount).VY = SinV(d)
            part(PartCount).x = Int(Rnd * 8) - 4 + Mx
            part(PartCount).y = My
            part(PartCount).color = 255
        Next t
Case 2
    
'Dim pos As Long, col As Long
'Dim r As Single, g As Single, b As Single
'pos = 25
'Open "c:\tmp\mpal.pal" For Binary As #1
'For t = 0 To 255

'r = DX7_DX.ColorGetRed(pal(t))
'g = DX7_DX.ColorGetGreen(pal(t))
'b = DX7_DX.ColorGetBlue(pal(t))
'col = DX7_DX.CreateColorRGBA(b, g, r, 0)
'    Put #1, pos, col
'    'put #1, pos, r
'    'pos = pos + 1
'    'Get #1, pos, g
'    'pos = pos + 1
'    'Get #1, pos, b
'    pos = pos + 4
'
'     'pal(t) = RGB(b, g, r)
'Next t
'Close #1
'    'Exit Sub
'    Open "c:\tmp\tmp.raw" For Binary As #1
'        pos = 1
'        For Y = 0 To 239
'            For X = 0 To 319
'                Put #1, pos, img(X, Y)
'                pos = pos + 1
'            Next X
'        Next Y
'    Close #1
End Select


End Sub

Private Sub DDPIC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If y > 240 Or x > 320 Then
    If x > 320 Then x = 320
    If y > 240 Then y = 240
    SetCursorPos x, y
End If
If Foyer = 0 Then
    Mx = x
    My = y
End If
Select Case Button
Case 1
        For t = 1 To 10
            PartCount = PartCount + 1
            If PartCount > Maxparticul Then
                PartCount = 0
            End If
            d = Int(Rnd * 360)
            part(PartCount).VX = CosV(d) * 0.5
            part(PartCount).VY = SinV(d) * 0.5
            part(PartCount).x = Int(Rnd * 8) - 4 + Mx
            part(PartCount).y = My
            part(PartCount).color = 255
        Next t
End Select



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
running = False

End Sub

Private Sub Form_Load()
Dim t As Long, x As Long
Dim ps As Boolean
Dim bs As Boolean
Dim gs As Boolean
Dim cs As Boolean
Dim ks As Boolean
Dim ss As Boolean
Dim ns As Boolean
Dim CMs As Boolean

On Error Resume Next
PartColorDecRate = 0.2
PartColorDecRate2 = 0
ClearSurface = False
partrate2 = 0
'Evaluate command line arguments
For t = 1 To Len(Command)
    Select Case Mid$(Command, t, 2)
    Case "-f"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        Foyer = Val(Mid$(Command, t, x - t))
        t = x
        ps = True
    Case "-p"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        Maxparticul = Val(Mid$(Command, t, x - t))
        t = x
        ps = True
    Case "-b"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        BlurFactor = Val(Mid$(Command, t, x - t))
        t = x
        bs = True
    Case "-g"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        GravVY = Val(Mid$(Command, t, x - t))
        t = x
        gs = True
    Case "-y"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        Yoffset = Val(Mid$(Command, t, x - t))
        Yoffsetenabled = True
        t = x
        gs = True
    Case "-k"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        onclickPartnum = Val(Mid$(Command, t, x - t))
        t = x
        gs = True
    Case "-n"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        partrate = Val(Mid$(Command, t, x - t))
        t = x
        ns = True
    Case "-N"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        partrate2 = Val(Mid$(Command, t, x - t))
        t = x
        ns = True
    Case "-c"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        PGFrom = Val(Mid$(Command, t, x - t))
        t = x

        t = t + 1
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        b = Val(Mid$(Command, t, x - t))
        t = x
        PGVaray = b - PGFrom
        cs = True
        PGEnable = True
    Case "-C"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        Partcolorfrom = Val(Mid$(Command, t, x - t))
        t = x

        t = t + 1
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        b = Val(Mid$(Command, t, x - t))
        t = x
        partcolorVaray = b - Partcolorfrom
        CMs = True

        
    Case "-s"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        SpeedFrom = Val(Mid$(Command, t, x - t))
        t = x

        t = t + 1
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        b = Val(Mid$(Command, t, x - t))
        t = x
        SpeedVaray = b - SpeedFrom
        ss = True

    Case "-l"
        windenable = True
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        palletteind = Val(Mid$(Command, t, x - t))
        t = x
        bs = True
    Case "-w"
        windenable = True
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        windpower = Val(Mid$(Command, t, x - t))
        t = x
        bs = True
    Case "-d"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        PartColorDecRate = Val(Mid$(Command, t, x - t))
        t = x
    Case "-D"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        PartColorDecRate2 = Val(Mid$(Command, t, x - t))
        t = x
    Case "-r"
        t = t + 3
        x = t
        Do While Mid$(Command, x, 1) <> " "
            x = x + 1
            If x >= Len(Command) Then
                x = x + 1
                Exit Do
            End If
        Loop
        ClearSurface = IIf(Val(Mid$(Command, t, x - t)) = 1, True, False)
        t = x
    End Select
    
Next t

'Affect default if argument not found
If ps = False Then
    Maxparticul = 2000
End If
If bs = False Then
    BlurFactor = 1
End If
If gs = False Then
    GravVY = -0.01
End If
If cs = False Then
    PGVaray = 180
    PGFrom = 180
End If
If ss = False Then
    SpeedFrom = 0
    SpeedVaray = 1
End If
If ns = False Then
    partrate = 3
End If
If CMs = False Then
    partcolorVaray = 0
    Partcolorfrom = 255
End If

'initialize
ReDim part(Maxparticul)
ShowCursor False
Init
initpal
ReDim img((DDPIC.Width), (DDPIC.Height))
'Randomize the image
For x = 0 To (DDPIC.Width)
    For y = 0 To (DDPIC.Height)
        img(x, y) = Int(Rnd * 255)
    Next y
Next x

'If display is 16bit color: convert color to 16bit
f = displaymode.ddsCaps.lCaps
If displaymode.ddpfPixelFormat.lRGBBitCount = 16 Then
    ConvertPal32to16 pal()
End If

Me.Show
Initpart
InitTrigo
running = True
PTimecount = GetTickCount

'Main loop
Do While running
    animPart
    If windenable Then AnimWind

    DrawPart
    vbClearEdge DDPIC.Width, DDPIC.Height, img()    'Set to 0 edges of the image

    For t = 1 To BlurFactor
        vbBlurImg DDPIC.Width, DDPIC.Height, img()
    Next t
    render
    DoEvents

Loop
'Shutdown
ShowCursor True
End
End Sub
Public Function vbClearEdge(w As Long, h As Long, img() As Byte)
Dim x As Long, y As Long, add As Long
For y = 0 To h
    img(0, y) = 0
    img(w, y) = 0
Next y
For x = 0 To w
    img(x, 0) = 0
    img(x, h) = 0
Next x

End Function
Public Function vbBlurImg(w As Long, h As Long, img() As Byte)
'Blur the image
Dim x As Long, y As Long, add As Long, L1 As Long, L2 As Long, l3 As Long

For y = 1 To h - 1
        L2 = 0& + img(0, y - 1) + img(0, y) + img(0, y + 1)
        l3 = 0& + img(1, y - 1) + img(1, y) + img(1, y + 1)
    For x = 2 To w
        L1 = L2
        L2 = l3
        l3 = 0& + img(x, y - 1) + img(x, y) + img(x, y + 1)
        img(x - 1, y) = (L1 + L2 + l3) \ 9
    Next x
Next y


End Function


Public Function render()
'Render to the display
Dim tmpr As RECT
tmpr.Left = 0
tmpr.Top = 0
tmpr.Right = (DDPIC.Width) - 2
tmpr.Bottom = (DDPIC.Height) - 2

dd8b.bltToDDsurf_SafeBound img(), pal(), 0, 0, tmpr, DX7_BS
bltprim

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DDPIC_MouseDown Button, Shift, x, y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DDPIC_MouseMove Button, Shift, x, y

End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowCursor True
running = False
End Sub

