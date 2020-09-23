Attribute VB_Name = "Module1"
Private Type Lset32L
    c As Long
End Type
Private Type Lset32b
    b0 As Byte
    b1 As Byte
    b2 As Byte
    b3 As Byte
End Type
Global DX7_DX As New DirectX7
Global DX7_DD As DirectDraw7
Global DX7_PS As DirectDrawSurface7
Global DX7_BS As DirectDrawSurface7
Global DX7_BSR As RECT
Global DDC As New ddclass
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Global displaymode As DDSURFACEDESC2

Public Declare Function GetTickCount Lib "kernel32" () As Long




Public Function Init()
DDC.Init Form1.DDPIC.hwnd

Set DX7_PS = DDC.mk_prim_surf(Form1.DDPIC.hwnd)
Set DX7_BS = DDC.mk_surf(DX7_BSR, "", Form1.DDPIC.Width, Form1.DDPIC.Height)
DDC.dd.GetDisplayMode displaymode

End Function

Public Function bltprim()
Dim tmpr As RECT
DX7_DX.GetWindowRect Form1.DDPIC.hwnd, tmpr
DX7_PS.Blt tmpr, DX7_BS, DX7_BSR, DDBLT_WAIT
End Function
Public Sub Main()
Form1.Show

End Sub

Public Function Convert32bto16b(c As Long) As Long
Dim tmp1 As Lset32L
Dim tmp2 As Lset32b
tmp1.c = c
LSet tmp2 = tmp1
Convert32bto16b = (tmp2.b0 And &HF8) \ 8 + _
                (tmp2.b1 And &HFC) * 8 + _
                ((tmp2.b2 And &HF8) / 8 * 2048)
End Function
Public Function ConvertPal32to16(pal() As Long)
For t = 0 To 255
    pal(t) = Convert32bto16b(pal(t))
Next t
End Function
