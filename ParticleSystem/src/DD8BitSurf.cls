VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DD8BitSurf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'* DD8bitsurf:Methods to use an array as a 8bits picture                      *
'******************************************************************************
'*Author: Stéphanie Rancourt                                                  *
'*Contact:asuka_tetsuo@hotmail.com                                            *
'*Date:20/04/2002                                                             *
'*Version: beta 1                                                             *
'*Note: 16bit mode not supported completely                                   *
'*Comment from the author: You may use or alter this code for your personal   *
'*  use.If you do please credit the author.You can't sell or use it for       *
'*  commercial purpose without the approval of the author but you may         *
'*  distribute it in it's integrity.                                          *
'******************************************************************************
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Friend Function bltToDDsurf_SafeBound(img() As Byte, pal() As Long, ByVal X As Long, ByVal Y As Long, SRC_Rect As RECT, ddsurface As DirectDrawSurface7, Optional trans As Boolean = False, Optional TcolorIndex As Long = 0) As Long
Dim xs As Long, ys As Long, xd As Long, yd As Long
Dim ubx As Long, uby As Long, tmp As Long
Dim tmpdesc As DDSURFACEDESC2
ddsurface.GetSurfaceDesc tmpdesc
ubx = UBound(img(), 1)
uby = UBound(img(), 2)
If SRC_Rect.Right > ubx Then SRC_Rect.Right = ubx
If SRC_Rect.Left > ubx Then SRC_Rect.Left = ubx
If SRC_Rect.Bottom > uby Then SRC_Rect.Bottom = uby
If SRC_Rect.Top > uby Then SRC_Rect.Top = uby
If SRC_Rect.Top < 0 Then SRC_Rect.Top = 0
If SRC_Rect.Left < 0 Then SRC_Rect.Left = 0
If SRC_Rect.Right < 0 Then SRC_Rect.Right = 0
If SRC_Rect.Bottom < 0 Then SRC_Rect.Bottom = 0
If SRC_Rect.Bottom < SRC_Rect.Top Then
    tmp = SRC_Rect.Top
    SRC_Rect.Top = SRC_Rect.Bottom
    SRC_Rect.Bottom = tmp
End If
If SRC_Rect.Right < SRC_Rect.Left Then
    tmp = SRC_Rect.Left
    SRC_Rect.Left = SRC_Rect.Right
    SRC_Rect.Right = tmp
End If
If SRC_Rect.Bottom = SRC_Rect.Top Then bltToDDsurf_SafeBound = 1: Exit Function
If SRC_Rect.Left = SRC_Rect.Right Then bltToDDsurf_SafeBound = 1: Exit Function
If X + (SRC_Rect.Right - SRC_Rect.Left) > tmpdesc.lWidth Then
    SRC_Rect.Right = tmpdesc.lWidth - X
End If
If Y + (SRC_Rect.Bottom - SRC_Rect.Top) > tmpdesc.lHeight Then
    SRC_Rect.Bottom = tmpdesc.lHeight - Y
End If
If X < 0 Then
    SRC_Rect.Left = -X
End If
If Y < 0 Then
    SRC_Rect.Top = -Y
End If

Select Case trans
Case False
    ddsurface.Lock SRC_Rect, tmpdesc, DDLOCK_WAIT, 0
    xd = X
    For xs = SRC_Rect.Left To SRC_Rect.Right
        yd = Y
        For ys = SRC_Rect.Top To SRC_Rect.Bottom
            ddsurface.SetLockedPixel xd, yd, pal(img(xs, ys))
            yd = yd + 1
        Next ys
        xd = xd + 1
    Next xs
    ddsurface.Unlock SRC_Rect
Case True
    ddsurface.Lock SRC_Rect, tmpdesc, DDLOCK_WAIT, 0
    xd = X
    For xs = SRC_Rect.Left To SRC_Rect.Right
        yd = Y
        For ys = SRC_Rect.Top To SRC_Rect.Bottom
            tmp = img(xs, ys)
            If tmp <> TcolorIndex Then
                ddsurface.SetLockedPixel xd, yd, pal(tmp)
            End If
            yd = yd + 1
        Next ys
        xd = xd + 1
    Next xs
    ddsurface.Unlock SRC_Rect
End Select
End Function
Friend Function Load_MPAL(Filename As String, pal() As Long)
Dim pos As Long, t As Long, r As Byte, g As Byte, b As Byte
pos = 25
Open Filename For Binary As #1
For t = 0 To 255
    Get #1, pos, r
    pos = pos + 1
    Get #1, pos, g
    pos = pos + 1
    Get #1, pos, b
    pos = pos + 2
    
     pal(t) = RGB(b, g, r)

Next t
Close #1
End Function
