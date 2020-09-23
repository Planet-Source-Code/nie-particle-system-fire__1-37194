VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ddclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public dx As New DirectX7
Public dd As DirectDraw7

Friend Function Init(hdl As Long, Optional flags As CONST_DDSCLFLAGS = DDSCL_NORMAL, Optional fullscr As Boolean = False)
Set dd = dx.DirectDrawCreate("")
dd.SetCooperativeLevel hdl, flags

If fullscr Then
'dd.SetDisplayMode 320, 240, 32, 0, DDSDM_DEFAULT
End If

End Function
Friend Function mk_prim_surf(hwnd As Long) As DirectDrawSurface7
Dim desc As DDSURFACEDESC2, clip As DirectDrawClipper
desc.lFlags = DDSD_CAPS
desc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
Set mk_prim_surf = dd.CreateSurface(desc)
Set clip = dd.CreateClipper(0)
clip.SetHWnd hwnd
mk_prim_surf.SetClipper clip
End Function

Friend Function mk_surf(ByRef r As RECT, Optional file As String = "", Optional ByVal w As Long = 0, Optional ByVal h As Long = 0, Optional f As CONST_DDSURFACECAPSFLAGS = 64, Optional colorkey As Boolean = False, Optional colorkeyhigh As Long = 0, Optional colorkeylow As Long = 0, Optional ckf As CONST_DDCKEYFLAGS = DDCKEY_SRCBLT) As DirectDrawSurface7
Dim desc As DDSURFACEDESC2
If w <> 0 And h <> 0 Then
    desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    desc.lWidth = w
    desc.lHeight = h
    With r
    .Bottom = h
    .Right = w
    End With
    
Else
    desc.lFlags = DDSD_CAPS
End If
desc.ddsCaps.lCaps = f
desc.ddpfPixelFormat.lRGBBitCount = 8
        Dim c As DDCOLORKEY
If file = "" Then
    Set mk_surf = dd.CreateSurface(desc)
    If colorkey = True Then

        c.high = colorkeyhigh
        c.low = colorkeylow
        mk_surf.SetColorKey ckf, c
    End If
Else
    Set mk_surf = dd.CreateSurfaceFromFile(file, desc)
    
    If colorkey = True Then

        c.high = colorkeyhigh
        c.low = colorkeylow
        mk_surf.SetColorKey ckf, c
    End If

    mk_surf.GetSurfaceDesc desc
    
    With r
    .Bottom = desc.lHeight
    .Right = desc.lWidth
    End With
End If
End Function
