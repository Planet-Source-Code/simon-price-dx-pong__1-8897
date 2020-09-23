Attribute VB_Name = "ModSurfaces"
Public BackBuffer As DirectDrawSurface7
Public View As DirectDrawSurface7
Public Balls As DirectDrawSurface7
Public Paddles As DirectDrawSurface7
Public Table As DirectDrawSurface7
Public Phrases As DirectDrawSurface7

Public ViewDesc As DDSURFACEDESC2
Public BackBufferDesc As DDSURFACEDESC2
Public BallsDesc As DDSURFACEDESC2
Public PaddlesDesc As DDSURFACEDESC2
Public TableDesc As DDSURFACEDESC2
Public PhrasesDesc As DDSURFACEDESC2

Public BackBufferCaps As DDSCAPS2

Public ColorKey As DDCOLORKEY

Sub CreatePrimaryAndBackBuffer()
Set View = Nothing
Set BackBuffer = Nothing

ViewDesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
ViewDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
ViewDesc.lBackBufferCount = 1
Set View = DX_Draw.CreateSurface(ViewDesc)

BackBufferCaps.lCaps = DDSCAPS_BACKBUFFER
Set BackBuffer = View.GetAttachedSurface(BackBufferCaps)
BackBuffer.GetSurfaceDesc ViewDesc

BackBuffer.SetFontTransparency True

End Sub

Sub LoadAllPics()
Dim Path As String

CreatePrimaryAndBackBuffer

Set Balls = Nothing
Set Paddles = Nothing
Set Table = Nothing

ModDX7.CreateSurfaceFromFile Balls, BallsDesc, App.Path & "\Graphics\Balls.bmp", 600, 50
ModDX7.CreateSurfaceFromFile Paddles, PaddlesDesc, App.Path & "\Graphics\Batz.bmp", 50, 100
ModDX7.CreateSurfaceFromFile Table, TableDesc, App.Path & "\Graphics\Table.bmp", 320, 240
ModDX7.CreateSurfaceFromFile Phrases, PhrasesDesc, App.Path & "\Graphics\Phrases.bmp", 200, 60

ModDX7.AddColorKey BackBuffer, ColorKey, vbWhite, vbWhite
ModDX7.AddColorKey Balls, ColorKey, vbWhite, vbWhite
ModDX7.AddColorKey Phrases, ColorKey, vbWhite, vbWhite

End Sub

