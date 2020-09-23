VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "FX.DLL Demo - Window Effects"
   ClientHeight    =   1920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuRender 
         Caption         =   "Render"
         Begin VB.Menu mnuAlphaBlend 
            Caption         =   "Alpha Blend"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuAmbientLight 
            Caption         =   "Ambient Light"
         End
         Begin VB.Menu mnuBitBlt 
            Caption         =   "BitBlt"
         End
         Begin VB.Menu mnuBlur 
            Caption         =   "Blur"
         End
         Begin VB.Menu mnuBrightness 
            Caption         =   "Brightness"
         End
         Begin VB.Menu mnuGamma 
            Caption         =   "Gamma"
         End
         Begin VB.Menu mnuGreyscale 
            Caption         =   "Greyscale"
         End
         Begin VB.Menu mnuHue 
            Caption         =   "Hue"
         End
         Begin VB.Menu mnuInversion 
            Caption         =   "Inversion"
         End
         Begin VB.Menu mnuInvert 
            Caption         =   "Invert"
         End
         Begin VB.Menu mnuSaturation 
            Caption         =   "Saturation"
         End
         Begin VB.Menu mnuScanlines 
            Caption         =   "Scanlines"
         End
         Begin VB.Menu mnuSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrails 
            Caption         =   "Trails Effect"
         End
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCP 
         Caption         =   "Control Panel"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   

Private Sub Form_Load()
On Error Resume Next
    
    ' Initialize rendering
    With Render
        .Buffer = GetDC(GetDesktopWindow)
        .Color = vbBlack
        .Flags = 64
        .Mode = "fxAlphaBlend"
        Set .Window = Me
    End With
    
    Hook Me.hWnd
    
    DoEvents
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unhook Me.hWnd
End Sub


Private Sub ResetRenderOptions()
    mnuAlphaBlend.Checked = False
    mnuAmbientLight.Checked = False
    mnuBitBlt.Checked = False
    mnuBlur.Checked = False
    mnuBrightness.Checked = False
    mnuGamma.Checked = False
    mnuGreyscale.Checked = False
    mnuHue.Checked = False
    mnuInversion.Checked = False
    mnuInvert.Checked = False
    mnuSaturation.Checked = False
    mnuScanlines.Checked = False
End Sub

Private Sub mnuAlphaBlend_Click()
    ResetRenderOptions
    mnuAlphaBlend.Checked = True: Render.Mode = "fxAlphaBlend"
    Update
End Sub

Private Sub mnuAmbientLight_Click()
    ResetRenderOptions
    mnuAmbientLight.Checked = True: Render.Mode = "fxAmbientLight"
    Update
End Sub

Private Sub mnuBitBlt_Click()
    ResetRenderOptions
    mnuBitBlt.Checked = True: Render.Mode = "fxBitBlt"
    Update
End Sub

Private Sub mnuBkgClear_Click()
    Me.Picture = LoadPicture(vbNullString)
    Update
End Sub


Private Sub mnuBlur_Click()
    ResetRenderOptions
    mnuBlur.Checked = True: Render.Mode = "fxBlur"
    Update
End Sub

Private Sub mnuBrightness_Click()
    ResetRenderOptions
    mnuBrightness.Checked = True: Render.Mode = "fxBrightness"
    Update
End Sub

Private Sub mnuCP_Click()
    frmControlPanel.Show vbModal
End Sub

Private Sub mnuGamma_Click()
    ResetRenderOptions
    mnuGamma.Checked = True: Render.Mode = "fxGamma"
    Update
End Sub

Private Sub mnuGreyscale_Click()
    ResetRenderOptions
    mnuGreyscale.Checked = True: Render.Mode = "fxGreyscale"
    Update
End Sub

Private Sub mnuHue_Click()
    ResetRenderOptions
    mnuHue.Checked = True: Render.Mode = "fxHue"
    Update
End Sub

Private Sub mnuInversion_Click()
    ResetRenderOptions
    mnuInversion.Checked = True: Render.Mode = "fxInversion"
    Update
End Sub

Private Sub mnuInvert_Click()
    ResetRenderOptions
    mnuInvert.Checked = True: Render.Mode = "fxInvert"
    Update
End Sub

Private Sub mnuSaturation_Click()
    ResetRenderOptions
    mnuSaturation.Checked = True: Render.Mode = "fxSaturation"
    Update
End Sub

Private Sub mnuScanlines_Click()
    ResetRenderOptions
    mnuScanlines.Checked = True: Render.Mode = "fxScanlines"
    Update
End Sub

Private Sub mnuTrails_Click()
    mnuTrails.Checked = Not mnuTrails.Checked
    
    With Render
        .Flags = 127
        .Mode = "fxAlphaBlend": If mnuAlphaBlend.Checked = False Then mnuAlphaBlend_Click
        .Trails = mnuTrails.Checked
    End With
    
    mnuAlphaBlend.Enabled = Not mnuTrails.Checked
    mnuAmbientLight.Enabled = Not mnuTrails.Checked
    mnuBitBlt.Enabled = Not mnuTrails.Checked
    mnuBlur.Enabled = Not mnuTrails.Checked
    mnuBrightness.Enabled = Not mnuTrails.Checked
    mnuGamma.Enabled = Not mnuTrails.Checked
    mnuGreyscale.Enabled = Not mnuTrails.Checked
    mnuHue.Enabled = Not mnuTrails.Checked
    mnuInversion.Enabled = Not mnuTrails.Checked
    mnuInvert.Enabled = Not mnuTrails.Checked
    mnuSaturation.Enabled = Not mnuTrails.Checked
    mnuScanlines.Enabled = Not mnuTrails.Checked

End Sub
