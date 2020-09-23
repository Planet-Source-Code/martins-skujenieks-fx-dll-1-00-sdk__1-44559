Attribute VB_Name = "FX"


'   ========================================================================
'    Module Information
'   ========================================================================
'
'    File Version:      1.00.0001
'    Description:       FX.DLL Procedure Declarations, Constant and Tag
'                       Definitions
'    Copyright:         Copyright © Martins Skujenieks 2002-2003
'    Product Name:      FX.DLL
'    Product Version:   1.00
'
'
'   ========================================================================
'    End User License Agreement (EULA)
'   ========================================================================
'
'    This product is provided "as is", with no guarantee of completeness or
'    accuracy and without warranty of any kind, express or implied.
'
'    In no event will developer be liable for damages of any kind that may
'    be incurred with your hardware, peripherals or software programs.
'
'    You may create one copy of the product on any single computer for
'    your personal, non-commercial, home use only, provided you keep intact
'    all copyright and other proprietary notices.
'
'    This product and all of its parts may not be copied, emulated, cloned,
'    rented, leased, sold, reproduced, modified, decompiled, disassembled,
'    otherwise reverse engineered, republished, uploaded, posted,
'    transmitted or distributed in any way, without prior written consent
'    of the developer.
'
'
'   ========================================================================
'    Contact Developer
'   ========================================================================
'
'    Website:           http://www.exe.times.lv
'    E-Mail:            exe.times@e-apollo.lv
'
'   ========================================================================


Option Explicit


    '/* Ternary Raster Operations */
    Public Const SRCCOPY = &HCC0020
    Public Const SRCPAINT = &HEE0086
    Public Const SRCAND = &H8800C6
    Public Const SRCINVERT = &H660046
    Public Const SRCERASE = &H440328
    Public Const NOTSRCCOPY = &H330008
    Public Const NOTSRCERASE = &H1100A6
    Public Const MERGECOPY = &HC000CA
    Public Const MERGEPAINT = &HBB0226
    Public Const PATCOPY = &HF00021
    Public Const PATPAINT = &HFB0A09
    Public Const PATINVERT = &H5A0049
    Public Const DSTINVERT = &H550009
    Public Const BLACKNESS = &H42
    Public Const WHITENESS = &HFF0062


    '/* StretchBlt() Modes */
    Public Const BLACKONWHITE = 1
    Public Const WHITEONBLACK = 2
    Public Const COLORONCOLOR = 3
    Public Const HALFTONE = 4
    Public Const MAXSTRETCHBLTMODE = 4


    '/* New StretchBlt() Modes */
    Public Const STRETCH_ANDSCANS = BLACKONWHITE
    Public Const STRETCH_ORSCANS = WHITEONBLACK
    Public Const STRETCH_DELETESCANS = COLORONCOLOR
    Public Const STRETCH_HALFTONE = HALFTONE


    '/* Text Alignment Options */
    Public Const TA_NOUPDATECP = 0
    Public Const TA_UPDATECP = 1
    Public Const TA_LEFT = 0
    Public Const TA_RIGHT = 2
    Public Const TA_CENTER = 6
    Public Const TA_TOP = 0
    Public Const TA_BOTTOM = 8
    Public Const TA_BASELINE = 24
    Public Const TA_RTLREADING = 256
    Public Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP + TA_RTLREADING)
    
    
    '/* Vertical Text Alignment Options */
    Public Const VTA_BASELINE = TA_BASELINE
    Public Const VTA_LEFT = TA_BOTTOM
    Public Const VTA_RIGHT = TA_TOP
    Public Const VTA_CENTER = TA_CENTER
    Public Const VTA_BOTTOM = TA_RIGHT
    Public Const VTA_TOP = TA_LEFT


    '/* struct tagPOINT */
    Public Type POINT
        X       As Long
        Y       As Long
    End Type

    
    '/* struct tagRECT */
    Public Type RECT
        Left    As Long
        Top     As Long
        Right   As Long
        Bottom  As Long
    End Type


    Public Declare Function fxAlphaBlend Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Alpha As Long, ByVal TransparentColor As Long) As Long
    Public Declare Function fxAmbientLight Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Color As Long, ByVal ensity As Long) As Long
    Public Declare Function fxAntiAlias Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Purity As Long) As Long
    Public Declare Function fxArc Lib "fx.dll" (ByVal DC As Long, ByVal LeftRect As Long, ByVal TopRect As Long, ByVal RightRect As Long, ByVal BottomRect As Long, ByVal XStartArc As Long, ByVal YStartArc As Long, ByVal XEndArc As Long, ByVal YEndArc As Long) As Long
    Public Declare Function fxBitBlt Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal RasterOperation As Long) As Long
    Public Declare Function fxBlur Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long) As Long
    Public Declare Function fxBrightness Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Brightness As Long) As Long
    Public Declare Function fxCMYK Lib "fx.dll" (ByVal C As Long, ByVal M As Long, ByVal Y As Long, ByVal K As Long) As Long
    Public Declare Function fxEllipse Lib "fx.dll" (ByVal DC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
    Public Declare Function fxExpand Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Horizontal As Long, ByVal Vertical As Long) As Long
    Public Declare Function fxFilter Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Filter As Long, ByVal Flags As Long) As Long
    Public Declare Function fxFog Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Color As Long, ByVal ensity As Long) As Long
    Public Declare Function fxGamma Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Gamma As Double) As Long
    Public Declare Function fxGetBlue Lib "fx.dll" (ByVal RGB As Long) As Long
    Public Declare Function fxGetGreen Lib "fx.dll" (ByVal RGB As Long) As Long
    Public Declare Function fxGetRed Lib "fx.dll" (ByVal RGB As Long) As Long
    Public Declare Function fxGreyscale Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long) As Long
    Public Declare Function fxGridelines Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Color As Long, ByVal Transparency As Long, ByVal Distance As Long, ByVal Flags As Long) As Long
    Public Declare Function fxHSLtoRGB Lib "fx.dll" (ByVal H As Double, ByVal S As Double, ByVal L As Double) As Long
    Public Declare Function fxHue Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Hue As Long) As Long
    Public Declare Function fxInversion Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Inversion As Long) As Long
    Public Declare Function fxInvert Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long) As Long
    Public Declare Function fxPaletteIndex Lib "fx.dll" (byvalIndex As Long) As Long
    Public Declare Function fxPaletteRGB Lib "fx.dll" (ByVal R As Long, ByVal G As Long, ByVal B As Long) As Long
    Public Declare Function fxReduceColors Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Level As Long) As Long
    Public Declare Function fxReplaceColor Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Color As Long, ByVal ByColor As Long) As Long
    Public Declare Function fxReplaceColors Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Color As Long, ByVal ByColor As Long, ByVal Similarity As Long) As Long
    Public Declare Function fxRGB Lib "fx.dll" (ByVal R As Long, ByVal G As Long, ByVal B As Long) As Long
    Public Declare Function fxRGBtoHSL Lib "fx.dll" (ByVal RGB As Long, ByRef H As Double, ByRef S As Double, ByRef L As Double) As Long
    Public Declare Function fxRotate Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Angle As Long) As Long
    Public Declare Function fxSaturation Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Saturation As Long) As Long
    Public Declare Function fxScanlines Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Space As Long, ByVal Thickness As Long, ByVal Color As Long, ByVal Transparency As Long, ByVal Horizontal As Boolean, ByVal Vertical As Boolean) As Long
    Public Declare Function fxScreenShot Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal GetCursor As Boolean) As Long
    Public Declare Function fxShadeColors Lib "fx.dll" (ByVal DestColor As Long, ByVal SrcColor As Long, ByVal Shade As Long) As Long
    Public Declare Function fxStretchBlt Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal RasterOperation As Long) As Long
    Public Declare Function fxStretchBltMode Lib "fx.dll" (ByVal DC As Long, ByVal Mode As Long) As Long
    Public Declare Function fxTextOut Lib "fx.dll" (ByVal DC As Long, ByVal X As Long, ByVal Y As Long, ByVal Text As String, ByVal Color As Long, ByVal Alignment As Long) As Long
    Public Declare Function fxTransparentBlt Lib "fx.dll" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Color As Long, ByVal Transparency As Long) As Long




