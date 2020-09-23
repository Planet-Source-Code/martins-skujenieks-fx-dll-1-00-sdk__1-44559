Attribute VB_Name = "Main"

'   ========================================================================
'    Information
'   ========================================================================
'
'    Description:       FX.DLL Demo: Window Effects
'    Copyright:         Copyright © Martins Skujenieks 2003
'
'
'   ========================================================================
'    Warning!!!
'   ========================================================================
'
'    This module uses WindowProc(), so ALWAYS exit application by pressing
'    Close button on form, or Visaul Basic will crash!!!
'
'   ========================================================================


Option Explicit

    Private Type RENDERDESC
        Allowed As Boolean
        Buffer As Long
        Color As Long
        Flags As Long
        Mode As String
        Picture As String
        Trails As Boolean
        Updating As Boolean
        Window As Form
        WindowHor As Long
        WindowVer As Long
    End Type
    
    Public Render As RENDERDESC
    
    Public Const SM_CXSCREEN = 0           'X Size of screen
    Public Const SM_CYSCREEN = 1           'Y Size of Screen
    Public Const SM_CXVSCROLL = 2          'X Size of arrow in vertical scroll bar
    Public Const SM_CYHSCROLL = 3          'Y Size of arrow in horizontal scroll bar
    Public Const SM_CYCAPTION = 4          'Height of windows caption
    Public Const SM_CXBORDER = 5           'Width of non-sizable borders
    Public Const SM_CYBORDER = 6           'Height of non-sizable borders
    Public Const SM_CXDLGFRAME = 7         'Width of dialog box borders
    Public Const SM_CYDLGFRAME = 8         'Height of dialog box borders
    Public Const SM_CYMENU = 15            'Height of menu
              
    Public Const GWL_WNDPROC = (-4)
    
    Public Const WM_NULL = &H0
    Public Const WM_CREATE = &H1
    Public Const WM_DESTROY = &H2
    Public Const WM_MOVE = &H3
    Public Const WM_SIZE = &H5
    Public Const WM_ACTIVATE = &H6
    Public Const WM_SETFOCUS = &H7
    Public Const WM_KILLFOCUS = &H8
    Public Const WM_ENABLE = &HA
    Public Const WM_SETREDRAW = &HB
    Public Const WM_PAINT = &HF
    Public Const WM_CLOSE = &H10
    Public Const WM_QUIT = &H12
    
    Public PrevProc As Long
    
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Declare Function GetTickCount Lib "kernel32" () As Long
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
    Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
      
        
Public Sub Hook(hWnd As Long)
       
    With Render
        ' ====
        ' Get neccessayr System Metrics (like caption & menu height,
        ' border size etc. values):
        ' ====
        Select Case .Window.BorderStyle
            Case 0 'None
                .WindowHor = (GetSystemMetrics(SM_CXBORDER) * 2) + 1
                .WindowVer = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU) + (GetSystemMetrics(SM_CYBORDER) * 2) + 1
                
            Case 1 'Fixed Single
                .WindowHor = (GetSystemMetrics(SM_CXBORDER) * 2) + 1
                .WindowVer = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU) + (GetSystemMetrics(SM_CYBORDER) * 2) + 1
                
            Case 2 'Sizeable
                .WindowHor = (GetSystemMetrics(SM_CXBORDER) * 2) + 2
                .WindowVer = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU) + (GetSystemMetrics(SM_CYBORDER) * 2) + 2
                
            Case 3 'Dialog
                .WindowHor = (GetSystemMetrics(SM_CXDLGFRAME) * 2) - 3
                .WindowVer = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU) + (GetSystemMetrics(SM_CYDLGFRAME) * 2) - 3
        End Select
    End With
       
    ' ====
    ' Activate window hooking:
    ' ====
    PrevProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub


Public Sub Unhook(hWnd As Long)
    SetWindowLong hWnd, GWL_WNDPROC, PrevProc
End Sub


Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)
      
    Select Case uMsg
        Case WM_MOVE: Update
        Case WM_SIZE: Update
        Case WM_SETFOCUS: If Render.Updating = False Then Update
    End Select
            
End Function


Public Sub Update()

    With Render
                   
        .Updating = True
        DoEvents
    
        Dim MS As Long: MS = GetTickCount
        Dim mLeft As Long: mLeft = (.Window.Left / Screen.TwipsPerPixelX) + .WindowHor
        Dim mTop As Long: mTop = (.Window.Top / Screen.TwipsPerPixelY) + .WindowVer
        Dim mWidth As Long: mWidth = .Window.Width / Screen.TwipsPerPixelX
        Dim mHeight As Long: mHeight = .Window.Height / Screen.TwipsPerPixelY
    
        .Window.Visible = False: DoEvents
        
        If .Trails = False Then
            .Window.Cls
        End If
    
        Select Case .Mode
            Case "fxAlphaBlend": fxAlphaBlend .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight, .Flags, -1
            Case "fxAmbientLight": fxAmbientLight .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight, .Color, .Flags
            Case "fxBitBlt": fxBitBlt .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, SRCCOPY
            Case "fxBlur": fxBlur .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight
            Case "fxBrightness": fxBrightness .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight, .Flags
            Case "fxGamma": fxGamma .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight, .Flags
            Case "fxGreyscale": fxGreyscale .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight
            Case "fxHue": fxHue .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight, .Flags
            Case "fxInversion": fxInversion .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight, .Flags
            Case "fxInvert": fxInvert .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight
            Case "fxSaturation": fxSaturation .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight, .Flags
            Case "fxScanlines": fxScanlines .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, mWidth, mHeight, 1, 1, .Color, .Flags, True, False
            Case Else: fxBitBlt .Window.hDC, 0, 0, mWidth, mHeight, .Buffer, mLeft, mTop, SRCCOPY
        End Select
    
        .Window.Refresh
        
        If .Trails = True Then
            .Window.Picture = .Window.Image
            .Window.Refresh
        End If
        
        .Window.Visible = True ': DoEvents
    
        .Window.Caption = "Render: " & (GetTickCount - MS) & " Miliseconds"
        
        .Updating = False
        DoEvents
    
    End With
    
End Sub


