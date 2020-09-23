VERSION 5.00
Begin VB.Form frmControlPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Panel"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2760
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Render"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5175
      Begin VB.PictureBox picCP 
         AutoRedraw      =   -1  'True
         Height          =   795
         Left            =   3840
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   69
         TabIndex        =   12
         Top             =   2880
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   11
         Top             =   3420
         Width           =   3495
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   10
         Top             =   3150
         Width           =   3495
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   9
         Top             =   2880
         Width           =   3495
      End
      Begin VB.PictureBox picP 
         AutoRedraw      =   -1  'True
         Height          =   1935
         Left            =   3840
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   69
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   5
         Top             =   2280
         Width           =   3495
      End
      Begin VB.ListBox List2 
         Height          =   1230
         ItemData        =   "frmControlPanel.frx":0000
         Left            =   240
         List            =   "frmControlPanel.frx":002B
         TabIndex        =   3
         Top             =   600
         Width           =   3495
      End
      Begin VB.PictureBox picP2 
         AutoRedraw      =   -1  'True
         Height          =   1935
         Left            =   3840
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   69
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Color (RGB):"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Preview:"
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Flags:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "Mode:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton Command4 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Background Picture:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


    Private OldColor As Long
    Private OldFlags As Long
    Private OldMode As String


    '====
    'Stuff for Common Dialog:
    '====
    
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
    
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

    

Private Sub Command1_Click()

    With Render
        .Flags = HScroll1.Value
        .Color = fxRGB(HScroll2.Value, HScroll3.Value, HScroll4.Value)
        .Mode = List2.List(List2.ListIndex)
        frmMain.Picture = picP.Picture
    End With
       
    Unload Me
    
    Update
    
End Sub

Private Sub Command2_Click()
    Render.Color = OldColor
    Render.Flags = OldFlags
    Render.Mode = OldMode
    Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
    picP.Picture = LoadPicture(vbNullString)
    HScroll1_Scroll
End Sub

Private Sub Command4_Click()
On Error Resume Next

    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    'OFName.hwndOwner = Me.hWnd          'Set the parent window
    OFName.hInstance = App.hInstance    'Set the application's instance
    'OFName.lpstrFilter = "Bitmap Images (*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "JPEG Images (*.jpg)" + Chr$(0) + "*.jpg" + Chr$(0) + "GIF Images (*.gif)" + Chr$(0) + "*.gif" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFile = Space$(254)      'Create a buffer for the file
    OFName.nMaxFile = 255               'Set the maximum length of a returned file name
    OFName.lpstrFileTitle = Space$(254) 'Create a buffer for the file title
    OFName.nMaxFileTitle = 255          'Set the maximum length of a returned file title
    'OFName.lpstrInitialDir = App.Path   'Set the initial directory
    OFName.lpstrTitle = "Load Background Image"    'Set the title
    OFName.Flags = 0                    'No flags

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        picP.Picture = LoadPicture(Trim$(OFName.lpstrFile))
        HScroll1_Scroll
    Else
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next

    OldColor = Render.Color
    OldFlags = Render.Flags
    OldMode = Render.Mode
    
    HScroll1.Value = OldFlags
    HScroll2.Value = fxGetRed(OldColor)
    HScroll3.Value = fxGetGreen(OldColor)
    HScroll4.Value = fxGetBlue(OldColor)
    DoEvents

    picP.Picture = frmMain.Picture
    
    List2.ListIndex = 0
    
    picP2.Cls
    fxBitBlt picP2.hDC, 0, 0, 70, 125, Render.Buffer, 0, 0, SRCCOPY
    picP2.Refresh
    
    DoEvents
    
End Sub

Private Sub HScroll1_Change()
    HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()

    With Render
        
        .Flags = HScroll1.Value
        .Color = fxRGB(HScroll2.Value, HScroll3.Value, HScroll4.Value)
    
        picCP.Cls
        picCP.BackColor = Render.Color
        picCP.Refresh
        
        picP.Cls
        
        Select Case List2.List(List2.ListIndex)
            Case "fxAlphaBlend": fxAlphaBlend picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125, .Flags, -1
            Case "fxAmbientLight": fxAmbientLight picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125, .Color, .Flags
            Case "fxBitBlt": fxBitBlt picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, SRCCOPY
            Case "fxBlur": fxBlur picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125
            Case "fxBrightness": fxBrightness picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125, .Flags
            Case "fxGamma": fxGamma picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125, .Flags
            Case "fxGreyscale": fxGreyscale picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125
            Case "fxHue": fxHue picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125, .Flags
            Case "fxInversion": fxInversion picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125, .Flags
            Case "fxInvert": fxInvert picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125
            Case "fxSaturation": fxSaturation picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125, .Flags
            Case "fxScanlines": fxScanlines picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125, 1, 1, .Color, .Flags, True, False
            Case Else: fxBitBlt picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, SRCCOPY
        End Select
        
        picP.Refresh
        
    End With
    
End Sub

Private Sub HScroll2_Change()
    HScroll1_Scroll
End Sub

Private Sub HScroll2_Scroll()
    HScroll1_Scroll
End Sub

Private Sub HScroll3_Change()
    HScroll1_Scroll
End Sub

Private Sub HScroll3_Scroll()
    HScroll1_Scroll
End Sub

Private Sub HScroll4_Change()
    HScroll1_Scroll
End Sub

Private Sub HScroll4_Scroll()
    HScroll1_Scroll
End Sub

Private Sub List2_Click()
    HScroll1_Scroll
        Select Case List2.List(List2.ListIndex)
            Case "fxAlphaBlend": HScroll1.Min = 0: HScroll1.Max = 255
            Case "fxAmbientLight": HScroll1.Min = 0: HScroll1.Max = 255
            Case "fxBitBlt": HScroll1.Min = 0: HScroll1.Max = 255
            Case "fxBlur": HScroll1.Min = 0: HScroll1.Max = 255
            Case "fxBrightness": HScroll1.Min = -255: HScroll1.Max = 255
            Case "fxGamma": HScroll1.Min = 0: HScroll1.Max = 100
            Case "fxGreyscale": HScroll1.Min = 0: HScroll1.Max = 255
            Case "fxHue": HScroll1.Min = -100: HScroll1.Max = 100
            Case "fxInversion": HScroll1.Min = 0: HScroll1.Max = 255
            Case "fxInvert": fxInvert picP.hDC, 0, 0, 70, 125, picP2.hDC, 0, 0, 70, 125
            Case "fxSaturation": HScroll1.Min = 0: HScroll1.Max = 255
            Case "fxScanlines": HScroll1.Min = 0: HScroll1.Max = 255
            Case Else: HScroll1.Min = 0: HScroll1.Max = 255
        End Select
End Sub

Private Sub picP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ptCol As Long: ptCol = picP.POINT(X, Y)
    HScroll2.Value = fxGetRed(ptCol)
    HScroll3.Value = fxGetGreen(ptCol)
    HScroll4.Value = fxGetBlue(ptCol)
End Sub
