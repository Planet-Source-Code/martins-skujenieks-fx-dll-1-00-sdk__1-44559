VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Hello World"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   186
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox picSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Height          =   855
      Left            =   1080
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================================
' I added the Hello World project to FX.DLL 1.00 SDK for those who
' want to learn how to create 2D Engine with FX.DLL 1.00.
' I included very simple 2D Engine + mass of comments!
' I hope You will like it, and if You do, please vote for FX.DLL 1.00 SDK!!! :)
'
'BTW: add some background image, it will look better :)
'===============================================================================

Option Explicit

    Private bugX As Double
    Private bugY As Double
    
    Private tgtX As Long
    Private tgtY As Long
    Private tgtTrans As Long
    
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    
    
Private Sub Form_Load()

    '---------------------------------------------------------------
    'If You want that Your engine worked properly,
    'always be sure that:
    '
    '   for Destination PictureBox following properties are:
    '       AutoRedraw = True
    '       ScaleMode = vbPixels
    '
    '   for Source PictureBox following properties are:
    '       AutoRedraw = True
    '       AutoSize = True
    '       ScaleMode = vbPixels
    '       Visible = False
    '---------------------------------------------------------------
    
    Me.ScaleMode = vbPixels

    With picDest
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
        .AutoRedraw = True
        .ScaleMode = vbPixels
    End With

    With picSrc
        .AutoRedraw = True
        .AutoSize = True
        .ScaleMode = vbPixels
    End With
    
    
    'Load images:
    picSrc.Picture = LoadPicture("gfx.gif")
    
    
    'Show window:
    Me.Show
    DoEvents
    
    
    'Start 2D Engine:
    Engine
        
End Sub


Private Sub Engine()

    '---------------------------------------------------------------
    'Here is the very simple 2D Engine based on FX library.
    'Engine structure is based on one rendering loop in which
    'all graphics are drawn in specific order, preventing
    'flickering and showing good speed.
    '
    'The core of the engine always should look like this:
    '   1. Start Frame      begin loop, command "Do"
    '   2. DoEvents         allow system to complete other tasks,
    '                       call "DoEvents"
    '   3. Process Data     calculate new coordinates, colisions etc.
    '   4. Clear Surface    clear destination surface, call ".Cls"
    '   5. Render           draw all new objects on dest. surface
    '   6. Repaint          repaint destination surface,
    '                       call ".Refresh"
    '   7. End Frame        end loop, command "Loop"
    '
    'Of course You can add other stuff in this loop, this is just
    'basic to get into it!
    '---------------------------------------------------------------


    'Used to get framerate:
    Dim tStart As Long
    Dim tEnd As Long
    Dim tFPS As Long
    

    '1. Start next frame:
    Do
    
        '2. Allow system to complete all tasks:
            DoEvents
        
        
        '3. Calculate new coordinates, transparency values etc.
            If bugX < tgtX Then bugX = bugX + 1
            If bugX > tgtX Then bugX = bugX - 1
            If bugY < tgtY Then bugY = bugY + 1
            If bugY > tgtY Then bugY = bugY - 1
            
            tgtTrans = tgtTrans - 5
        
        
        '4. Clear destination surface:
            picDest.Cls
        
        
        '5. Draw all new objects to the destination surface:
            'Draw target:
            If tgtTrans >= 0 Then
                fxAlphaBlend picDest.hDC, tgtX, tgtY, 51, 64, picSrc.hDC, 0, 64, 51, 64, tgtTrans, vbGreen
            End If
            
            'Draw the bug:
            fxAlphaBlend picDest.hDC, bugX, bugY, 51, 64, picSrc.hDC, 0, 0, 51, 64, 255, vbGreen
            
            'Draw OSD:
            fxTextOut picDest.hDC, 10, 10, "Target X: " & tgtX, vbBlack, TA_LEFT
            fxTextOut picDest.hDC, 10, 22, "Target Y: " & tgtY, vbBlack, TA_LEFT
            
            tEnd = GetTickCount: tFPS = Int(1000 / (tEnd - tStart + 1))
            tStart = GetTickCount

            fxTextOut picDest.hDC, 10, 34, "FPS: " & tFPS, vbBlack, TA_LEFT
            
            
        '6. Repaint destination surface:
            picDest.Refresh
    
    
    '7. End frame, go to next frame :)
    Loop

End Sub


Private Sub Form_Resize()
    'Change destination PictureBox size if form size is changed:
    With picDest
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    DoEvents
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Exit all loops, stop all processes, etc.:
    End
End Sub


Private Sub picDest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Get new bug target coordinates:
    If Button = 1 Then
        tgtX = X - 25
        tgtY = Y - 32
        tgtTrans = 255
    End If

End Sub


'===============================================================================
'Now when You have gone through it You may try to upgrade this engine,
'then you can create Your own engines!
'I hope You liked!
'===============================================================================

