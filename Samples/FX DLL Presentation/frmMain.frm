VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FX.DLL Presentation"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picShow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   360
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   585
      TabIndex        =   112
      Top             =   720
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton Command9 
         Caption         =   "Toggle OSD"
         Height          =   375
         Left            =   2880
         TabIndex        =   120
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "+"
         Height          =   375
         Left            =   2160
         TabIndex        =   118
         Top             =   5400
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "-"
         Height          =   375
         Left            =   1680
         TabIndex        =   119
         Top             =   5400
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Stop"
         Height          =   375
         Left            =   4440
         TabIndex        =   117
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   116
         Top             =   5400
         Width           =   1335
      End
      Begin VB.PictureBox picBufferBase 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   615
         Left            =   8040
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   114
         Top             =   5280
         Visible         =   0   'False
         Width           =   615
         Begin VB.PictureBox picBuffer 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   2400
            Left            =   0
            Picture         =   "frmMain.frx":0000
            ScaleHeight     =   160
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   160
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   2400
         End
      End
      Begin VB.PictureBox picScreen 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   186
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   5325
         Left            =   120
         MouseIcon       =   "frmMain.frx":2515
         ScaleHeight     =   351
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   564
         TabIndex        =   113
         Top             =   0
         Width           =   8520
      End
   End
   Begin VB.Frame Frame25 
      Caption         =   "Credits"
      Height          =   2295
      Left            =   480
      TabIndex        =   121
      Top             =   4200
      Width           =   3495
      Begin VB.Label Label17 
         Caption         =   "Please vote at Planet-Source-Code!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         MouseIcon       =   "frmMain.frx":30D7
         MousePointer    =   99  'Custom
         TabIndex        =   123
         Top             =   900
         Width           =   3015
      End
      Begin VB.Label Label15 
         Caption         =   $"frmMain.frx":3229
         Height          =   495
         Left            =   240
         TabIndex        =   122
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.PictureBox picGeometry 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   109
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame24 
         Caption         =   "Rotation"
         Enabled         =   0   'False
         Height          =   855
         Left            =   0
         TabIndex        =   110
         Top             =   0
         Width           =   4815
         Begin VB.HScrollBar HScroll27 
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            Max             =   360
            Min             =   -360
            TabIndex        =   111
            Top             =   360
            Width           =   4335
         End
      End
   End
   Begin VB.PictureBox picFilters 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   100
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame23 
         Caption         =   "Replace Colors"
         Height          =   1455
         Left            =   0
         TabIndex        =   103
         Top             =   960
         Width           =   4815
         Begin VB.PictureBox picRepBy 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4320
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   108
            Top             =   1020
            Width           =   255
         End
         Begin VB.PictureBox picRep 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4320
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   107
            Top             =   720
            Width           =   255
         End
         Begin VB.HScrollBar HScroll26 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   104
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label14 
            Caption         =   "Replace By Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   1020
            Width           =   3855
         End
         Begin VB.Label Label13 
            Caption         =   "Replace Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   720
            Width           =   3975
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Reduce Colors"
         Height          =   855
         Left            =   0
         TabIndex        =   101
         Top             =   0
         Width           =   4815
         Begin VB.HScrollBar HScroll25 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   102
            Top             =   360
            Width           =   4335
         End
      End
   End
   Begin VB.PictureBox picDrawing 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   78
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame20 
         Caption         =   "3D Objects"
         Enabled         =   0   'False
         Height          =   855
         Left            =   0
         TabIndex        =   98
         Top             =   2880
         Width           =   4815
         Begin VB.Label Label12 
            Caption         =   "FX.DLL Version 1.00 Doesn't Support 3D Drawing"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "2D Objects"
         Height          =   2775
         Left            =   0
         TabIndex        =   79
         Top             =   0
         Width           =   4815
         Begin VB.OptionButton Option1 
            Caption         =   "RoundRect"
            Enabled         =   0   'False
            Height          =   375
            Index           =   17
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   2160
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rectangle"
            Enabled         =   0   'False
            Height          =   375
            Index           =   16
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   2160
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PolyPolyline"
            Enabled         =   0   'False
            Height          =   375
            Index           =   15
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   2160
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PolyPolygon"
            Enabled         =   0   'False
            Height          =   375
            Index           =   14
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PolylineTo"
            Enabled         =   0   'False
            Height          =   375
            Index           =   13
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Polyline"
            Enabled         =   0   'False
            Height          =   375
            Index           =   12
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Polygon"
            Enabled         =   0   'False
            Height          =   375
            Index           =   11
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PolyDraw"
            Enabled         =   0   'False
            Height          =   375
            Index           =   10
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PolyBezier To"
            Enabled         =   0   'False
            Height          =   375
            Index           =   9
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PolyBezier"
            Enabled         =   0   'False
            Height          =   375
            Index           =   8
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Pie"
            Enabled         =   0   'False
            Height          =   375
            Index           =   7
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "LineTo"
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ellipse"
            Height          =   375
            Index           =   5
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Circle"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Chord"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "ArcTo"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Arc"
            Height          =   375
            Index           =   1
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "AngleArc"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox picLight 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame19 
         Caption         =   "Fog"
         Height          =   1215
         Left            =   0
         TabIndex        =   74
         Top             =   2400
         Width           =   4815
         Begin VB.HScrollBar HScroll23 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   76
            Top             =   360
            Width           =   4335
         End
         Begin VB.PictureBox picFog2Color 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4320
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   75
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "Fog Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   720
            Width           =   3135
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Ambient Light"
         Height          =   2295
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   4815
         Begin VB.HScrollBar HScroll11 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   27
            Top             =   1800
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll10 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   26
            Top             =   1200
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll9 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   25
            Top             =   900
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll8 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   24
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label Label2 
            Caption         =   "Intensity:"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1560
            Width           =   4215
         End
         Begin VB.Label Label1 
            Caption         =   "RGB:"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   4215
         End
      End
   End
   Begin VB.PictureBox picEffects 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   65
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame18 
         Caption         =   "Fog"
         Height          =   1215
         Left            =   0
         TabIndex        =   70
         Top             =   2040
         Width           =   4815
         Begin VB.PictureBox picFogColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4320
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   73
            Top             =   720
            Width           =   255
         End
         Begin VB.HScrollBar HScroll16 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   71
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label10 
            Caption         =   "Fog Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   720
            Width           =   3135
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Blur"
         Height          =   975
         Left            =   0
         TabIndex        =   68
         Top             =   960
         Width           =   4815
         Begin VB.CommandButton Command4 
            Caption         =   "Blur"
            Height          =   375
            Left            =   240
            TabIndex        =   69
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Anti-Alias"
         Height          =   855
         Left            =   0
         TabIndex        =   66
         Top             =   0
         Width           =   4815
         Begin VB.HScrollBar HScroll24 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   67
            Top             =   360
            Width           =   4335
         End
      End
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   42
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame14 
         Caption         =   "HSL To RGB"
         Height          =   1695
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   4815
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   63
            Text            =   "0"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.HScrollBar HScroll18 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   62
            Top             =   960
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll19 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   61
            Top             =   660
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll20 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   60
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label7 
            Caption         =   "Resulting RGB Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   1320
            Width           =   2655
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "RGB"
         Height          =   1695
         Left            =   0
         TabIndex        =   53
         Top             =   4200
         Width           =   4815
         Begin VB.HScrollBar HScroll17 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   57
            Top             =   360
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll21 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   56
            Top             =   660
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll22 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   55
            Top             =   960
            Width           =   4335
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   54
            Text            =   "0"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Resulting RGB Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   1320
            Width           =   2655
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "RGB To HSL"
         Height          =   2295
         Left            =   0
         TabIndex        =   43
         Top             =   1800
         Width           =   4815
         Begin VB.HScrollBar HScroll13 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   49
            Top             =   360
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll14 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   48
            Top             =   660
            Width           =   4335
         End
         Begin VB.HScrollBar HScroll15 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   47
            Top             =   960
            Width           =   4335
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   46
            Text            =   "0"
            Top             =   1275
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   45
            Text            =   "0"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   44
            Text            =   "0"
            Top             =   1845
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Resulting Hue Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "Resulting Saturation Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   1590
            Width           =   2655
         End
         Begin VB.Label Label9 
            Caption         =   "Resulting Luminiscency Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   1845
            Width           =   2655
         End
      End
   End
   Begin VB.PictureBox picBitTrans 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame9 
         Caption         =   "BitBlt"
         Height          =   1095
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   4815
         Begin VB.CommandButton Command1 
            Caption         =   "BitBlt"
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "StretchBlt"
         Height          =   1095
         Left            =   0
         TabIndex        =   38
         Top             =   1200
         Width           =   4815
         Begin VB.CommandButton Command2 
            Caption         =   "StretchBlt"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "MaskBlt"
         Enabled         =   0   'False
         Height          =   975
         Left            =   0
         TabIndex        =   36
         Top             =   2400
         Width           =   4815
         Begin VB.CommandButton Command3 
            Caption         =   "MaskBlt"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "TransparentBlt"
         Height          =   1455
         Left            =   0
         TabIndex        =   31
         Top             =   3480
         Width           =   4815
         Begin VB.PictureBox picTransColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4320
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   15
            TabIndex        =   33
            Top             =   960
            Width           =   255
         End
         Begin VB.HScrollBar HScroll12 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   32
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label Label3 
            Caption         =   "Transparency:"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label4 
            Caption         =   "Transparent Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   3975
         End
      End
   End
   Begin VB.PictureBox picBlending 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame7 
         Caption         =   "Alpha Blending"
         Height          =   1215
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   4815
         Begin VB.PictureBox picTransparentColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4320
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   21
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Use Transparent Color"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   3975
         End
         Begin VB.HScrollBar HScroll7 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   19
            Top             =   360
            Width           =   4335
         End
      End
   End
   Begin VB.PictureBox picSrcBase 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   480
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   3495
      Begin VB.PictureBox picSrc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3750
         Left            =   0
         ScaleHeight     =   250
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   250
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   3750
      End
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3405
      Left            =   480
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   232
      TabIndex        =   14
      Top             =   720
      Width           =   3540
   End
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4200
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   1
      Top             =   720
      Width           =   4935
      Begin VB.Frame Frame6 
         Caption         =   "Saturation"
         Height          =   855
         Left            =   0
         TabIndex        =   12
         Top             =   4800
         Width           =   4815
         Begin VB.HScrollBar HScroll6 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   13
            Top             =   360
            Value           =   255
            Width           =   4335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Luminiscency"
         Enabled         =   0   'False
         Height          =   855
         Left            =   0
         TabIndex        =   10
         Top             =   3840
         Width           =   4815
         Begin VB.HScrollBar HScroll5 
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   -100
            TabIndex        =   11
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Hue"
         Height          =   855
         Left            =   0
         TabIndex        =   5
         Top             =   2880
         Width           =   4815
         Begin VB.HScrollBar HScroll4 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   -100
            TabIndex        =   9
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Gamma Ramp"
         Height          =   855
         Left            =   0
         TabIndex        =   4
         Top             =   1920
         Width           =   4815
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   240
            Max             =   1000
            TabIndex        =   8
            Top             =   360
            Value           =   100
            Width           =   4335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contrast"
         Enabled         =   0   'False
         Height          =   855
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   4815
         Begin VB.HScrollBar HScroll2 
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   7
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Brightness"
         Height          =   855
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4815
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   240
            Max             =   255
            Min             =   -255
            TabIndex        =   6
            Top             =   360
            Width           =   4335
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11668
      TabWidthStyle   =   2
      TabFixedWidth   =   1482
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Balance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bits"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Blending"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Colors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Drawing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Effects"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Geometry"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Light"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Show"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

    Private Type ShowObject
        X As Long
        Y As Long
        Vector As POINT
        Alive As Boolean
        Face As RECT
        Transparency As Long
    End Type
    
    Private StopShow As Boolean
    Private ShowStuff As Boolean
    Private ObjCount As Long
    
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


'===============================================================================
'   Bit-Block-Transfers
'===============================================================================
Private Sub Command1_Click()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxBitBlt picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picSrc.hDC, 0, 0, SRCCOPY
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "BitBlt (" & tEnd - tStart & " ms)"

End Sub

Private Sub Command2_Click()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxStretchBlt picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picSrc.hDC, 0, 0, 100, 100, SRCCOPY
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "StretchBlt (" & tEnd - tStart & " ms)"

End Sub

Private Sub Command4_Click()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxBlur picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Blur (" & tEnd - tStart & " ms)"

End Sub


'===============================================================================
'   2D Engine
'===============================================================================
Private Sub Command5_Click()


    'This is a sample of a simpliest 2d engine that
    'even beginner can create.
    'If you never did something like this, analize code -
    'it is very simple. You will cache it very fast :)
    
    Me.Caption = "2D Engine"
    
    '2D engine (named Show) is running:
    StopShow = False
    
    'Set default object (rock) count:
    ObjCount = 16

    'These all are used to get framerate:
    Dim tStart As Long
    Dim tEnd As Long
    Dim oFPS As Long    'old FPS
    Dim tFPS As Long    'current FPS
    Dim tFrames As Long 'current frame
    
    'Allocate memory for 16 objects:
    Dim Obj(16) As ShowObject



    'Setup:
    Dim id As Long
    For id = 0 To 16
        With Obj(id)
            'Specifies wether object is alive (not used in this engine):
            .Alive = True
            
            'Takes random face for object (from 16 aveilable faces in picBuffer):
            SetRect .Face, 40 * Int(Rnd(1) * 3), 40 * Int(Rnd(1) * 3), 40, 40
            
            'Set starting positions of the object:
            .X = Int(Rnd(1) * picScreen.ScaleWidth)
            .Y = Int(Rnd(1) * picScreen.ScaleHeight)
            
            'Set vector for the object
            .Vector.X = RndVect
            .Vector.Y = RndVect
            
            'Set transparency for the object (min=63):
            .Transparency = Int(Rnd(1) * 192) + 63
        End With
    Next
    
   
   
    'Rendering:
    Do
        If StopShow = True Then Exit Do
        
        DoEvents
        
        tStart = GetTickCount
        
        picScreen.Cls
              
        For id = 0 To ObjCount
            With Obj(id)
                If .Alive = True Then
                    If .X < 0 Then .Vector.X = RndVect
                    If .Y < 0 Then .Vector.Y = RndVect
                    If .X > picScreen.ScaleWidth - 40 Then .Vector.X = -RndVect
                    If .Y > picScreen.ScaleHeight - 40 Then .Vector.Y = -RndVect
                    .X = .X + .Vector.X
                    .Y = .Y + .Vector.Y
                    fxAlphaBlend picScreen.hDC, .X, .Y, 40, 40, picBuffer.hDC, .Face.Left, .Face.Top, .Face.Right, .Face.Bottom, .Transparency, vbMagenta
                End If
            End With
        Next
        
        If ShowStuff Then
            'Render osd:
            fxTextOut picScreen.hDC, 10, 10, "2D Engine", vbBlack, TA_LEFT
            fxTextOut picScreen.hDC, 10, 30, "Running " & tFPS & " FPS", vbBlack, TA_LEFT
            fxTextOut picScreen.hDC, 10, 42, "Frame " & tFrames, vbBlack, TA_LEFT
        End If
        
        picScreen.Refresh
        
        'Calculate framerate:
        tEnd = GetTickCount
        oFPS = tFPS
        tFPS = Int(oFPS + Int(1000 / (tEnd - tStart + 1))) / 2  'Get average FPS
        tFrames = tFrames + 1
        
    Loop
    
    
End Sub

Private Function RndVect() As Long
    Dim Vect As Long
    Vect = Int(Rnd(1) * 4)
    RndVect = Vect + 1
End Function

Private Sub Command6_Click()
    StopShow = True 'stop 2d engine
End Sub

Private Sub Command7_Click()
    ObjCount = ObjCount + 1
    If ObjCount > 16 Then ObjCount = 16
End Sub

Private Sub Command8_Click()
    ObjCount = ObjCount - 1
    If ObjCount < 0 Then ObjCount = 0
End Sub

Private Sub Command9_Click()
    ShowStuff = Not ShowStuff   'toggle 2d engine osd
End Sub

Private Sub Form_Load()
        
    picScreen.Picture = LoadPicture("gfx.jpg")
    DoEvents

    picDest.Cls
    fxBitBlt picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picScreen.hDC, 200, 0, SRCCOPY
    picDest.Refresh
    
    picSrc.Cls
    fxBitBlt picSrc.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picScreen.hDC, 300, 100, SRCCOPY
    picSrc.Refresh
    
    picDest.Picture = picDest.Image
    
    DoEvents
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopShow = True         'exit 2d engine, if it is running
    End                     'stop all stuff (i.e. 2d engine)
    Unload frmMain          'unload form
    Set frmMain = Nothing   'release grabbed resources
End Sub


'===============================================================================
'   Brightness
'===============================================================================
Private Sub HScroll1_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxBrightness picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, HScroll1.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Brightness: " & HScroll1.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub


'===============================================================================
'   TransparentBlt
'===============================================================================
Private Sub HScroll12_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxTransparentBlt picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picSrc.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picTransColor.BackColor, HScroll12.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "TransparentBlt: " & HScroll12.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll12_Scroll()
    HScroll12_Change
End Sub


'===============================================================================
'   RGB To HSL
'===============================================================================
Private Sub HScroll13_Change()
    
    Dim H As Double
    Dim S As Double
    Dim L As Double
    
    Dim RGBVal As Long: RGBVal = fxRGB(HScroll13.Value, HScroll14.Value, HScroll15.Value)
    
    fxRGBtoHSL RGBVal, H, S, L
    
    Text1.Text = Int(H * 100) / 100     'Return only two decimal numbers
    Text4.Text = Int(S * 100) / 100
    Text5.Text = Int(L * 100) / 100
    
    Text1.BackColor = RGBVal            'Set TextBox colors to specified RGB
    Text4.BackColor = RGBVal
    Text5.BackColor = RGBVal
    
    Me.Caption = "RGB To HSL Calculation"
    
End Sub

Private Sub HScroll13_Scroll()
    HScroll13_Change
End Sub

Private Sub HScroll14_Change()
    HScroll13_Change
End Sub

Private Sub HScroll14_Scroll()
    HScroll13_Change
End Sub

Private Sub HScroll15_Change()
    HScroll13_Change
End Sub

Private Sub HScroll15_Scroll()
    HScroll13_Change
End Sub


'===============================================================================
'   Fog
'===============================================================================
Private Sub HScroll16_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxFog picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picFogColor.BackColor, HScroll16.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Fog: " & HScroll16.Value & " (" & tEnd - tStart & " ms)"

End Sub

Private Sub HScroll16_Scroll()
    HScroll16_Change
End Sub


'===============================================================================
'   HSL To RGB
'===============================================================================
Private Sub HScroll18_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    Text2.Text = fxHSLtoRGB(HScroll20.Value / 100, HScroll19.Value / 100, HScroll18.Value / 100)
    Text2.BackColor = Val(Text2.Text)
    
    tEnd = GetTickCount
    
    Me.Caption = "HSL To RGB Calculation"

End Sub

Private Sub HScroll18_Scroll()
    HScroll18_Change
End Sub

Private Sub HScroll19_Change()
    HScroll18_Change
End Sub

Private Sub HScroll19_Scroll()
    HScroll18_Change
End Sub

Private Sub HScroll20_Change()
    HScroll18_Change
End Sub

Private Sub HScroll20_Scroll()
    HScroll18_Change
End Sub

Private Sub HScroll17_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    Text3.Text = fxRGB(HScroll17.Value, HScroll21.Value, HScroll22.Value)
    Text3.BackColor = Val(Text3.Text)
    
    tEnd = GetTickCount
    
    Me.Caption = "RGB Calculation"

End Sub

Private Sub HScroll17_Scroll()
    HScroll17_Change
End Sub

Private Sub HScroll21_Change()
    HScroll17_Change
End Sub

Private Sub HScroll21_Scroll()
    HScroll17_Change
End Sub

Private Sub HScroll22_Change()
    HScroll17_Change
End Sub

Private Sub HScroll22_Scroll()
    HScroll17_Change
End Sub


'===============================================================================
'   Fog
'===============================================================================
Private Sub HScroll23_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxFog picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picFog2Color.BackColor, HScroll23.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Fog: " & HScroll23.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll23_Scroll()
    HScroll23_Change
End Sub


'===============================================================================
'   Anti-Alias
'===============================================================================
Private Sub HScroll24_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxAntiAlias picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, HScroll24.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Anti-Alias: " & HScroll24.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll24_Scroll()
    HScroll24_Change
End Sub


Private Sub HScroll25_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxReduceColors picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, HScroll25.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Reduce Colors: " & HScroll25.Value & " (" & tEnd - tStart & " ms)"
    

End Sub

Private Sub HScroll25_Scroll()
    HScroll25_Change
End Sub


'===============================================================================
'   Replace Colors
'===============================================================================
Private Sub HScroll26_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxReplaceColors picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picRep.BackColor, picRepBy.BackColor, HScroll26.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Replace Colors: " & HScroll26.Value & " (" & tEnd - tStart & " ms)"
    

End Sub

Private Sub HScroll26_Scroll()
    HScroll26_Change
End Sub


'===============================================================================
'   Rotation
'===============================================================================
Private Sub HScroll27_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxRotate picDest.hDC, 50, 50, picDest.hDC, 50, 50, 150, 150, HScroll27.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Rotation: " & HScroll27.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll27_Scroll()
    HScroll27_Change
End Sub


'===============================================================================
'   Gamma Ramp
'===============================================================================
Private Sub HScroll3_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxGamma picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, HScroll3.Value / 100
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Gamma Ramp: " & HScroll3.Value / 100 & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll3_Scroll()
    HScroll3_Change
End Sub


'===============================================================================
'   Hue
'===============================================================================
Private Sub HScroll4_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxHue picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, HScroll4.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Hue: " & HScroll4.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll4_Scroll()
    HScroll4_Change
End Sub


'===============================================================================
'   Saturation
'===============================================================================
Private Sub HScroll6_Change()

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxSaturation picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, HScroll6.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Saturation: " & HScroll6.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll6_Scroll()
    HScroll6_Change
End Sub


'===============================================================================
'   Alpha Blending
'===============================================================================
Private Sub HScroll7_Change()
    
    Dim Flags As Long
    
    If Check1.Value = vbChecked Then
        Flags = picTransparentColor.BackColor
    Else
        Flags = -1
    End If

    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxAlphaBlend picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picSrc.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, HScroll7.Value, Flags
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Alpha Blending: " & HScroll7.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll7_Scroll()
    HScroll7_Change
End Sub

Private Sub Check1_Click()
    HScroll7_Change
End Sub


'===============================================================================
'   Ambient Light
'===============================================================================
Private Sub HScroll8_Change()
    
    Dim tStart As Long
    Dim tEnd As Long
    
    tStart = GetTickCount
    
    picDest.Cls
    fxAmbientLight picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hDC, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, fxRGB(HScroll8.Value, HScroll9.Value, HScroll10.Value), HScroll11.Value
    picDest.Refresh
    
    tEnd = GetTickCount
    
    Me.Caption = "Ambient Light: " & HScroll11.Value & " (" & tEnd - tStart & " ms)"
    
End Sub

Private Sub HScroll8_Scroll()
    HScroll8_Change
End Sub

Private Sub HScroll9_Change()
    HScroll8_Change
End Sub

Private Sub HScroll9_Scroll()
    HScroll8_Change
End Sub

Private Sub HScroll10_Change()
    HScroll8_Change
End Sub

Private Sub HScroll10_Scroll()
    HScroll8_Change
End Sub

Private Sub HScroll11_Change()
    HScroll8_Change
End Sub

Private Sub HScroll11_Scroll()
    HScroll8_Change
End Sub


Private Sub Label16_Click()
On Error Resume Next
    ShellExecute hwnd, "open", "http://www.exe.times.lv", vbNullString, App.Path, vbMaximizedFocus
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label17_Click()
On Error Resume Next
    ShellExecute hwnd, _
        "open", _
        "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=44559", _
        vbNullString, _
        App.Path, _
        vbMaximizedFocus
End Sub

'===============================================================================
'   Drawing 2D & 3D Objects
'===============================================================================
Private Sub Option1_Click(Index As Integer)

    'Get random object dimensions:
    Dim X1 As Long: X1 = Int(Rnd(1) * 100)
    Dim Y1 As Long: Y1 = Int(Rnd(1) * 100)
    Dim X2 As Long: X2 = Int(Rnd(1) * 100) + 100
    Dim Y2 As Long: Y2 = Int(Rnd(1) * 100) + 100
    Dim X3 As Long: X3 = Int(Rnd(1) * 100)
    Dim Y3 As Long: Y3 = Int(Rnd(1) * 100)
    Dim X4 As Long: X4 = Int(Rnd(1) * 100) + 100
    Dim Y4 As Long: Y4 = Int(Rnd(1) * 100) + 100
    
    picDest.Cls
    
    Select Case Index
        Case 1  'Arc
            fxArc picDest.hDC, X1, Y1, X2, Y2, X3, Y3, X4, Y4
            
        Case 5  'Ellipse
            fxEllipse picDest.hDC, X1, Y1, X2, Y2
            
        Case Else
    End Select
    
    picDest.Refresh
    
End Sub


'===============================================================================
'   Destination DC Events
'===============================================================================
Private Sub picDest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case TabStrip.SelectedItem.Caption
        Case "Blending"
            picTransparentColor.BackColor = picDest.POINT(X, Y): HScroll7_Change
        Case "Bits"
            picTransColor.BackColor = picDest.POINT(X, Y): HScroll12_Change
        Case "Effects"
            picFogColor.BackColor = picDest.POINT(X, Y): HScroll16_Change
        Case "Filters"
            Dim Col As Long: Col = picDest.POINT(X, Y)
            If Button = 1 Then
                picRep.BackColor = Col
            Else
                picRepBy.BackColor = Col
            End If
            HScroll26_Change
        Case "Light"
            picFog2Color.BackColor = picDest.POINT(X, Y): HScroll16_Change
        Case Else
    End Select
End Sub


'===============================================================================
'   TabStrip
'===============================================================================
Private Sub TabStrip_Click()
    Select Case TabStrip.SelectedItem.Caption
        Case "Balance"
            ResetTabs
            picBalance.Visible = True
            
        Case "Bits"
            ResetTabs
            picBitTrans.Visible = True
            
        Case "Blending"
            ResetTabs
            picBlending.Visible = True
                       
        Case "Colors"
            ResetTabs
            picColors.Visible = True
            
        Case "Drawing"
            ResetTabs
            picDrawing.Visible = True
            
        Case "Effects"
            ResetTabs
            picEffects.Visible = True
            
        Case "Filters"
            ResetTabs
            picFilters.Visible = True
            
        Case "Geometry"
            ResetTabs
            picGeometry.Visible = True
            
        Case "Light"
            ResetTabs
            picLight.Visible = True
            
        Case "Show"
            ResetTabs
            picShow.Visible = True
            
        Case Else
            ResetTabs
    End Select
End Sub

Private Sub ResetTabs()
    StopShow = True 'stop engine if it is running
    picBalance.Visible = False
    picBitTrans.Visible = False
    picBlending.Visible = False
    picColors.Visible = False
    picDrawing.Visible = False
    picEffects.Visible = False
    picFilters.Visible = False
    picGeometry.Visible = False
    picLight.Visible = False
    picShow.Visible = False
    DoEvents
End Sub
