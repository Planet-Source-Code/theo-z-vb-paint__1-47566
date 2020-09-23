VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaint 
   AutoRedraw      =   -1  'True
   Caption         =   "VB Paint"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   720
   ClientWidth     =   7440
   Icon            =   "frmPaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   7440
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5400
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   65
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picImageEffect 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4560
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraTools 
      Height          =   6525
      Left            =   0
      TabIndex        =   30
      Top             =   -90
      WhatsThisHelpID =   10296
      Width           =   855
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   16
         Left            =   50
         Picture         =   "frmPaint.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Brush"
         Top             =   3120
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   17
         Left            =   435
         Picture         =   "frmPaint.frx":11A6
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Hand"
         Top             =   3120
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   15
         Left            =   435
         Picture         =   "frmPaint.frx":18A8
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Zoom"
         Top             =   120
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   14
         Left            =   435
         Picture         =   "frmPaint.frx":1C36
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Filter Brush"
         Top             =   1245
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   13
         Left            =   435
         Picture         =   "frmPaint.frx":1CA8
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Curve"
         Top             =   2745
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   12
         Left            =   435
         Picture         =   "frmPaint.frx":1D00
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Polygon"
         Top             =   2370
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   11
         Left            =   435
         Picture         =   "frmPaint.frx":1D73
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Rounded Rectangle"
         Top             =   1995
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   10
         Left            =   50
         Picture         =   "frmPaint.frx":1DFD
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Air Brush"
         Top             =   1245
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   2
         Left            =   435
         Picture         =   "frmPaint.frx":2107
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Eraser"
         Top             =   870
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10295
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   4
         Left            =   50
         Picture         =   "frmPaint.frx":2186
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Pencil"
         Top             =   870
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         WhatsThisHelpID =   10298
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   5
         Left            =   50
         Picture         =   "frmPaint.frx":2205
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Line"
         Top             =   1620
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   3
         Left            =   435
         Picture         =   "frmPaint.frx":22EA
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Fill"
         Top             =   495
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10300
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   7
         Left            =   50
         Picture         =   "frmPaint.frx":236C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Ellipse"
         Top             =   2370
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10301
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   6
         Left            =   50
         Picture         =   "frmPaint.frx":23D9
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Rectangle"
         Top             =   1995
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10302
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   8
         Left            =   50
         Picture         =   "frmPaint.frx":2446
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Text"
         Top             =   2745
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   9
         Left            =   435
         Picture         =   "frmPaint.frx":27C8
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Arrow"
         Top             =   1620
         WhatsThisHelpID =   10340
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   0
         Left            =   50
         Picture         =   "frmPaint.frx":2813
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Select Area"
         Top             =   120
         WhatsThisHelpID =   10359
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   1
         Left            =   50
         Picture         =   "frmPaint.frx":2B91
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Pick Color"
         Top             =   495
         WhatsThisHelpID =   10361
         Width           =   390
      End
      Begin VB.Frame fraOptDot 
         Height          =   1215
         Left            =   90
         TabIndex        =   31
         Top             =   3600
         WhatsThisHelpID =   10335
         Width           =   660
         Begin VB.Label lblDot 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   75
            TabIndex        =   32
            Top             =   150
            WhatsThisHelpID =   10336
            Width           =   255
         End
         Begin VB.Shape shpDot 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   30
            Index           =   0
            Left            =   195
            Shape           =   3  'Circle
            Top             =   270
            Width           =   30
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  'Solid
            Height          =   45
            Index           =   1
            Left            =   435
            Shape           =   3  'Circle
            Top             =   255
            Width           =   45
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  'Solid
            Height          =   60
            Index           =   2
            Left            =   165
            Shape           =   3  'Circle
            Top             =   495
            Width           =   60
         End
         Begin VB.Shape shpDot 
            BorderStyle     =   0  'Transparent
            FillStyle       =   0  'Solid
            Height          =   75
            Index           =   3
            Left            =   420
            Shape           =   3  'Circle
            Top             =   495
            Width           =   75
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   4
            Left            =   150
            Shape           =   3  'Circle
            Top             =   730
            Width           =   90
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   5
            Left            =   405
            Shape           =   3  'Circle
            Top             =   715
            Width           =   105
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  'Solid
            Height          =   120
            Index           =   6
            Left            =   140
            Shape           =   3  'Circle
            Top             =   970
            Width           =   120
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   7
            Left            =   390
            Shape           =   3  'Circle
            Top             =   960
            Width           =   135
         End
      End
      Begin VB.Frame fraBrush 
         Height          =   1545
         Left            =   90
         TabIndex        =   68
         Top             =   4815
         Visible         =   0   'False
         WhatsThisHelpID =   10335
         Width           =   660
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   9
            Left            =   405
            Picture         =   "frmPaint.frx":2F19
            Top             =   1290
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   8
            Left            =   120
            Picture         =   "frmPaint.frx":2F5B
            Top             =   1290
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   1
            Left            =   405
            Picture         =   "frmPaint.frx":2F9A
            Top             =   210
            Width           =   135
         End
         Begin VB.Label lblBrush 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   60
            TabIndex        =   69
            Top             =   150
            WhatsThisHelpID =   10336
            Width           =   255
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   0
            Left            =   120
            Picture         =   "frmPaint.frx":2FDE
            Top             =   210
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   3
            Left            =   405
            Picture         =   "frmPaint.frx":3022
            Top             =   480
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   2
            Left            =   120
            Picture         =   "frmPaint.frx":3066
            Top             =   480
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   6
            Left            =   120
            Picture         =   "frmPaint.frx":30AB
            Top             =   1020
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   7
            Left            =   405
            Picture         =   "frmPaint.frx":30ED
            Top             =   1020
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   5
            Left            =   405
            Picture         =   "frmPaint.frx":312F
            Top             =   750
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   4
            Left            =   120
            Picture         =   "frmPaint.frx":3173
            Top             =   750
            Width           =   135
         End
      End
      Begin VB.Frame fraOptFill 
         Height          =   1110
         Left            =   75
         TabIndex        =   33
         Top             =   4815
         Visible         =   0   'False
         WhatsThisHelpID =   10333
         Width           =   705
         Begin VB.Label lblFill 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   275
            Left            =   60
            TabIndex        =   34
            Top             =   150
            WhatsThisHelpID =   10334
            Width           =   570
         End
         Begin VB.Shape shpRect 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FFFFFF&
            Height          =   150
            Index           =   0
            Left            =   140
            Top             =   210
            Width           =   420
         End
         Begin VB.Shape shpRect 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   150
            Index           =   1
            Left            =   135
            Top             =   525
            Width           =   420
         End
         Begin VB.Shape shpRect 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   150
            Index           =   2
            Left            =   140
            Top             =   840
            Width           =   420
         End
      End
   End
   Begin VB.HScrollBar hscPaint 
      Height          =   255
      LargeChange     =   100
      Left            =   855
      Max             =   0
      SmallChange     =   10
      TabIndex        =   55
      Top             =   6150
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.VScrollBar vscPaint 
      Height          =   6165
      LargeChange     =   1000
      Left            =   7215
      Max             =   0
      SmallChange     =   100
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraScroll 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7200
      TabIndex        =   57
      Top             =   5670
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraColor 
      Height          =   860
      Left            =   0
      TabIndex        =   0
      Top             =   6315
      Width           =   7455
      Begin MSComDlg.CommonDialog cdlPrint 
         Left            =   4680
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin MSComDlg.CommonDialog cdlFonts 
         Left            =   5190
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdlOpen 
         Left            =   6210
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   $"frmPaint.frx":31B6
         Flags           =   4
      End
      Begin MSComDlg.CommonDialog cdlColor 
         Left            =   5715
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdlSave 
         Left            =   6720
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "*.brg"
         DialogTitle     =   "Save As"
         Filter          =   "Bitmap Files (*.bmp) |*.bmp"
      End
      Begin VB.Label lblColor 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   25
         Left            =   4080
         TabIndex        =   29
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   24
         Left            =   4080
         TabIndex        =   28
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   23
         Left            =   3825
         TabIndex        =   27
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   22
         Left            =   3825
         TabIndex        =   26
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   21
         Left            =   3555
         TabIndex        =   25
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   20
         Left            =   3555
         TabIndex        =   24
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   19
         Left            =   3285
         TabIndex        =   23
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   18
         Left            =   3285
         TabIndex        =   22
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   17
         Left            =   3015
         TabIndex        =   21
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00004040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   16
         Left            =   3015
         TabIndex        =   20
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   15
         Left            =   2745
         TabIndex        =   19
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   14
         Left            =   2745
         TabIndex        =   18
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   2475
         TabIndex        =   17
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   12
         Left            =   2475
         TabIndex        =   16
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   2200
         TabIndex        =   15
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   2200
         TabIndex        =   14
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   1935
         TabIndex        =   13
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   1935
         TabIndex        =   12
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   1660
         TabIndex        =   11
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   1660
         TabIndex        =   10
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   1400
         TabIndex        =   9
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1400
         TabIndex        =   8
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1125
         TabIndex        =   7
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1125
         TabIndex        =   6
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblForeColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lblFillColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   375
         TabIndex        =   5
         Top             =   420
         Width           =   255
      End
      Begin VB.Label label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   555
         Left            =   150
         TabIndex        =   3
         Top             =   210
         Width           =   555
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   850
         TabIndex        =   2
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   850
         TabIndex        =   1
         Top             =   225
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   54
      Top             =   7140
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9051
            MinWidth        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPaint 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6030
      Left            =   840
      MousePointer    =   99  'Custom
      ScaleHeight     =   398
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   411
      TabIndex        =   49
      Top             =   15
      Width           =   6225
      Begin VB.PictureBox picClipboard 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   630
         Left            =   1200
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtText 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   52
         Top             =   180
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.PictureBox picBuffer 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   2040
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picSelect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   480
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   50
         Top             =   120
         Width           =   615
      End
      Begin VB.Image imgBezier 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   60
         Index           =   0
         Left            =   2880
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgBezier 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   60
         Index           =   3
         Left            =   3240
         Top             =   600
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgBezier 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   60
         Index           =   2
         Left            =   3240
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgBezier 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   60
         Index           =   1
         Left            =   2880
         Top             =   600
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label lblTextSize 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.PictureBox picPaintResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Index           =   0
      Left            =   7080
      MousePointer    =   9  'Size W E
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   58
      Top             =   3000
      Width           =   70
   End
   Begin VB.PictureBox picPaintResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Index           =   2
      Left            =   7080
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   60
      Top             =   6045
      Width           =   70
   End
   Begin VB.PictureBox picPaintResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Index           =   1
      Left            =   3930
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   59
      Top             =   6045
      Width           =   70
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "C&rop"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuBorderStyle 
         Caption         =   "&Border Style"
         Begin VB.Menu mnuBS 
            Caption         =   "&Solid"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuBS 
            Caption         =   "&Dash"
            Index           =   1
         End
         Begin VB.Menu mnuBS 
            Caption         =   "D&ot"
            Index           =   2
         End
         Begin VB.Menu mnuBS 
            Caption         =   "D&ashDot"
            Index           =   3
         End
         Begin VB.Menu mnuBS 
            Caption         =   "Da&shDotDot"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFillStyle 
         Caption         =   "Fi&ll Style"
         Begin VB.Menu mnuFS 
            Caption         =   "&Solid"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Transparent"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Horizontal Line"
            Index           =   2
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Vertical Line"
            Index           =   3
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Downward Diagonal"
            Index           =   4
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Upward Diagonal"
            Index           =   5
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Cross"
            Index           =   6
         End
         Begin VB.Menu mnuFS 
            Caption         =   "Diagona&l Cross"
            Index           =   7
         End
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuForegroundColor 
         Caption         =   "F&oreground Color..."
      End
      Begin VB.Menu mnuFillColor 
         Caption         =   "Fi&ll Color..."
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuEffect 
      Caption         =   "Effec&t"
      Begin VB.Menu mnuResize 
         Caption         =   "Re&size"
         Begin VB.Menu mnuResize25 
            Caption         =   "25%"
         End
         Begin VB.Menu mnuResize50 
            Caption         =   "50%"
         End
         Begin VB.Menu mnuResize75 
            Caption         =   "75%"
         End
         Begin VB.Menu mnuResize125 
            Caption         =   "125%"
         End
         Begin VB.Menu mnuResize150 
            Caption         =   "150%"
         End
         Begin VB.Menu mnuResize175 
            Caption         =   "175%"
         End
         Begin VB.Menu mnuResize200 
            Caption         =   "200%"
         End
         Begin VB.Menu mnuSep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuResizeBoth 
            Caption         =   "&Both"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuResizeWidth 
            Caption         =   "&Width"
         End
         Begin VB.Menu mnuResizeHeight 
            Caption         =   "&Height"
         End
      End
      Begin VB.Menu mnuFlip 
         Caption         =   "&Flip"
         Begin VB.Menu mnuFlipHorizontal 
            Caption         =   "&Horizontal"
         End
         Begin VB.Menu mnuFlipVertical 
            Caption         =   "&Vertical"
         End
      End
      Begin VB.Menu mnuRotate 
         Caption         =   "&Rotate"
         Begin VB.Menu mnuRotate45 
            Caption         =   "By 45°"
         End
         Begin VB.Menu mnuRotate90 
            Caption         =   "By 90°"
         End
         Begin VB.Menu mnuRotate135 
            Caption         =   "By 135°"
         End
         Begin VB.Menu mnuRotate180 
            Caption         =   "By 180°"
         End
         Begin VB.Menu mnuRotate225 
            Caption         =   "By 225°"
         End
         Begin VB.Menu mnuRotate270 
            Caption         =   "By 270°"
         End
         Begin VB.Menu mnuRotate315 
            Caption         =   "By 315°"
         End
         Begin VB.Menu mnuSep7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRotateClockwise 
            Caption         =   "&Clockwise"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRotateAntiClockwise 
            Caption         =   "&Anti-Clockwise"
         End
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "&Filte&r"
      Begin VB.Menu mnuBlacknWhite 
         Caption         =   "&Black && White"
      End
      Begin VB.Menu mnuBlur 
         Caption         =   "B&lur"
      End
      Begin VB.Menu mnuBrightness 
         Caption         =   "B&rightness"
      End
      Begin VB.Menu mnuCrease 
         Caption         =   "&Crease"
      End
      Begin VB.Menu mnuDarkness 
         Caption         =   "&Darkness"
      End
      Begin VB.Menu mnuDiffuse 
         Caption         =   "Di&ffuse"
      End
      Begin VB.Menu mnuEmboss 
         Caption         =   "&Emboss"
      End
      Begin VB.Menu mnuGrayBlacknWhite 
         Caption         =   "Gra&y Black && White"
      End
      Begin VB.Menu mnuGrayscale 
         Caption         =   "&Grayscale"
      End
      Begin VB.Menu mnuInvertColors 
         Caption         =   "&Invert Colors"
      End
      Begin VB.Menu mnuReplaceColors 
         Caption         =   "&Replace Colors"
      End
      Begin VB.Menu mnuSharpen 
         Caption         =   "&Sharpen"
      End
      Begin VB.Menu mnuSnow 
         Caption         =   "S&now"
      End
      Begin VB.Menu mnuWave 
         Caption         =   "&Wave"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuTFilter 
      Caption         =   "&TFilter"
      Visible         =   0   'False
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Black && White"
         Index           =   0
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "B&lur"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Light"
         Index           =   2
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Crease"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Dirty"
         Index           =   4
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "Di&ffuse"
         Index           =   5
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Emboss"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "Gra&y Black && White"
         Index           =   7
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Grayscale"
         Index           =   8
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Invert Colors"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Replace Color"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Sharpen"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "S&now"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Wave"
         Index           =   13
      End
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name  : frmPaint.frm                                                 **
'** Language   : Visual Basic 6.0                                             **
'** References : Microsoft Scripting Runtime (only for mdlGeneral.ForceSave)  **
'** Components : * Microsoft Common Dialog Control 6.0 (SP3)                  **
'**              * Microsoft Windows Common Controls 6.0                      **
'** Modules    : mdlAPI, mdlEffect, mdlFilter, mdlGeneral                     **
'** Developer  : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A simple drawing program similar to Microsoft Paint plus     **
'**              several image filters                                        **
'** Features   :                                                              **
'** - Drawing tools: curve, polygon, filter brush, brush (10 different        **
'**                  shapes), air brush, text, fill, rectangle, square,       **
'**                  rounded rectangle, rounded square, ellipse, circle,      **
'**                  pencil, eraser and pick                                  **
'** - Drawing properties: foreground color, fill color, fill style,           **
'**                       draw width, border style and font                   **
'** - Selection tool: move, cut, copy, paste, delete, crop, apply effects,    **
'**                   apply filters                                           **
'** - Effects: resize, flip horizontal/vertical, rotate, clear                **
'** - Filters: black and white, blur, brightness, crease, darkness, diffuse,  **
'**            emboss, gray black and white, grayscale, invert colors,        **
'**            replace colors, sharpen, snow and wave                         **
'** - Undo/redo (limited only by memory, currently I set it to 10x undo/redo) **
'** - Others: scroll bars, zoom, resizable paint area, hand, status bar,      **
'**           open, save, and print                                           **
'** Version    : 1.02                                                         **
'** - 1.00 (August 10, 2003): First release                                   **
'** - 1.01 (August 13, 2003):                                                 **
'**     * bugs fixed on pressing cancel on open, save as and print dialog box **
'**     * bugs fixed on filter brushing on the top left of the paint area     **
'** - 1.02 (August 15, 2003):                                                 **
'**     bugs fixed on resizing and zooming the image multiple times           **
'*******************************************************************************

Option Explicit

'Enumeration declaration
Public Enum enmStatusBar
  conStPaintArea = 0
  conStColorBox = 1
  conStForeColorBox = 2
  conStBackColorBox = 3
  conStFiltering = 4
  conStRetrieveingColor = 5
End Enum
Dim sng As Single

Private Enum enmTool
  'the values below must match optTools index
  conTSelect = 0
  conTPick = 1
  conTEraser = 2
  conTFill = 3
  conTPencil = 4
  conTLine = 5
  conTRect = 6
  conTEllipse = 7
  conTText = 8
  conTArrow = 9
  conTAirBrush = 10
  conTRoundRect = 11
  conTPolygon = 12
  conTCurve = 13
  conTFilter = 14
  conTZoom = 15
  conTBrush = 16
  conTHand = 17
End Enum

Private Enum enmFillStyle
  conTsBorderOnly = 0
  conTsBorderFill = 1
  conTsFillOnly = 2
End Enum

Private Enum enmBrushShape
  'the values below must match imgBrush index
  conFilledRect = 0
  conFilledCircle = 1
  conRect = 2
  conCircle = 3
  conCross = 4
  conDiagonalCross = 5
  conUpwardDiagonal = 6
  conDownwardDiagonal = 7
  conHorizontal = 8
  conVertical = 9
End Enum

'Paint area resize direction constants declaration
Private Const conResizeWE = 0
Private Const conResizeNS = 1
Private Const conResizeNWSE = 2

'Default value
Private Const conDefaultActiveTool = conTPencil
Private Const conDefaultActiveFilterTool = conFltBrightness
Private Const conDefaultBorderStyle = vbBSSolid
Private Const conDefaultBrushShape = conFilledRect
Private Const conDefaultDotWidth = 0
Private Const conDefaultFillStyle = conTsBorderOnly
Private Const conDefaultInsideFillStyle = vbFSSolid
Private Const conDefaultPaintHeight = 6000
Private Const conDefaultPaintWidth = 6400

'Other constants declaration
Private Const conBufMax = 10               'maximum buffer for undo/redo feature
                                           '  (be careful increasing this value,
                                           '           it can make your computer
                                           '                  run out of memory)
Private Const conProgramTitle = "VB Paint"

'Variable Declaration
Private blnDrag As Boolean              'condition whether mouse move is to drag
Private blnDrawing As Boolean              'condition when mouse move is to draw
Private blnDrawingPolygon As Boolean                  'condition to draw polygon
Private blnFirstMoving As Boolean              'condition whether it's the first
                                               '   selected object moving action
Private blnMoving As Boolean                       'condition when mouse move is
                                                   ' to move the selected object
Private blnPicChanged As Boolean    'condition that the picture has been changed
                                    ' so the save confirmation on exit is needed
Private blnResize As Boolean      'condition that the paint area is being resize
Private lngDragStart As mdlAPI.typPoint  'coordinate where the drag action start
Private lngP1 As mdlAPI.typPoint         'the starting coordinate marked by user
Private lngP2 As mdlAPI.typPoint           'the ending coordinate marked by user
Private lngPolygon() As mdlAPI.typPoint     'to store polygon points information
Private intActiveFilterTool As enmFilter              'the active filter tool id
Private intActiveTool As enmTool     'the active tool id (active optTools index)
Private intBrushShape As enmBrushShape               'current active brush shape
Private intBufCur As Integer             'current buffer (for undo/redo feature)
Private intBufEnd As Integer           'last buffer used (for undo/redo feature)
Private intBufStart As Integer        'first buffer used (for undo/redo feature)
Private intDot As Integer                          'the width of the dot to draw
Private intDrawStyle As Integer                      'current .DrawSyle property
Private intFillStyle As enmFillStyle                'the current fill style used
Private intInsideFillStyle As Integer               'current .FillStyle property
Private sngZoomFactor As Single                             'current zoom factor
Private strFileName As String   'image file name (null string for unnamed image)

' Purpose    : Adjust lngP2 coordinate to agree with Shift or Ctrl key status as
'              specified below:
'              - Shift key pressed is to draw a square shape like square,
'                circle, 45-degree line, etc.
'              - Ctrl key pressed and blnEnableCtrl = true are to draw a
'                horizontal or vertical shape, like horizontal line, vertical
'                line, etc.
' Assumption : These global variables has been initiated:
'                lngP1, lngP2
' Effect     : As specified
' Inputs     : * X (current X coordinate)
'              * Y (current Y coordinate)
'              * Shift (shift key status)
'              * blnEnableCtrl (condition whether ctrl key status will effect
'                               the drawing or not)
' Returns    : -
Private Sub AdjustP2(X As Single, Y As Single, Shift As Integer, _
                     Optional blnEnableCtrl As Boolean = False)
  On Error GoTo ErrorHandler
  
  If Shift = vbShiftMask Then
    'Draw a square shape
    If Abs(X - lngP1.X) <= Abs(Y - lngP1.Y) Then
      lngP2.X = X
      If Y > lngP1.Y Then
        lngP2.Y = lngP1.Y + Abs(X - lngP1.X)
      Else
        lngP2.Y = lngP1.Y - Abs(X - lngP1.X)
      End If
    Else
      If X > lngP1.X Then
        lngP2.X = lngP1.X + Abs(Y - lngP1.Y)
      Else
        lngP2.X = lngP1.X - Abs(Y - lngP1.Y)
      End If
      lngP2.Y = Y
    End If
  ElseIf (Shift = vbCtrlMask) And blnEnableCtrl Then
    'Draw a horizontal or vertical shape
    If Abs(X - lngP1.X) <= Abs(Y - lngP1.Y) Then
      '- Horizontal shape
      lngP2.X = lngP1.X
      lngP2.Y = Y
    Else
      '- Vertical shape
      lngP2.X = X
      lngP2.Y = lngP1.Y
    End If
  Else
    'Draw a free shape
    lngP2.X = X
    lngP2.Y = Y
  End If
  Exit Sub
  
ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Adjust paint resize boxes (the little box on the right, bottom
'              abd bottom-right of the paint area) position to agree with paint
'              area width and height
' Assumption : These components exist in this form:
'                picPaint, picPaintResize
' Effect     : The paint resize boxes have been positioned to the middle right,
'              middle bottom and bottom-right next to the paint area
' Inputs     : -
' Returns    : -
Public Sub AdjustPaintResizeBox()
  On Error GoTo ErrorHandler
  
  picPaintResize(conResizeWE).Left = picPaint.Left + picPaint.Width
  picPaintResize(conResizeWE).Top = picPaint.Top + (picPaint.Height / 2)
  picPaintResize(conResizeNS).Left = picPaint.Left + (picPaint.Width / 2)
  picPaintResize(conResizeNS).Top = picPaint.Top + picPaint.Height
  picPaintResize(conResizeNWSE).Left = picPaintResize(conResizeWE).Left
  picPaintResize(conResizeNWSE).Top = picPaintResize(conResizeNS).Top
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Change the cursor in the paint area
' Assumptions: * This global variable has been initiated:
'                  intActiveTool
'              * This component exists in this form:
'                  picPaint
'              * The cursor file needed is exist in the sub directory "\Cursors"
' Effect     : The cursor in the paint area has been changed
' Inputs     : -
' Returns    : -
Private Sub ChangePaintCursor()
  On Error GoTo ErrorHandler                     'don't change the cursor if the
                                                 '     file needed doesn't exist
  With picPaint
    .MousePointer = vbCustom
    Select Case intActiveTool
      Case conTAirBrush
        .MouseIcon = LoadPicture(App.Path & "\Cursors\airbrush.cur")
      Case conTBrush
        .MouseIcon = LoadPicture(App.Path & "\Cursors\brush.cur")
      Case conTEraser
        .MouseIcon = LoadPicture(App.Path & "\Cursors\eraser.cur")
      Case conTFill
        .MouseIcon = LoadPicture(App.Path & "\Cursors\fill.cur")
      Case conTFilter
        .MouseIcon = LoadPicture(App.Path & "\Cursors\filter.cur")
      Case conTPencil
        .MouseIcon = LoadPicture(App.Path & "\Cursors\pencil.cur")
      Case conTPick
        .MouseIcon = LoadPicture(App.Path & "\Cursors\pick.cur")
      Case conTText
        .MouseIcon = LoadPicture(App.Path & "\Cursors\text.cur")
      Case conTSelect, conTCurve
        .MousePointer = vbDefault
      Case conTZoom
        .MouseIcon = LoadPicture(App.Path & "\Cursors\zoom.cur")
      Case conTHand
        .MouseIcon = LoadPicture(App.Path & "\Cursors\handflat.cur")
      Case Else
        .MouseIcon = LoadPicture(App.Path & "\Cursors\cross.cur")
    End Select
  End With

ErrorHandler:
End Sub

' Purpose    : Clear image buffer (for undo/redo feature)
' Assumption : These components exist in this form:
'                mnuRedo, mnuUndo, picBuffer(), picPaint
' Effects    : These global variables has been changed as following:
'              * intBufCur = 0
'              * intBufStart = 0
'              * intBufEnd = 0
'              * picBuffer.ubound = 0
'              * picBuffer(0).Picture = picPaint.Image
' Inputs     : -
' Returns    : -
Private Sub ClearImageBuffer()
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  intBufCur = 0
  intBufStart = 0
  intBufEnd = 0
  For i = 1 To picBuffer.UBound
    Unload picBuffer(i)
  Next
  picBuffer(intBufCur).Picture = picPaint.Image
  'save the paint area width and height for undo/redo action
  '  on resized paint area
  picBuffer(intBufCur).Tag = CStr((picPaint.Width * 100000) + picPaint.Height)
  mnuUndo.Enabled = False
  mnuRedo.Enabled = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw an air brush effect with current foreground color and
'              draw mode in the paint area
' Assumption : This component exist in this form:
'                picPaint
' Effects    : The air brush effect has been drawn in the paint area
' Inputs     : * X, Y (center coordinate of the air brush)
'              * R (half of the width or height of the air brush)
' Returns    : -
Private Sub DrawAirBrush(X As Integer, Y As Integer, R As Integer)
  Const conIntencity = 0.25
  
  Dim i As Integer
  Dim intDrawWidth As Integer                  'to keep current draw width value
  
  On Error GoTo ErrorHandler
  
  With picPaint
    intDrawWidth = .DrawWidth
    .DrawWidth = 1
    Randomize
    For i = 1 To ((R * R) * conIntencity)
      picPaint.PSet (X - (R / 2) + (Rnd() * R), Y - (R / 2) + (Rnd() * R))
    Next
    .DrawWidth = intDrawWidth
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw an arrow from (X1,Y1) to (X2,Y2) in the paint area with
'              current foreground color, draw mode and draw width in the paint
'              area
' Assumption : This component exists in this form:
'                picPaint
' Effect     : The arrow has been drawn in the paint area
' Inputs     : X1, Y1, X2, Y2
' Returns    : -
Private Sub DrawArrow(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
  Const conAlphaTip = 45                  'the angle of the lines in arrow's tip
  Const conLenTip = 10                       'length of the lines in arrow's tip
  Const conPi = 3.14159265358979
  
  'Variables to draw arrow's tip
  Dim intSign As Integer
  Dim X3 As Integer
  Dim Y3 As Integer
  Dim X4 As Integer
  Dim Y4 As Integer
  Dim sngBeta As Single
  
  On Error GoTo ErrorHandler
  
  'Draw arrow's line
  picPaint.Line (X1, Y1)-(X2, Y2)
  'Calculate variables for arrow's tip
  If X2 - X1 <> 0 Then
    sngBeta = Atn((Y2 - Y1) / (X2 - X1)) * 180 / conPi
  Else
    sngBeta = 90
  End If
  If X2 > X1 Then
    intSign = 1
  ElseIf X2 < X1 Then
    intSign = -1
  ElseIf Y2 > Y1 Then
    intSign = 1
  ElseIf Y2 < Y1 Then
    intSign = -1
  End If
  X3 = X2 - ((conLenTip * Cos((conAlphaTip + sngBeta) * conPi / 180)) * intSign)
  Y3 = Y2 - ((conLenTip * Sin((conAlphaTip + sngBeta) * conPi / 180)) * intSign)
  X4 = X2 - ((conLenTip * Cos((conAlphaTip - sngBeta) * conPi / 180)) * intSign)
  Y4 = Y2 + ((conLenTip * Sin((conAlphaTip - sngBeta) * conPi / 180)) * intSign)
  'Draw arrow's tip
  picPaint.Line (X2, Y2)-(X3, Y3)
  picPaint.Line (X2, Y2)-(X4, Y4)
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw a brush shape intBrushShape at (X,Y) with
'              current foreground color, draw style and fill style in the paint
'              area
' Assumption : This component exists in this form:
'                picPaint
' Effect     : The brush shape has been drawn in the paint area
' Inputs     : intBrushShape, X, Y
' Returns    : -
Private Sub DrawBrush(intBrushShape As enmBrushShape, X As Single, Y As Single)
  Const conBrushSize = 3
  
  Dim intDrawWidth As Integer       'to keep current picPaint.DrawWidth property
  
  On Error GoTo ErrorHandler
  
  With picPaint
    intDrawWidth = .DrawWidth
    .DrawWidth = 1
    Select Case intBrushShape
      Case conFilledRect
        picPaint.FillStyle = intInsideFillStyle
        picPaint.Line (X - (conBrushSize * intDrawWidth), _
                       Y - (conBrushSize * intDrawWidth))- _
                      (X + (conBrushSize * intDrawWidth), _
                       Y + (conBrushSize * intDrawWidth)), , BF
      Case conFilledCircle
        picPaint.FillStyle = intInsideFillStyle
        picPaint.Circle (X, Y), conBrushSize * intDrawWidth
      Case conRect
        picPaint.FillStyle = vbFSTransparent
        picPaint.Line (X - (conBrushSize * intDrawWidth), _
                       Y - (conBrushSize * intDrawWidth))- _
                      (X + (conBrushSize * intDrawWidth), _
                       Y + (conBrushSize * intDrawWidth)), , B
      Case conCircle
        picPaint.FillStyle = vbFSTransparent
        picPaint.Circle (X, Y), conBrushSize * intDrawWidth
      Case conCross
        picPaint.Line (X - (conBrushSize * intDrawWidth), Y)- _
                      (X + (conBrushSize * intDrawWidth), Y)
        picPaint.Line (X, Y - (conBrushSize * intDrawWidth))- _
                      (X, Y + (conBrushSize * intDrawWidth))
      Case conDiagonalCross
        picPaint.Line (X - (conBrushSize * intDrawWidth), _
                       Y + (conBrushSize * intDrawWidth))- _
                      (X + (conBrushSize * intDrawWidth), _
                       Y - (conBrushSize * intDrawWidth))
        picPaint.Line (X - (conBrushSize * intDrawWidth), _
                       Y - (conBrushSize * intDrawWidth))- _
                      (X + (conBrushSize * intDrawWidth), _
                       Y + (conBrushSize * intDrawWidth))
      Case conUpwardDiagonal
        picPaint.Line (X - (conBrushSize * intDrawWidth), _
                       Y + (conBrushSize * intDrawWidth))- _
                      (X + (conBrushSize * intDrawWidth), _
                       Y - (conBrushSize * intDrawWidth))
      Case conDownwardDiagonal
        picPaint.Line (X - (conBrushSize * intDrawWidth), _
                       Y - (conBrushSize * intDrawWidth))- _
                      (X + (conBrushSize * intDrawWidth), _
                       Y + (conBrushSize * intDrawWidth))
      Case conHorizontal
        picPaint.Line (X - (conBrushSize * intDrawWidth), Y)- _
                      (X + (conBrushSize * intDrawWidth), Y)
      Case conVertical
        picPaint.Line (X, Y - (conBrushSize * intDrawWidth))- _
                      (X, Y + (conBrushSize * intDrawWidth))
    End Select
    .DrawWidth = intDrawWidth
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw a bezier curve in the paint area with current foreground
'              color, draw mode (except for blnComplete = true which is
'              draw mode = copy) and draw width
' Assumption : These components exist in this form:
'                picPaint, imgBezier()
' Effect     : The bezier curve has been drawn in the paint area
' Inputs     : * blnCreate (condition to "create" [not to edit] the curve)
'              * blnComplete (condition to finish [draw with copy mode] the
'                             curve drawing)
'              * X, Y (center coordinate of the curve)
' Returns    : -
Private Sub DrawCurveBezier(Optional blnCreate As Boolean = False, _
                            Optional blnComplete As Boolean = False, _
                            Optional X As Single, Optional Y As Single)
  Const conCreateRadius = 50
  
  Dim i As Integer
  Dim intScaleMode                             'to keep current scale mode value
  Dim lngBezier(3) As typPoint
  
  On Error GoTo ErrorHandler
  
  intScaleMode = picPaint.ScaleMode
  picPaint.ScaleMode = vbPixels
  If blnCreate Then
    imgBezier(0).Top = Y - conCreateRadius
    imgBezier(0).Left = X - conCreateRadius
    imgBezier(1).Top = Y - conCreateRadius
    imgBezier(1).Left = X + conCreateRadius
    imgBezier(2).Top = Y + conCreateRadius
    imgBezier(2).Left = X - conCreateRadius
    imgBezier(3).Top = Y + conCreateRadius
    imgBezier(3).Left = X + conCreateRadius
    For i = 0 To 3
      imgBezier(i).Visible = True
    Next
  End If
  lngBezier(0).X = imgBezier(0).Left + (imgBezier(0).Width / 2)
  lngBezier(0).Y = imgBezier(0).Top + (imgBezier(0).Height / 2)
  lngBezier(1).X = imgBezier(1).Left + (imgBezier(0).Width / 2)
  lngBezier(1).Y = imgBezier(1).Top + (imgBezier(0).Height / 2)
  lngBezier(2).X = imgBezier(2).Left + (imgBezier(0).Width / 2)
  lngBezier(2).Y = imgBezier(2).Top + (imgBezier(0).Height / 2)
  lngBezier(3).X = imgBezier(3).Left + (imgBezier(0).Width / 2)
  lngBezier(3).Y = imgBezier(3).Top + (imgBezier(0).Height / 2)
  With picPaint
    If blnComplete Then
      .DrawMode = vbCopyPen
    End If
    mdlAPI.PolyBezier picPaint.hDC, lngBezier(0), 4
    .Refresh
  End With
  picPaint.ScaleMode = intScaleMode
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw polygon lngPolygon() in the paint area with current
'              foreground color, draw mode (except for blnComplete = true which
'              is draw mode = copy and foreground color depends on fill style)
'              and draw width (also with fill style intFillStyle and fill color
'              lblFillColor.BackColor for blnComplete = true)
' Assumption : * These components exist in this form:
'                  picPaint, imgBezier(), lnlForeColor, lblFillColor
'              * These global variables has been initiated
'                  lngPolygon()
' Effect     : The polygon has been drawn in the paint area
' Inputs     : * blnComplete (condition to finsih [draw with copy mode and fill
'                             style intFillStyle)
'              * blnOnlyDrawLastLine (condition to draw only the last line of
'                                     the polygon)
' Returns    : -
Private Sub DrawPolygon(Optional blnComplete As Boolean = True, _
                        Optional blnOnlyDrawLastLine = True)
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  With picPaint
    If blnComplete Then
      .DrawMode = vbCopyPen
      Select Case intFillStyle
        Case conTsBorderOnly
          .FillStyle = vbFSTransparent
          .ForeColor = lblForeColor.BackColor
        Case conTsBorderFill
          .FillStyle = intInsideFillStyle
          .ForeColor = lblForeColor.BackColor
          .FillColor = lblFillColor.BackColor
        Case conTsFillOnly
          .FillStyle = intInsideFillStyle
          .ForeColor = lblFillColor.BackColor
          .FillColor = lblFillColor.BackColor
      End Select
      mdlAPI.Polygon picPaint.hDC, lngPolygon(0), UBound(lngPolygon) + 1
      .Refresh
    Else
      If UBound(lngPolygon) > 0 Then
        If blnOnlyDrawLastLine Then
          picPaint.Line (lngPolygon(UBound(lngPolygon) - 1).X, _
                         lngPolygon(UBound(lngPolygon) - 1).Y)- _
                        (lngPolygon(UBound(lngPolygon)).X, _
                         lngPolygon(UBound(lngPolygon)).Y)
        Else
          For i = 1 To UBound(lngPolygon)
            picPaint.Line (lngPolygon(i - 1).X, lngPolygon(i - 1).Y)- _
                          (lngPolygon(i).X, lngPolygon(i).Y)
          Next
        End If
      End If
    End If
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw selection rectangle with xor mode and dot style
' Assumption : These components exist in this form:
'                picPaint, picSelect
' Effect     : As specified
' Inputs     : -
' Returns    : -
Public Sub DrawSelectionRect()
  'Variables to keep picPaint properties
  Dim intDrawStyle As Integer
  Dim intDrawMode As Integer
  Dim intDrawWidth As Integer
  
  On Error GoTo ErrorHandler
  
  If picSelect.Visible Then
    With picPaint
      intDrawMode = .DrawMode
      intDrawWidth = .DrawWidth
      picPaint.DrawStyle = vbDot
      picPaint.DrawMode = vbXorPen
      picPaint.DrawWidth = 1
      blnFirstMoving = False
      picPaint.Line (picSelect.Left - 1, picSelect.Top - 1)- _
                    (picSelect.Left + picSelect.Width, _
                     picSelect.Top + picSelect.Height), _
                    vbBlack Xor picPaint.BackColor, B
      .DrawStyle = intDrawStyle
      .DrawMode = intDrawMode
      .DrawWidth = intDrawWidth
    End With
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub Form_Activate()
  On Error GoTo ErrorHandler
  
  picPaint.SetFocus
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

  mnuNew_Click
  'Init default value
  intActiveFilterTool = conDefaultActiveFilterTool
  intActiveTool = conDefaultActiveTool
  intBrushShape = conDefaultBrushShape
  intDot = conDefaultDotWidth
  intInsideFillStyle = conDefaultFillStyle
  intFillStyle = conDefaultFillStyle
  mnuFilterTools(intActiveFilterTool).Checked = True
  picPaint.BorderStyle = conDefaultBorderStyle
  'Init dialogs' flags
  cdlSave.Flags = cdlOFNHideReadOnly Or _
                  cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  cdlOpen.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
  cdlFonts.Flags = cdlCFBoth Or cdlCFEffects Or cdlCFForceFontExist
  cdlPrint.Flags = cdlPDNoPageNums Or cdlPDNoSelection
  'Init fonts
  With picPaint
    .FontBold = txtText.FontBold
    .FontItalic = txtText.FontItalic
    .FontName = txtText.FontName
    .FontSize = txtText.FontSize
    .FontStrikethru = txtText.FontStrikethru
    .FontUnderline = txtText.FontUnderline
  End With
  'Init paint area size
  picPaint.Width = conDefaultPaintWidth
  picPaint.Height = conDefaultPaintHeight
  AdjustPaintResizeBox
  'Others
  UpdateStatusBar
  ChangePaintCursor
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Public Sub Form_Resize()
  On Error GoTo ErrorHandler
  
  If Me.WindowState <> vbMinimized Then
    'Limit the form's height
    If Me.Height < 2800 Then
      Me.Height = 2800
    End If
    'Adjust the tools and color box position and size
    fraTools.Height = Me.ScaleHeight - 900
    fraColor.Top = Me.ScaleHeight - 1110
    fraColor.Width = Me.Width - 90
    'Adjust the vertical scroll bar position, size and other properties
    With vscPaint
      If hscPaint.Visible Then
        .Max = (picPaint.Height - (Me.Height - hscPaint.Height - 1950)) / 10
      Else
        .Max = (picPaint.Height - (Me.Height - 1950)) / 10
      End If
      .Visible = (.Max > 0)
      If .Visible Then
        .Left = Me.Width - .Width - 110
        If hscPaint.Visible Then
          .Height = Me.ScaleHeight - fraColor.Height - hscPaint.Height - 150
        Else
          .Height = Me.ScaleHeight - fraColor.Height - 150
        End If
      End If
    End With
    'Adjust the horizontal scroll bar position, size and other properties
    With hscPaint
      If vscPaint.Visible Then
        .Max = (picPaint.Width - (Me.Width - vscPaint.Width - 1050)) / 10
      Else
        .Max = (picPaint.Width - (Me.Width - 1050)) / 10
      End If
      .Visible = (.Max > 0)
      If .Visible Then
        .Top = fraColor.Top - .Height + 110
        If vscPaint.Visible Then
          .Width = Me.Width - fraTools.Width - vscPaint.Width - 90
        Else
          .Width = Me.Width - fraTools.Width - 90
        End If
      End If
    End With
    'Re-adjust the vertical scroll bar max and height to match the new
    '  horizontal scroll bar properties
    If hscPaint.Visible Then
      vscPaint.Max = (picPaint.Height - _
                      (Me.Height - hscPaint.Height - 1850)) / 10
      vscPaint.Height = Me.ScaleHeight - fraColor.Height - hscPaint.Height - 150
    End If
    'Adjust the fraScroll properties
    If hscPaint.Visible And vscPaint.Visible Then
      fraScroll.Visible = True
      fraScroll.Left = vscPaint.Left
      fraScroll.Top = hscPaint.Top
    Else
      fraScroll.Visible = False
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ErrorHandler
  
  Dim intSave As Integer

  'Save confirmation
  If blnPicChanged = True Then
    intSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
      Case vbYes
        mnuSave_Click
        Cancel = blnPicChanged
      Case vbCancel
        Cancel = True
    End Select
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub fraOptDot_MouseDown(Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  For i = 0 To 7
    shpDot(i).FillColor = vbBlack
    shpDot(i).BorderColor = vbBlack
  Next
  'Set draw width intDot value and highlight the tool based on mouse click
  '  coordinate (X,Y)
  If Button = vbLeftButton Then
    If (Y >= 150) And (Y < 400) Then
      lblDot.Top = 150
      If (X >= 75) And (X < 325) Then
        intDot = 0
        lblDot.Left = 75
      ElseIf (X >= 325) And (X < 575) Then
        intDot = 1
        lblDot.Left = 325
      End If
    ElseIf (Y >= 400) And (Y < 650) Then
      lblDot.Top = 400
      If (X >= 75) And (X < 325) Then
        intDot = 2
        lblDot.Left = 75
      ElseIf (X >= 325) And (X < 575) Then
        intDot = 3
        lblDot.Left = 325
      End If
    ElseIf (Y >= 650) And (Y < 900) Then
      lblDot.Top = 650
      If (X >= 75) And (X < 325) Then
        shpDot(4).FillColor = vbWhite
        intDot = 4
        lblDot.Left = 75
      ElseIf (X >= 325) And (X < 575) Then
        intDot = 5
        lblDot.Left = 325
      End If
    ElseIf (Y >= 900) And (Y < 1150) Then
      lblDot.Top = 900
      If (X >= 75) And (X < 325) Then
        intDot = 6
        lblDot.Left = 75
      ElseIf (X >= 325) And (X < 575) Then
        intDot = 7
        lblDot.Left = 325
      End If
    End If
    shpDot(intDot).FillColor = vbWhite
    shpDot(intDot).BorderColor = vbWhite
    'Update the current drawing to match the new draw width
    UpdateDrawing
    picPaint.DrawWidth = intDot + 1
    UpdateDrawing
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub fraOptFill_MouseDown(Button As Integer, _
                                 Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'Set fill style intFillStyle and highlight the tool based on mouse click
  '  coordinate (X,Y)
  If Button = vbLeftButton Then
    If (Y >= 125) And (Y < 425) Then
      shpRect(0).BorderColor = vbWhite
      shpRect(1).BorderColor = vbBlack
      shpRect(2).BorderColor = vbBlack
      lblFill.Top = 150
      intFillStyle = conTsBorderOnly
    ElseIf (Y >= 450 And Y < 750) Then
      shpRect(0).BorderColor = vbBlack
      shpRect(1).BorderColor = vbWhite
      shpRect(2).BorderColor = vbBlack
      lblFill.Top = 465
      intFillStyle = conTsBorderFill
    ElseIf (Y >= 775 And Y < 1075) Then
      shpRect(0).BorderColor = vbBlack
      shpRect(1).BorderColor = vbBlack
      shpRect(2).BorderColor = vbWhite
      lblFill.Top = 780
      intFillStyle = conTsFillOnly
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub hscPaint_Change()
  Dim lngPicPaintLeft As Long
  
  On Error GoTo ErrorHandler
  
  lngPicPaintLeft = CLng(fraTools.Width) - (CLng(hscPaint.Value) * 10)
  picPaint.Left = lngPicPaintLeft
  AdjustPaintResizeBox
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Apply effect intEffect to the selection (if any) or to the paint
'              area
' Assumption : These components exist in this form:
'                mnuRotateClockWise, picPaint, picSelection, picImageEffect
' Effect     : As specified
' Inputs     : intImageEffect, sngResizeFactor
' Returns    : -
Private Sub ImageEffect(intEffect As enmEffect, _
                        Optional sngResizeFactor As Single, _
                        Optional sngRotateAngle As Single)
  Dim pic As PictureBox
  
  On Error GoTo ErrorHandler

  If picSelect.Visible Then
    Set pic = picSelect
  Else
    picPaint_DblClick
    Set pic = picPaint
  End If
  Select Case intEffect
    Case conEffResize
      If Not mnuResizeHeight.Checked Then
        mdlEffect.sngResizeWidth = sngResizeFactor
      End If
      If Not mnuResizeWidth.Checked Then
        mdlEffect.sngResizeHeight = sngResizeFactor
      End If
    Case conEffRotate
      mdlEffect.blnRotateClockWise = mnuRotateClockwise.Checked
      mdlEffect.sngRotateAngle = sngRotateAngle
  End Select
  If (intEffect <> conEffResize) Or _
     ((pic.ScaleWidth * Screen.TwipsPerPixelX * sngResizeFactor <= _
       mdlEffect.conMaxImageWidth) And _
      (pic.ScaleHeight * Screen.TwipsPerPixelY * sngResizeFactor <= _
       mdlEffect.conMaxImageHeight)) Then
    mdlEffect.ApplyEffect intEffect:=intEffect, _
                          pic:=pic, picTemp:=picImageEffect
  End If
  DrawSelectionRect
  If Not picSelect.Visible Then
    SetImageBuffer
  End If
  DrawSelectionRect
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Apply filter intFilter to the selection (if any) or to the paint
'              area
' Assumption : These components exist in this form:
'                picPaint, picSelection, lblForeColor, lblFillColor
' Effect     : As specified
' Input      : intFilter
' Returns    : -
Private Sub ImageFilter(intFilter As enmFilter, _
                        Optional X As Long = -1, Optional Y As Long = -1)
  On Error GoTo ErrorHandler
  
  Dim pic As PictureBox
  Dim X1 As Long
  Dim Y1 As Long
  Dim X2 As Long
  Dim Y2 As Long
  
  If picSelect.Visible Then
    Set pic = picSelect
  Else
    picPaint_DblClick
    Set pic = picPaint
  End If
  If intFilter = conFltReplaceColors Then
    mdlFilter.lngReplacedColor = lblForeColor.BackColor
    mdlFilter.lngReplaceWithColor = lblFillColor.BackColor
  End If
  If (intActiveTool = conTFilter) And ((X <> -1) Or (Y <> -1)) Then
    X1 = X - intDot
    Y1 = Y - intDot
    X2 = X + intDot
    Y2 = Y + intDot
    If (X2 >= 0) And (Y2 >= 0) Then
      mdlFilter.ApplyFilter intFilter:=intFilter, pic:=picPaint, _
                            X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2
    End If
  Else
    mdlFilter.ApplyFilter intFilter:=intFilter, pic:=pic
    DrawSelectionRect
    If Not picSelect.Visible Then
      SetImageBuffer
    End If
    DrawSelectionRect
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Zoom the paint area sngZoomFactor times (or no zoom if
'              blnNoZoom = true) and adjust the scroll bar so the coordinate
'              clicked by users is positioned in the center of visible
'              paint area
' Assumption : These components exist in this form:
'                picPaint, picZoom, picImageEffect, picPaintResize, hscPaint,
'                vcsPaint
' Effect     : * As specified
'              * if blnNoZoom = true then picPaintResize is shown, else
'                picPaintResize is hidden
' Inputs     : * X, Y (coordinate (in pixel) clicked by users that has been
'                      adjusted with zoom factor)
'              * blnNoZoom
' Returns    : -
Private Sub ImageZoom(Optional X As Long = 0, Optional Y As Long = 0, _
                      Optional blnNoZoom As Boolean = False)
  Dim lngHscValue As Long                  'adjusted horizontal scroll bar value
  Dim lngVscValue As Long                    'adjusted vertical scroll bar value
  Dim lngVisibleWidth As Long                   'the width of visible paint area
  Dim lngVisibleHeight As Long                 'the height of visible paint area
  
  On Error GoTo ErrorHandler
  
  If blnNoZoom Then
    If sngZoomFactor <> 1 Then
      sngZoomFactor = 1
      picPaint.Picture = picZoom.Image
      frmPaint.AdjustPaintResizeBox
      frmPaint.Form_Resize
      picPaintResize(0).Visible = True
      picPaintResize(1).Visible = True
      picPaintResize(2).Visible = True
    End If
  Else
    'Zoom the picture
    mdlEffect.sngResizeWidth = sngZoomFactor
    mdlEffect.sngResizeHeight = sngZoomFactor
    picPaintResize(0).Visible = False
    picPaintResize(1).Visible = False
    picPaintResize(2).Visible = False
    picPaint.Visible = False
    picPaint.Picture = picZoom.Image
    mdlEffect.ApplyEffect intEffect:=conEffResize, _
                          pic:=picPaint, picTemp:=picImageEffect
    'Arrange horizontal scroll bar value
    If hscPaint.Visible Then
      If vscPaint.Visible Then
        lngVisibleWidth = Me.Width - fraTools.Width - vscPaint.Width
      Else
        lngVisibleWidth = Me.Width - fraTools.Width
      End If
      lngHscValue = ((X - (lngVisibleWidth / 2)) / _
                     (picPaint.Width - lngVisibleWidth)) * hscPaint.Max
      If lngHscValue < 0 Then
        hscPaint.Value = 0
      ElseIf lngHscValue > hscPaint.Max Then
        hscPaint.Value = hscPaint.Max
      Else
        hscPaint.Value = lngHscValue
      End If
    End If
    'Arrange vertical scroll bar value
    If vscPaint.Visible Then
      If hscPaint.Visible Then
        lngVisibleHeight = Me.ScaleHeight - _
                           hscPaint.Height - fraColor.Height - sta.Height
      Else
        lngVisibleHeight = Me.ScaleHeight - fraColor.Height - sta.Height
      End If
      lngVscValue = ((Y - (lngVisibleHeight / 2)) / _
                     (picPaint.Height - lngVisibleHeight)) * vscPaint.Max
      If lngVscValue < 0 Then
        vscPaint.Value = 0
      ElseIf lngVscValue > vscPaint.Max Then
        vscPaint.Value = vscPaint.Max
      Else
        vscPaint.Value = lngVscValue
      End If
    End If
    picPaint.SetFocus
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub imgBezier_MouseDown(Index As Integer, Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'Start the drag operation on imgBezier(Index)
  lngDragStart.X = CLng(X)
  lngDragStart.Y = CLng(Y)
  blnDrag = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub imgBezier_MouseMove(Index As Integer, Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'Move imgBezier(Index) for the drag operation and update the bezier curve
  If blnDrag Then
    DrawCurveBezier
    picPaint.ScaleMode = vbTwips
    With imgBezier(Index)
      .Top = .Top + (Y - lngDragStart.Y)
      .Left = .Left + (X - lngDragStart.X)
    End With
    DrawCurveBezier
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub imgBezier_MouseUp(Index As Integer, Button As Integer, _
                              Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'End the drag operation on imgBezier(Index)
  blnDrag = False
  picPaint.ScaleMode = vbPixels
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub imgBrush_MouseDown(Index As Integer, Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
 On Error GoTo ErrorHandler
  
  intBrushShape = Index
  lblBrush.Top = imgBrush(Index).Top - (4 * Screen.TwipsPerPixelX)
  lblBrush.Left = imgBrush(Index).Left - (4 * Screen.TwipsPerPixelY)
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub lblColor_MouseDown(Index As Integer, Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  Select Case Button
    Case vbLeftButton
      'Set the foreground color and update the current drawing to match the new
      '  foreground color
      UpdateDrawing
      lblForeColor.BackColor = lblColor(Index).BackColor
      picPaint.DrawMode = vbXorPen
      picPaint.ForeColor = picPaint.BackColor Xor lblForeColor.BackColor
      UpdateDrawing
    Case vbRightButton
      'Set the background color
      lblFillColor.BackColor = lblColor(Index).BackColor
  End Select
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub lblColor_MouseMove(Index As Integer, Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  UpdateStatusBar intInfo:=conStColorBox
End Sub

Private Sub lblFillColor_DblClick()
  On Error GoTo ErrorHandler
  
  cdlColor.ShowColor
  lblFillColor.BackColor = cdlColor.Color
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub lblFillColor_MouseMove(Button As Integer, _
                                   Shift As Integer, X As Single, Y As Single)
  UpdateStatusBar intInfo:=conStBackColorBox
End Sub

Private Sub lblForeColor_DblClick()
  On Error GoTo ErrorHandler
  
  cdlColor.ShowColor
  'Update the current drawing to match with the new foreground color
  UpdateDrawing
  lblForeColor.BackColor = cdlColor.Color
  picPaint.DrawMode = vbXorPen
  picPaint.ForeColor = picPaint.BackColor Xor lblForeColor.BackColor
  UpdateDrawing
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub lblForeColor_MouseMove(Button As Integer, _
                                   Shift As Integer, X As Single, Y As Single)
  UpdateStatusBar intInfo:=conStForeColorBox
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show modal:=vbModal, ownerform:=Me
End Sub

Private Sub mnuBlacknWhite_Click()
  ImageFilter intFilter:=conFltBlacknWhite
End Sub

Private Sub mnuBlur_Click()
  ImageFilter intFilter:=conFltBlur
End Sub

Private Sub mnuBrightness_Click()
  ImageFilter intFilter:=conFltBrightness
End Sub

Private Sub mnuBS_Click(Index As Integer)
  On Error GoTo ErrorHandler
  
  Dim i As Integer
  
  For i = 0 To mnuBS.Count - 1
    mnuBS(i).Checked = False
  Next
  intDrawStyle = Index
  mnuBS(Index).Checked = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuClear_Click()
  On Error GoTo ErrorHandler
  
  picPaint_DblClick
  picPaint.Picture = Nothing
  SetImageBuffer
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuCopy_Click()
  On Error GoTo ErrorHandler
  
  picClipboard.Picture = picSelect.Image
  mnuPaste.Enabled = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuCrop_Click()
  picSelect.Visible = False
  picPaint.Picture = picSelect.Image
  SetImageBuffer
  Form_Resize
  AdjustPaintResizeBox
End Sub

Private Sub mnuCut_Click()
  mnuDelete_Click
  mnuCopy_Click
End Sub

Private Sub mnuDarkness_Click()
  ImageFilter intFilter:=conFltDarkness
End Sub

Private Sub mnuDelete_Click()
  On Error GoTo ErrorHandler
  
  picSelect.Visible = False
  With picPaint
    'Remove the selection rectangle
    .DrawMode = vbXorPen
    .DrawStyle = vbDot
    .DrawWidth = 1
    picPaint.Line (picSelect.Left - 1, picSelect.Top - 1)- _
                  (picSelect.Left + picSelect.ScaleWidth, _
                   picSelect.Top + picSelect.ScaleHeight), _
                  vbBlack Xor picPaint.BackColor, B
    'Delete the selection area
    .DrawMode = vbCopyPen
    .DrawStyle = intDrawStyle
    If blnFirstMoving Then
      picPaint.Line (lngP1.X + varIIf(lngP1.X < lngP2.X, 1, -1), _
                     lngP1.Y + varIIf(lngP1.Y < lngP2.Y, 1, -1))- _
                    (lngP2.X + varIIf(lngP2.X < lngP1.X, 1, -1), _
                     lngP2.Y + varIIf(lngP2.Y < lngP1.Y, 1, -1)), _
                    picPaint.BackColor, BF
    End If
    .SetFocus
  End With
  picSelect.Visible = False
  mnuCut.Enabled = False
  mnuCopy.Enabled = False
  mnuDelete.Enabled = False
  mnuCrop.Enabled = False
  SetImageBuffer
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuDiffuse_Click()
  ImageFilter intFilter:=conFltDiffuse
End Sub

Private Sub mnuEdit_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuEffect_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuEmboss_Click()
  ImageFilter intFilter:=conFltEmboss
End Sub

Private Sub mnuExit_Click()
  On Error GoTo ErrorHandler

  Unload Me
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuFile_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuFillColor_Click()
  lblFillColor_DblClick
End Sub

Private Sub mnuFilter_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuFilterTools_Click(Index As Integer)
  On Error GoTo ErrorHandler
  
  Dim i As Integer
  
  For i = 0 To mnuFilterTools.Count - 1
    mnuFilterTools(i).Checked = False
  Next
  mnuFilterTools(Index).Checked = True
  intActiveFilterTool = Index
  picPaint.SetFocus
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuFlipHorizontal_Click()
  ImageEffect intEffect:=conEffFlipHorizontal
End Sub

Private Sub mnuFlipVertical_Click()
  ImageEffect intEffect:=conEffFlipVertical
End Sub

Private Sub mnuFont_Click()
  On Error GoTo ErrorHandler
  
  With cdlFonts
    'Set font dialog box properties with current paint area font properties
    .FontBold = picPaint.FontBold
    .FontItalic = picPaint.FontItalic
    .FontName = picPaint.FontName
    .FontSize = picPaint.FontSize
    .FontStrikethru = picPaint.FontStrikethru
    .FontUnderline = picPaint.FontUnderline
    .Color = picPaint.ForeColor
    'Open font dialog box
    .ShowFont
    'Set paint area and text box txtText font properties with properties in font
    '  dialog box
    picPaint.FontBold = .FontBold
    picPaint.FontItalic = .FontItalic
    picPaint.FontName = .FontName
    picPaint.FontSize = .FontSize
    picPaint.FontStrikethru = .FontStrikethru
    picPaint.FontUnderline = .FontUnderline
    picPaint.ForeColor = .Color
    txtText.FontBold = .FontBold
    txtText.FontItalic = .FontItalic
    txtText.FontName = .FontName
    txtText.FontSize = .FontSize
    txtText.FontStrikethru = .FontStrikethru
    txtText.FontUnderline = .FontUnderline
    txtText.ForeColor = .Color
    lblTextSize.FontBold = .FontBold
    lblTextSize.FontItalic = .FontItalic
    lblTextSize.FontName = .FontName
    lblTextSize.FontSize = .FontSize
    lblTextSize.FontStrikethru = .FontStrikethru
    lblTextSize.FontUnderline = .FontUnderline
    lblForeColor.BackColor = .Color
    txtText_KeyDown 0, 0
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuForegroundCOlor_Click()
  lblForeColor_DblClick
End Sub

Private Sub mnuFS_Click(Index As Integer)
  Dim i As Integer
  
  On Error GoTo ErrorHandler

  For i = 0 To mnuFS.Count - 1
    mnuFS(i).Checked = False
  Next
  intInsideFillStyle = Index
  mnuFS(Index).Checked = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuGrayBlacknWhite_Click()
  ImageFilter intFilter:=conFltGrayBlacknWhite
End Sub

Private Sub mnuGrayscale_Click()
  ImageFilter intFilter:=conFltGrayscale
End Sub

Private Sub mnuHelp_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuInvertColors_Click()
  ImageEffect intEffect:=conEffInvertColors
End Sub

Private Sub mnuNew_Click()
  Dim i As Integer
  Dim intSave As Integer
  
  On Error GoTo ErrorHandler

  If blnPicChanged = True Then
    intSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
  Else
    intSave = vbNo
  End If
  If intSave = vbYes Then
    mnuSave_Click
  End If
  If intSave <> vbCancel Then
    picZoom.Width = picPaint.Width
    picZoom.Height = picPaint.Height
    picZoom.Picture = Nothing
    ImageZoom blnNoZoom:=True
    picPaint.Picture = Nothing
    blnPicChanged = False
    strFileName = ""
    UpdateFormTitle
    blnDrawingPolygon = False
    ReDim lngPolygon(0)
    For i = 0 To 3
      imgBezier(i).Visible = False
    Next
    sngZoomFactor = 1
    AdjustPaintResizeBox
    ClearImageBuffer
    picSelect.Visible = False
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuDelete.Enabled = False
    mnuCrop.Enabled = False
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuOpen_Click()
  Dim intSave As Integer
  
  On Error GoTo ErrorHandler
  
  If blnPicChanged Then
    intSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
  Else
    intSave = vbNo
  End If
  If intSave = vbYes Then
    mnuSave_Click
  End If
  If intSave <> vbCancel Then
    cdlOpen.ShowOpen
    If cdlOpen.FileName <> "" Then
      blnPicChanged = False
      mnuNew_Click
      picPaint.Picture = LoadPicture(cdlOpen.FileName)
      strFileName = cdlOpen.FileName
      UpdateFormTitle
      ClearImageBuffer
      optTools_Click Index:=conTZoom
    End If
  End If
  Form_Resize
  AdjustPaintResizeBox
  Exit Sub

ErrorHandler:
  If Err.Number <> conErrCancel Then
    ShowErrMessage intErr:=conErrReadImage
  End If
End Sub

Private Sub mnuPaste_Click()
  On Error GoTo ErrorHandler
  
  picPaint_DblClick
  If Not blnFirstMoving Then
    PlaceSelection
  End If
  picSelect.Picture = picClipboard.Image
  picPaint.DrawStyle = vbDot
  blnFirstMoving = False
  If picSelect.Visible Then
    picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y), _
                  vbBlack Xor picPaint.BackColor, B
  End If
  picPaint.DrawMode = vbXorPen
  picPaint.DrawWidth = 1
  picSelect.Left = 0
  picSelect.Top = 0
  picPaint.Line (-1, -1)-(picClipboard.Width, picClipboard.Height), _
                vbBlack Xor picPaint.BackColor, B
  picSelect.Visible = True
  If intActiveTool <> conTSelect Then
    intActiveTool = conTSelect
    optTools(conTSelect).SetFocus
  End If
  mnuCut.Enabled = True
  mnuCopy.Enabled = True
  mnuDelete.Enabled = True
  mnuCrop.Enabled = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuPrint_Click()
  Dim strImgTmpFile As String

  On Error GoTo ErrorHandler
  
  cdlPrint.ShowPrinter
  Printer.Copies = cdlPrint.Copies
  strImgTmpFile = "temp.bmp"
  If blnFileExist(strImgTmpFile) Then
    Kill strImgTmpFile
  End If
  ImageZoom blnNoZoom:=True
  SavePicture picPaint.Image, strImgTmpFile
  picPaint.Picture = LoadPicture(strImgTmpFile)
  Kill strImgTmpFile
  Printer.PaintPicture picPaint, 0, 0
  Printer.EndDoc
  Exit Sub

ErrorHandler:
  If Err.Number <> conErrCancel Then
    ShowErrMessage intErr:=conErrPrint
  End If
End Sub

Private Sub mnuRedo_Click()
  On Error GoTo ErrorHandler
  
  ImageZoom blnNoZoom:=True
  'Remove selection
  If picSelect.Visible Then
    picSelect.Visible = False
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuDelete.Enabled = False
    mnuCrop.Enabled = False
  End If
  'Set the current buffer index
  If intBufCur < conBufMax Then
    intBufCur = intBufCur + 1
  Else
    intBufCur = 0
  End If
  'Replace the paint area with image in picBuffer(intBufCur)
  picPaint.Picture = picBuffer(intBufCur).Image
  picPaint.Width = CLng(Left(picBuffer(intBufCur).Tag, _
                             Len(picBuffer(intBufCur).Tag) - 5))
  picPaint.Height = CLng(Right(picBuffer(intBufCur).Tag, 5))
  'Other settings
  mnuUndo.Enabled = True
  If intBufCur = intBufEnd Then
    mnuRedo.Enabled = False
  End If
  picPaint_DblClick
  AdjustPaintResizeBox
  Form_Resize
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuCrease_Click()
  ImageFilter intFilter:=conFltCrease
End Sub

Private Sub mnuReplaceColors_Click()
  ImageFilter intFilter:=conFltReplaceColors
End Sub

Private Sub mnuSnow_Click()
  ImageFilter intFilter:=conFltSnow
End Sub

Private Sub mnuResize125_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=1.25
End Sub

Private Sub mnuResize150_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=1.5
End Sub

Private Sub mnuResize175_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=1.75
End Sub

Private Sub mnuResize200_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=2
End Sub

Private Sub mnuResize25_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=0.25
End Sub

Private Sub mnuResize50_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=0.5
End Sub

Private Sub mnuResize75_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=0.75
End Sub

Private Sub mnuResizeBoth_Click()
  On Error GoTo ErrorHandler

  mnuResizeBoth.Checked = True
  mnuResizeWidth.Checked = False
  mnuResizeHeight.Checked = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuResizeHeight_Click()
  On Error GoTo ErrorHandler

  mnuResizeBoth.Checked = False
  mnuResizeWidth.Checked = False
  mnuResizeHeight.Checked = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuResizeWidth_Click()
  On Error GoTo ErrorHandler

  mnuResizeBoth.Checked = False
  mnuResizeWidth.Checked = True
  mnuResizeHeight.Checked = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuRotate135_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=135
End Sub

Private Sub mnuRotate180_Click()
  ImageEffect intEffect:=conEffFlipHorizontal
  ImageEffect intEffect:=conEffFlipVertical
End Sub

Private Sub mnuRotate225_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=225
End Sub

Private Sub mnuRotate270_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=270
End Sub

Private Sub mnuRotate315_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=315
End Sub

Private Sub mnuRotate45_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=45
End Sub

Private Sub mnuRotate90_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=90
End Sub

Private Sub mnuRotateAntiClockwise_Click()
  On Error GoTo ErrorHandler

  mnuRotateClockwise.Checked = False
  mnuRotateAntiClockwise.Checked = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuRotateClockwise_Click()
  On Error GoTo ErrorHandler

  mnuRotateClockwise.Checked = True
  mnuRotateAntiClockwise.Checked = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuSave_Click()
  On Error GoTo ErrorHandler
  
  If strFileName = "" Then
    mnuSaveAs_Click
  Else
    ImageZoom blnNoZoom:=True
    SavePicture picPaint.Image, strFileName
    blnPicChanged = False
    UpdateFormTitle
  End If
  Exit Sub
  
ErrorHandler:
  ShowErrMessage intErr:=conErrWrite
End Sub

Private Sub mnuSaveAs_Click()
  On Error GoTo ErrorHandler
  
  cdlSave.ShowSave
  If cdlSave.FileName <> "" Then
    strFileName = cdlSave.FileName
    mnuSave_Click
  End If
  Exit Sub
  
ErrorHandler:
  If Err.Number = conErrPermission Then
    If ForceSave(strFileName) Then
      Resume
    End If
  ElseIf Err.Number <> conErrCancel Then
    ShowErrMessage intErr:=conErrWrite
  End If
End Sub

Private Sub mnuSharpen_Click()
  ImageFilter intFilter:=conFltSharpen
End Sub



Private Sub mnuUndo_Click()
  On Error GoTo ErrorHandler

  ImageZoom blnNoZoom:=True
  'Place the selection
  If picSelect.Visible Then
    PlaceSelection
    picPaint.SetFocus
  Else
    picPaint_DblClick
  End If
  'Set the current buffer index
  If intBufCur > 0 Then
    intBufCur = intBufCur - 1
  Else
    intBufCur = conBufMax
  End If
  'Replace the paint area with image in picBuffer(intBufCur)
  picPaint.Picture = picBuffer(intBufCur).Image
  picPaint.Width = CLng(Left(picBuffer(intBufCur).Tag, _
                             Len(picBuffer(intBufCur).Tag) - 5))
  picPaint.Height = CLng(Right(picBuffer(intBufCur).Tag, 5))
  'Other settings
  If intBufCur = intBufStart Then
    mnuUndo.Enabled = False
  End If
  mnuRedo.Enabled = True
  AdjustPaintResizeBox
  Form_Resize
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuWave_Click()
  ImageFilter intFilter:=conFltWave
End Sub

Private Sub optTools_Click(Index As Integer)
  On Error GoTo ErrorHandler
  
  'Arrange draw width box and fill style box visibility
  Select Case intActiveTool
    Case conTAirBrush, conTArrow, conTCurve, conTEraser, _
         conTFilter, conTLine, conTPencil
      fraBrush.Visible = False
      fraOptDot.Visible = True
      fraOptFill.Visible = False
    Case conTRect, conTEllipse, conTRoundRect, conTPolygon
      fraBrush.Visible = False
      fraOptDot.Visible = True
      fraOptFill.Visible = True
    Case conTBrush
      fraBrush.Visible = True
      fraOptDot.Visible = True
      fraOptFill.Visible = False
    Case Else
      fraBrush.Visible = False
      fraOptDot.Visible = False
      fraOptFill.Visible = False
  End Select
  'Other settings
  If intActiveTool = conTFilter Then
    PopupMenu Menu:=mnuTFilter
  End If
  If intActiveTool = conTZoom Then
    picZoom.Width = picPaint.Width
    picZoom.Height = picPaint.Height
    picZoom.Picture = picPaint.Image
  End If
  If intActiveTool <> conTSelect Then
    PlaceSelection
  End If
  If (intActiveTool <> conTPick) And (intActiveTool <> conTHand) Then
    ImageZoom blnNoZoom:=True
  End If
  UpdateStatusBar
  ChangePaintCursor
  picPaint.SetFocus
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub optTools_MouseDown(Index As Integer, Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  If Button = vbLeftButton Then
    picPaint_DblClick
    intActiveTool = Index
    If intActiveTool = conTFilter Then
      PopupMenu Menu:=mnuTFilter
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_DblClick()
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  Select Case intActiveTool
    Case conTCurve
      If imgBezier(0).Visible Then
        DrawCurveBezier
        picPaint.DrawMode = vbCopyPen
        picPaint.ForeColor = lblForeColor.BackColor
        DrawCurveBezier blnComplete:=True
        For i = 0 To 3
          imgBezier(i).Visible = False
        Next
        SetImageBuffer
      End If
    Case conTPolygon
      If blnDrawingPolygon Then
        DrawPolygon blnComplete:=False
        DrawPolygon
        blnDrawingPolygon = False
        SetImageBuffer
      End If
    Case conTSelect
      PlaceSelection
  End Select
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim blnSuccess As Boolean

  On Error GoTo ErrorHandler

  If KeyCode = vbKeyReturn Then
    picPaint_DblClick
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_MouseDown(Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  Dim i As Long
  
  If Button = vbLeftButton Then
    blnDrawing = True
    lngP1.X = X
    lngP1.Y = Y
    With picPaint
      If intActiveTool = conTSelect Then
        .DrawStyle = vbDot
        .DrawWidth = 1
      Else
        .DrawStyle = intDrawStyle
        .DrawWidth = intDot + 1
      End If
      Select Case intActiveTool
        Case conTAirBrush
          .DrawMode = vbCopyPen
          .ForeColor = lblForeColor.BackColor
          DrawAirBrush CInt(X), CInt(Y), .DrawWidth * 4
        Case conTBrush
          .DrawMode = vbCopyPen
          .ForeColor = lblForeColor.BackColor
          .FillColor = lblForeColor.BackColor
          DrawBrush intBrushShape:=intBrushShape, X:=X, Y:=Y
        Case conTCurve
          If Not imgBezier(0).Visible Then
            .DrawMode = vbXorPen
            .ForeColor = picPaint.BackColor Xor lblForeColor.BackColor
            DrawCurveBezier blnCreate:=True, X:=X, Y:=Y
          End If
          lngP1.X = X
          lngP1.Y = Y
        Case conTEraser
          .DrawMode = vbCopyPen
          .ForeColor = .BackColor
          picPaint.Line (X, Y)-(X + .DrawWidth, Y - .DrawWidth), , B
        Case conTFill
          .DrawMode = vbCopyPen
          .FillColor = lblForeColor.BackColor
          .FillStyle = intInsideFillStyle
          mdlAPI.ExtFloodFill .hDC, X, Y, .Point(X, Y), 1
        Case conTFilter
          ImageFilter intFilter:=intActiveFilterTool, X:=CLng(X), Y:=CLng(Y)
        Case conTHand
          .ScaleMode = vbTwips
          .MouseIcon = LoadPicture(App.Path & "\Cursors\handgrab.cur")
          lngP1.X = (X * Screen.TwipsPerPixelX) + .Left
          lngP1.Y = (Y * Screen.TwipsPerPixelY) + .Top
          lngDragStart.X = .Left
          lngDragStart.Y = .Top
          blnDrag = True
        Case conTPencil
          .DrawMode = vbCopyPen
          .ForeColor = lblForeColor.BackColor
          picPaint.Line (X, Y)-(X, Y), , B
        Case conTPick
          lblForeColor.BackColor = picPaint.Point(X, Y)
        Case conTPolygon
          If Not blnDrawingPolygon Then
            blnDrawingPolygon = True
            ReDim lngPolygon(0)
            lngPolygon(0).X = X
            lngPolygon(0).Y = Y
          Else
            ReDim Preserve lngPolygon(UBound(lngPolygon) + 1)
            lngPolygon(UBound(lngPolygon)).X = X
            lngPolygon(UBound(lngPolygon)).Y = Y
            DrawPolygon blnComplete:=False
          End If
          .DrawMode = vbXorPen
          .FillStyle = vbFSTransparent
          .ForeColor = .BackColor Xor lblForeColor.BackColor
        Case conTText
          With txtText
            If Not .Visible Then
              .BackColor = picPaint.BackColor
              .ForeColor = lblForeColor.BackColor
              .Left = X
              .Top = Y
              .Text = ""
              .Visible = True
              .SetFocus
            Else
              .Tag = "moving"
              .Move X, Y
              .SetFocus
            End If
          End With
        Case Else
          If (intActiveTool = conTArrow) Or _
             (intActiveTool = conTSelect) Or (intActiveTool = conTLine) Then
            picPaint.Line (X, Y)-(X, Y)
          End If
          If intActiveTool = conTSelect Then
            .DrawWidth = 1
            PlaceSelection
          End If
          .DrawMode = vbXorPen
          If (intActiveTool = conTLine) Or _
             (intActiveTool = conTArrow) Or (intActiveTool = conTSelect) Then
            .ForeColor = .BackColor Xor lblForeColor.BackColor
            .FillStyle = vbFSTransparent
          Else
            Select Case intFillStyle
              Case conTsBorderOnly
                .FillStyle = vbFSTransparent
                .ForeColor = .BackColor Xor lblForeColor.BackColor
              Case conTsBorderFill
                .FillStyle = intInsideFillStyle
                .ForeColor = .BackColor Xor lblForeColor.BackColor
                .FillColor = .BackColor Xor lblFillColor.BackColor
              Case conTsFillOnly
                .FillStyle = intInsideFillStyle
                .ForeColor = .BackColor Xor lblFillColor.BackColor
                .FillColor = .BackColor Xor lblFillColor.BackColor
            End Select
          End If
          lngP2 = lngP1
      End Select
    End With
  ElseIf (Button = vbRightButton) Then
    If intActiveTool = conTPick Then
      lblFillColor.BackColor = picPaint.Point(X, Y)
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrDrawing
End Sub

Private Sub picPaint_MouseMove(Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  Dim intHscPaintValue As Integer       'adjusted horizontal and vertical scroll
  Dim intVscPaintValue As Integer       '                bar value for hand tool
  
  On Error GoTo ErrorHandler
  
  If Button = vbLeftButton Then
    If blnDrawing Then
      With picPaint
        Select Case intActiveTool
          Case conTAirBrush
            DrawAirBrush CInt(X), CInt(Y), .DrawWidth * 4
          Case conTArrow
            DrawArrow lngP1.X, lngP1.Y, lngP2.X, lngP2.Y
            AdjustP2 X:=X, Y:=Y, Shift:=Shift, blnEnableCtrl:=True
            DrawArrow lngP1.X, lngP1.Y, lngP2.X, lngP2.Y
          Case conTBrush
            .DrawMode = vbCopyPen
            .ForeColor = lblForeColor.BackColor
            .FillColor = lblForeColor.BackColor
            DrawBrush intBrushShape:=intBrushShape, X:=X, Y:=Y
          Case conTCurve
            DrawCurveBezier
            imgBezier(0).Top = imgBezier(0).Top + (Y - lngP1.Y)
            imgBezier(0).Left = imgBezier(0).Left + (X - lngP1.X)
            imgBezier(1).Top = imgBezier(1).Top + (Y - lngP1.Y)
            imgBezier(1).Left = imgBezier(1).Left + (X - lngP1.X)
            imgBezier(2).Top = imgBezier(2).Top + (Y - lngP1.Y)
            imgBezier(2).Left = imgBezier(2).Left + (X - lngP1.X)
            imgBezier(3).Top = imgBezier(3).Top + (Y - lngP1.Y)
            imgBezier(3).Left = imgBezier(3).Left + (X - lngP1.X)
            DrawCurveBezier
            lngP1.X = X
            lngP1.Y = Y
          Case conTEllipse
            If (lngP2.X <> lngP1.X) Then
              picPaint.Circle ((lngP1.X + lngP2.X) / 2, _
                                 (lngP1.Y + lngP2.Y) / 2), _
                               varIIf(Abs(lngP2.X - lngP1.X) > _
                                        Abs(lngP2.Y - lngP1.Y), _
                                      Abs(lngP2.X - lngP1.X) / 2, _
                                      Abs(lngP2.Y - lngP1.Y) / 2), , , , _
                               Abs((lngP2.Y - lngP1.Y) / _
                                   (lngP2.X - lngP1.X))
            End If
            AdjustP2 X:=X, Y:=Y, Shift:=Shift
            If (lngP2.X <> lngP1.X) Then
              picPaint.Circle ((lngP1.X + lngP2.X) / 2, _
                                 (lngP1.Y + lngP2.Y) / 2), _
                               varIIf(Abs(lngP2.X - lngP1.X) > _
                                        Abs(lngP2.Y - lngP1.Y), _
                                      Abs(lngP2.X - lngP1.X) / 2, _
                                      Abs(lngP2.Y - lngP1.Y) / 2), , , , _
                               Abs((lngP2.Y - lngP1.Y) / _
                                   (lngP2.X - lngP1.X))
            End If
          Case conTEraser
            picPaint.Line (X, Y)-(X + .DrawWidth, Y - .DrawWidth), , B
          Case conTFilter
            ImageFilter intFilter:=intActiveFilterTool, X:=CLng(X), Y:=CLng(Y)
          Case conTHand
            If blnDrag Then
              If hscPaint.Visible Then
                intHscPaintValue = lngDragStart.X - _
                                   (lngP1.X - (X + picPaint.Left))
                intHscPaintValue = hscPaint.Value + _
                                   ((picPaint.Left - intHscPaintValue) / _
                                    Screen.TwipsPerPixelX)
                If intHscPaintValue < hscPaint.Min Then
                  hscPaint.Value = hscPaint.Min
                ElseIf intHscPaintValue > hscPaint.Max Then
                  hscPaint.Value = hscPaint.Max
                Else
                  hscPaint.Value = intHscPaintValue
                End If
              End If
              If vscPaint.Visible Then
                intVscPaintValue = lngDragStart.Y - _
                                   (lngP1.Y - (Y + picPaint.Top))
                intVscPaintValue = vscPaint.Value + _
                                   ((picPaint.Top - intVscPaintValue) / _
                                    Screen.TwipsPerPixelY)
                If intVscPaintValue < vscPaint.Min Then
                  vscPaint.Value = vscPaint.Min
                ElseIf intVscPaintValue > vscPaint.Max Then
                  vscPaint.Value = vscPaint.Max
                Else
                  vscPaint.Value = intVscPaintValue
                End If
              End If
              picPaint.Refresh
            End If
          Case conTLine
            picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y)
            AdjustP2 X:=X, Y:=Y, Shift:=Shift, blnEnableCtrl:=True
            picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y)
          Case conTPencil
            lngP2 = lngP1
            lngP1.X = X
            lngP1.Y = Y
            picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y)
          Case conTPolygon
            If UBound(lngPolygon) = 0 Then
              ReDim Preserve lngPolygon(UBound(lngPolygon) + 1)
            Else
              DrawPolygon blnComplete:=False
            End If
            lngPolygon(UBound(lngPolygon)).X = X
            lngPolygon(UBound(lngPolygon)).Y = Y
            DrawPolygon blnComplete:=False
          Case conTRect
            If (lngP1.X <> lngP2.X) Or (lngP1.Y <> lngP2.Y) Then
              picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y), , B
            End If
            AdjustP2 X:=X, Y:=Y, Shift:=Shift
            picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y), , B
          Case conTRoundRect
            mdlAPI.RoundRect picPaint.hDC, _
                             lngP1.X, lngP1.Y, lngP2.X, lngP2.Y, 10, 10
            AdjustP2 X:=X, Y:=Y, Shift:=Shift
            mdlAPI.RoundRect picPaint.hDC, _
                             lngP1.X, lngP1.Y, lngP2.X, lngP2.Y, 10, 10
            .Refresh
          Case conTSelect
            picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y), _
                          vbBlack Xor picPaint.BackColor, B
            AdjustP2 X:=X, Y:=Y, Shift:=Shift
            picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y), _
                          vbBlack Xor picPaint.BackColor, B
        End Select
      End With
    End If
  End If
  UpdateStatusBar X:=X, Y:=Y
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_MouseUp(Button As Integer, _
                             Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  If Button = vbLeftButton Then
    If blnDrawing Then
      lngP2.X = X
      lngP2.Y = Y
      Select Case intActiveTool
        Case conTArrow, conTEllipse, conTLine, conTRect, conTRoundRect
          With picPaint
            .DrawMode = vbCopyPen
            If intActiveTool = conTLine Then
              .ForeColor = lblForeColor.BackColor
            Else
              .ForeColor = .BackColor Xor .ForeColor
              .FillColor = .BackColor Xor .FillColor
            End If
          End With
          Select Case intActiveTool
            Case conTArrow
              AdjustP2 X:=X, Y:=Y, Shift:=Shift, blnEnableCtrl:=True
              DrawArrow lngP1.X, lngP1.Y, lngP2.X, lngP2.Y
            Case conTEllipse
              AdjustP2 X:=X, Y:=Y, Shift:=Shift
              If (lngP2.X <> lngP1.X) Then
                picPaint.Circle ((lngP1.X + lngP2.X) / 2, _
                                   (lngP1.Y + lngP2.Y) / 2), _
                                 varIIf(Abs(lngP2.X - lngP1.X) > _
                                          Abs(lngP2.Y - lngP1.Y), _
                                        Abs(lngP2.X - lngP1.X) / 2, _
                                        Abs(lngP2.Y - lngP1.Y) / 2), , , , _
                                 Abs((lngP2.Y - lngP1.Y) / _
                                     (lngP2.X - lngP1.X))
              End If
            Case conTLine
              AdjustP2 X:=X, Y:=Y, Shift:=Shift, blnEnableCtrl:=True
              picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y)
            Case conTRect
              AdjustP2 X:=X, Y:=Y, Shift:=Shift
              If (lngP1.X <> lngP2.X) Or (lngP1.Y <> lngP2.Y) Then
                picPaint.Line (lngP1.X, lngP1.Y)- _
                              (lngP2.X, lngP2.Y), , B
              End If
            Case conTRoundRect
              AdjustP2 X:=X, Y:=Y, Shift:=Shift
              mdlAPI.RoundRect picPaint.hDC, _
                               lngP1.X, lngP1.Y, lngP2.X, lngP2.Y, 10, 10
          End Select
        Case conTHand
          blnDrag = False
          picPaint.ScaleMode = vbPixels
          picPaint.MouseIcon = LoadPicture(App.Path & "\Cursors\handflat.cur")
        Case conTSelect
          With picSelect
            If (Abs(lngP2.X - lngP1.X) > 1) And _
               (Abs(lngP2.Y - lngP1.Y) > 1) Then
              AdjustP2 X:=X, Y:=Y, Shift:=Shift
              .Width = Abs(lngP2.X - lngP1.X) - 1
              .Height = Abs(lngP2.Y - lngP1.Y) - 1
              .Left = IIf(lngP1.X <= lngP2.X, lngP1.X, lngP2.X) + 1
              .Top = IIf(lngP1.Y <= lngP2.Y, lngP1.Y, lngP2.Y) + 1
              .Visible = True
              .Picture = Nothing
              .PaintPicture picPaint.Image, 0, 0, _
                            .Width, .Height, .Left, .Top, .Width, .Height
              mnuCut.Enabled = True
              mnuCopy.Enabled = True
              mnuDelete.Enabled = True
              mnuCrop.Enabled = True
              blnFirstMoving = True
            Else
              .Visible = False
              picPaint.Line (lngP1.X, lngP1.Y)-(lngP2.X, lngP2.Y), _
                            vbBlack Xor picPaint.BackColor, B
              mnuCut.Enabled = False
              mnuCopy.Enabled = False
              mnuDelete.Enabled = False
              mnuCrop.Enabled = False
              blnFirstMoving = False
            End If
          End With
          picPaint.DrawWidth = intDot + 1
        Case conTZoom
          If sngZoomFactor = 1 Then
            picZoom.Width = picPaint.Width
            picZoom.Height = picPaint.Height
            picZoom.Picture = picPaint.Image
          End If
          If Shift <> vbCtrlMask Then
            'Zoom in
            If ((picZoom.Width * sngZoomFactor * conZoomFactor * 2) <= _
                (mdlEffect.conMaxImageWidth * 2)) And _
               ((picZoom.Height * sngZoomFactor * conZoomFactor * 2) <= _
                (mdlEffect.conMaxImageHeight * 2)) Then
              sngZoomFactor = sngZoomFactor * conZoomFactor
              ImageZoom X:=CLng(X * Screen.TwipsPerPixelX * conZoomFactor), _
                        Y:=CLng(Y * Screen.TwipsPerPixelY * conZoomFactor)
            End If
          Else
            'Zoom out
            sngZoomFactor = sngZoomFactor / conZoomFactor
            ImageZoom X:=CLng(X * Screen.TwipsPerPixelX / conZoomFactor), _
                      Y:=CLng(Y * Screen.TwipsPerPixelY / conZoomFactor)
          End If
      End Select
      blnDrawing = False
      If (intActiveTool <> conTText) And (intActiveTool <> conTSelect) And _
         (intActiveTool <> conTPolygon) And (intActiveTool <> conTCurve) And _
         (intActiveTool <> conTZoom) Then
        SetImageBuffer
      End If
    End If
  End If
  UpdateStatusBar
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_Resize()
  blnResize = True
End Sub

Private Sub picPaintResize_MouseDown(Index As Integer, Button As Integer, _
                                     Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'Start the drag operation on picPaintResize(Index)
  lngDragStart.X = CLng(X)
  lngDragStart.Y = CLng(Y)
  blnDrag = True
  blnResize = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaintResize_MouseMove(Index As Integer, Button As Integer, _
                                     Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'Move picPaintResize(Index) for the drag operation and resize the paint area
  '  to match picPaintResize(Index) position
  If blnDrag Then
    With picPaintResize(Index)
      If Index <> conResizeNS Then
        If (picPaint.Width + (X - lngDragStart.X)) > 0 Then
          .Left = .Left + (X - lngDragStart.X)
          picPaint.Width = picPaint.Width + (X - lngDragStart.X)
        End If
      End If
      If Index <> conResizeWE Then
        If (picPaint.Height + (Y - lngDragStart.Y)) > 0 Then
          .Top = .Top + (Y - lngDragStart.Y)
          picPaint.Height = picPaint.Height + (Y - lngDragStart.Y)
        End If
      End If
    End With
    AdjustPaintResizeBox
    Form_Resize
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaintResize_MouseUp(Index As Integer, Button As Integer, _
                                   Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'End the drag operation on picPaintResize(Index)
  blnDrag = False
  If blnResize Then
    SetImageBuffer
  End If
  blnResize = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picSelect_MouseDown(Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  If Button = vbLeftButton Then
    'Start the drag operation on the selection object
    blnMoving = True
    With picSelect
      picPaint.DrawWidth = 1
      If blnFirstMoving And (Shift <> vbCtrlMask) Then
        'Erase the drawing behind the selection object
        picPaint.DrawStyle = intDrawStyle
        picPaint.DrawMode = vbCopyPen
        picPaint.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height - 1), _
                      picPaint.BackColor, BF
        blnFirstMoving = False
      End If
      picPaint.DrawStyle = vbDot
      picPaint.DrawMode = vbXorPen
      picPaint.Line (.Left - 1, .Top - 1)- _
                    (.Left + .Width, .Top + .Height), _
                    vbBlack Xor picPaint.BackColor, B
      lngP1.X = X
      lngP1.Y = Y
    End With
  End If
  UpdateStatusBar
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picSelect_MouseMove(Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'Move the selection object for the drag operation
  If (Button = vbLeftButton) And blnMoving Then
    lngP2.X = X
    lngP2.Y = Y
    picSelect.Left = picSelect.Left + (lngP2.X - lngP1.X)
    picSelect.Top = picSelect.Top + (lngP2.Y - lngP1.Y)
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picSelect_MouseUp(Button As Integer, _
                              Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler
  
  'End the drag operation on picSelect
  If Button = vbLeftButton Then
    With picSelect
      picPaint.Line (.Left - 1, .Top - 1)- _
                    (.Left + .Width, .Top + .Height), _
                    vbBlack Xor picPaint.BackColor, B
    End With
    blnFirstMoving = False
    blnMoving = False
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Place selection image in picSelect to the paint area
' Assumptions: * These components exists in this form:
'                - picPaint
'                - picSelect
'              * Meet assumptions in this procedure:
'                  SetImageBuffer
' Effects    : * picSelect.Visible = False
'              * The selection rectangle has been erased
'              * Effects from SetImageBUffer
'              * Menu "Delete" is not enabled
' Inputs     : -
' Returns    : -
Private Sub PlaceSelection()
  On Error GoTo ErrorHandler

  With picSelect
    If .Visible Then
      .Visible = False
      picPaint.PaintPicture .Image, .Left, .Top
      'Erase the selection rectangle
      picPaint.DrawMode = vbXorPen
      picPaint.DrawWidth = 1
      picPaint.Line (.Left - 1, .Top - 1)-(.Left + .Width, .Top + .Height), _
                    vbBlack Xor picPaint.BackColor, B
      If Not blnFirstMoving Then
        SetImageBuffer
      End If
      mnuCopy.Enabled = False
      mnuCut.Enabled = False
      mnuCrop.Enabled = False
      mnuDelete.Enabled = False
    End If
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Safe the image on paint area (picPaint) to image buffer array
'              (picBuffer)
' Assumptions: These components exists in this form:
'                - picPaint
'                - picBuffer
' Effects    : * Image in paint area has been saved to image buffer array
'              * Buffer pointer (intBufStart) has been set to the next buffer
' Inputs     : -
' Returns    : -
Public Sub SetImageBuffer()
  On Error GoTo ErrorHandler

  If intBufCur < conBufMax Then
    intBufCur = intBufCur + 1
  Else
    intBufCur = 0
  End If
  If intBufCur > picBuffer.UBound Then
    Load picBuffer(intBufCur)
  End If
  picBuffer(intBufCur).Picture = picPaint.Image
  picBuffer(intBufCur).Tag = CStr((picPaint.Width * 100000) + picPaint.Height)
  intBufEnd = intBufCur
  If intBufStart = intBufEnd Then
    If intBufStart < conBufMax Then
      intBufStart = intBufStart + 1
    Else
      intBufStart = 0
    End If
  End If
  blnPicChanged = True
  mnuUndo.Enabled = True
  mnuRedo.Enabled = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub txtText_DblClick()
  On Error GoTo ErrorHandler
  
  With txtText
    picPaint.CurrentX = .Left
    picPaint.CurrentY = .Top
    picPaint.ForeColor = lblForeColor.BackColor
    picPaint.Print .Text
    .Visible = False
    SetImageBuffer
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorHandler
  
  With txtText
    lblTextSize.Caption = .Text & "M"
    .Width = lblTextSize.Width
    .Height = lblTextSize.Height
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorHandler
  
  With txtText
    Select Case KeyCode
      Case vbKeyReturn
        txtText_DblClick
      Case vbKeyEscape
        .Visible = False
      Case Else
        lblTextSize.Caption = .Text & "O"
        .Width = lblTextSize.Width
        .Height = lblTextSize.Height
    End Select
    If Not .Visible Then
      picPaint.SetFocus
    End If
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub txtText_LostFocus()
  On Error GoTo ErrorHandler
  
  With txtText
    If (.Visible) And (.Tag <> "moving") Then
      txtText_KeyUp vbKeyReturn, 0
    End If
    .Tag = ""
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Update certain drawing to match drawing properties changes
'              (dot width, foreground color, etc.)
' Assumption : This global variable has been initiated:
'                intActiveTool
' Effect     : As specified
' Inputs     : -
' Returns    : -
Private Sub UpdateDrawing()
  On Error GoTo ErrorHandler
  
  Select Case intActiveTool
    Case conTCurve
      DrawCurveBezier
    Case conTPolygon
      If blnDrawingPolygon Then
        DrawPolygon blnComplete:=False, blnOnlyDrawLastLine:=False
      End If
  End Select
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Update form's title based on file name strFileName
' Assumptions: This global variable has been initiated:
'                strFileName
'              This global constant has been initiated:
'                conProgramTitle (the title of this program)
' Effects    : The form's title has been updated
' Inputs     : -
' Returns    : -
Private Sub UpdateFormTitle()
  On Error GoTo ErrorHandler
  
  If strFileName <> "" Then
    Me.Caption = strFileName & " - " & conProgramTitle
  Else
    Me.Caption = "untitled - " & conProgramTitle
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Update the status bar sta content
' Assumptions: These components exists in this form:
'                sta, imgBezier()
'              This global variable has been initiated:
'                intActiveTool, blnFirstMoving, blnDrawing, blnDrawingPolygon
' Effect     : The status bar content has been updated
' Inputs     : * intArea (the new content of the status bar)
'              * X, Y (current drawing coordinates)
'              * blnClear (condition to remove all texts in status bar)
' Return     : -
Public Sub UpdateStatusBar(Optional intInfo As enmStatusBar = conStPaintArea, _
                           Optional X As Single, Optional Y As Single, _
                           Optional intPercentage As Integer, _
                           Optional blnClear As Boolean = False)
  On Error GoTo ErrorHandler
  
  If blnClear Then
    sta.Panels(1).Text = ""
    sta.Panels(2).Text = ""
    sta.Panels(3).Text = ""
  Else
    'First panel
    With sta.Panels(1)
      Select Case intInfo
        Case conStPaintArea
          Select Case intActiveTool
            Case conTAirBrush
              .Text = "Draws using an airbrush with the selected airbrush size"
            Case conTArrow
              If Not blnDrawing Then
                .Text = "Draws an arrow with the selected arrow width"
              Else
                .Text = "Press and hold down " & _
                        "CTRL to draw a horizontal or vertical arrow; " & _
                        "SHIFT to draw a 45-degree arrow"
              End If
            Case conTBrush
              .Text = "Draws using a brush with the selected shape"
            Case conTCurve
              If Not imgBezier(0).Visible Then
                .Text = "Draws a bezier curve with the selected curve width"
              Else
                .Text = "Press ENTER or double-click " & _
                        "to finish drawing the curve"
              End If
            Case conTEllipse
              If Not blnDrawing Then
                .Text = "Draws an ellips " & _
                        "with the selected outline width and fill style"
              Else
                .Text = "Press and hold down SHIFT to draw a circle"
              End If
            Case conTEraser
              .Text = "Erases a partion of the picture " & _
                      "using the selected eraser width"
            Case conTFilter
              .Text = "Apply the selected filter to the image"
            Case conTFill
              .Text = "Fills an area"
            Case conTHand
              .Text = "Pan to see other part of the picture"
            Case conTLine
              If Not blnDrawing Then
                .Text = "Draws a straight line with the selected line width"
              Else
                .Text = "Press and hold down " & _
                        "CTRL to draw a horizontal or vertical line; " & _
                        "SHIFT to draw a 45-degree line"
              End If
            Case conTPencil
              .Text = "Draws using a pencil with the selected dot size"
            Case conTPick
              .Text = "Picks up a foreground color (click) or " & _
                      "background color (right-click) " & _
                      "from the picture for drawing"
            Case conTPolygon
              If Not blnDrawingPolygon Then
                .Text = "Draws a polygon " & _
                        "with the selected outline width and fill area"
              Else
                .Text = "Press ENTER or double-click " & _
                        "to finish drawing the polygon"
              End If
            Case conTRect
              If Not blnDrawing Then
                .Text = "Draws a rectangle " & _
                        "with the selected outline width and fill style"
              Else
                .Text = "Press and hold down SHIFT to draw a square"
              End If
            Case conTRoundRect
              If Not blnDrawing Then
                .Text = "Draws a rounded rectangle " & _
                        "with the selected outline width and fill style"
              Else
                .Text = "Press and hold down SHIFT to draw a rounded-square"
              End If
            Case conTSelect
              If blnFirstMoving Then
                .Text = "Press and hold down CTRL " & _
                        "before moving the selection to copy it"
              ElseIf Not blnDrawing Then
                .Text = "Selects a rectangular part of the picture " & _
                        "to move or delete"
              Else
                .Text = "Press and hold down SHIFT to select a square part"
              End If
            Case conTText
              If Not txtText.Visible Then
                .Text = "Insert text into the picture"
              Else
                .Text = "Press ENTER or double-click " & _
                        "to finish inserting the text"
              End If
            Case conTZoom
              .Text = "Zoom in or zoom out the image 1.25x " & _
                      "(press and hold down CTRL to zoom out)"
          End Select
        Case conStColorBox
          .Text = "Click to set the foreground color; " & _
                               "Right-click to set the background color"
        Case conStForeColorBox
          .Text = "Double-click " & _
                  "to set the foreground color with custom color"
        Case conStBackColorBox
          .Text = "Double-click " & _
                  "to set the background color with custom color"
        Case conStFiltering
          .Text = "Filtering... " & _
                 "(" & CStr(intPercentage) & "% complete)"
        Case conStRetrieveingColor
          .Text = "Retrieving color information... " & _
                  "(" & CStr(intPercentage) & "% complete)"
        Case Else
          .Text = ""
      End Select
    End With
    'Second and third panels
    If intInfo = conStPaintArea Then
      If blnDrawing And _
         ((intActiveTool = conTArrow) Or (intActiveTool = conTEllipse) Or _
          (intActiveTool = conTLine) Or (intActiveTool = conTRect) Or _
          (intActiveTool = conTRoundRect) Or (intActiveTool = conTSelect)) Then
        sta.Panels(3).Text = CStr(lngP2.X - lngP1.X) & "x" & _
                             CStr(lngP2.Y - lngP1.Y)
      Else
        sta.Panels(2).Text = CStr(X) & "," & CStr(Y)
        sta.Panels(3).Text = ""
      End If
    Else
      sta.Panels(2).Text = ""
      sta.Panels(3).Text = ""
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub vscPaint_Change()
  Dim lngPicPaintTop As Long
  
  On Error GoTo ErrorHandler
  
  lngPicPaintTop = -(CLng(vscPaint.Value) * 10)
  picPaint.Top = lngPicPaintTop
  AdjustPaintResizeBox
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub
