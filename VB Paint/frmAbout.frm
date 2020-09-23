VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VB Paint"
   ClientHeight    =   2730
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5415
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -105
      TabIndex        =   3
      Top             =   2025
      WhatsThisHelpID =   10385
      Width           =   5550
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3945
      TabIndex        =   0
      Top             =   2250
      WhatsThisHelpID =   10379
      Width           =   1260
   End
   Begin VB.Label Label4 
      Caption         =   " Thank you for using this program."
      Height          =   255
      Left            =   315
      TabIndex        =   7
      Top             =   1560
      Width           =   2505
   End
   Begin VB.Label lblEMail 
      Caption         =   "(theo_yz@yahoo.com)"
      Height          =   255
      Left            =   3390
      MouseIcon       =   "frmAbout.frx":1042
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   645
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "This program is freeware and open source. You may freely use and modify any part of the code for your personal needs."
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   360
      TabIndex        =   4
      Top             =   1155
      Width           =   4740
   End
   Begin VB.Label Label1 
      Caption         =   "VB Paint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1425
      TabIndex        =   1
      Top             =   285
      WhatsThisHelpID =   10382
      Width           =   4605
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Written by Theo Zacharias"
      Height          =   225
      Left            =   1440
      TabIndex        =   2
      Top             =   645
      WhatsThisHelpID =   10383
      Width           =   1965
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   255
      Picture         =   "frmAbout.frx":134C
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   255
      TabIndex        =   5
      Top             =   1065
      Width           =   4935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, X As Single, Y As Single)
  lblEMail.Font.Underline = False
End Sub

Private Sub lblEMail_Click()
  mdlAPI.ShellExecute hwnd:=Me.hwnd, lpOperation:=vbNullString, _
                      lpFile:="mailto:theo_yz@yahoo.com", _
                      lpParameters:=vbNullString, _
                      lpDirectory:=vbNullString, nShowCmd:=1
End Sub

Private Sub lblEMail_MouseMove(Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  lblEMail.Font.Underline = True
End Sub
