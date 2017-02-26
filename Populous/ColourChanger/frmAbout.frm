VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Populous Level Changer"
   ClientHeight    =   2310
   ClientLeft      =   7140
   ClientTop       =   5925
   ClientWidth     =   6195
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1594.403
   ScaleMode       =   0  'User
   ScaleWidth      =   5817.425
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4800
      TabIndex        =   0
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "ted@brambles.org"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5746.997
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Label lblDescription 
      Caption         =   "Copies Populous level files to allow single player levels to be played as multiplayer."
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Populous Level Changer"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5746.997
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 0.1"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Copyright 2003 Ted John"
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub

