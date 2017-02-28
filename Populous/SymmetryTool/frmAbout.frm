VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Populous Symmetry Tool"
   ClientHeight    =   3075
   ClientLeft      =   10575
   ClientTop       =   8805
   ClientWidth     =   5865
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About PopulousSymmetryTool"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   780
      Left            =   120
      Picture         =   "frmAbout.frx":1CCA
      ScaleHeight     =   720
      ScaleMode       =   0  'User
      ScaleWidth      =   720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2640
      Width           =   1467
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":3994
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   1050
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   1125
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      Caption         =   "Populous Symmetry Tool"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Tag             =   "Application Title"
      Top             =   240
      Width           =   4092
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.5"
      Height          =   225
      Left            =   1050
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   780
      Width           =   4092
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Copyright 2005 TedTycoon"
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Tag             =   "Warning: ..."
      Top             =   2640
      Width           =   3870
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

