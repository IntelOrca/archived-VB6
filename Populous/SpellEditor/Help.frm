VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Populous Spell Editor Help"
   ClientHeight    =   2565
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "ted@brambles.org"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Please report any bugs to me."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Help.frx":0000
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
End Sub
