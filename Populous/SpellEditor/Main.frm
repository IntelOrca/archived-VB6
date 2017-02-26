VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   7770
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7095
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Script Filenames (cpscr???.dat)"
      Height          =   855
      Left            =   3960
      TabIndex        =   24
      Top             =   6840
      Width           =   3015
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Mason"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "000"
         Mask            =   "###"
         PromptChar      =   "0"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Mason"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "000"
         Mask            =   "###"
         PromptChar      =   "0"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Mason"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "000"
         Mask            =   "###"
         PromptChar      =   "0"
      End
      Begin VB.Label Label2 
         Caption         =   "   Dakini          Chumara         Matak"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tribes available to level"
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   6840
      Width           =   3735
      Begin VB.CommandButton Tribe 
         Caption         =   "Blue"
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "Main.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Tribe 
         Caption         =   "Matak"
         Height          =   495
         Index           =   3
         Left            =   2880
         Picture         =   "Main.frx":0C6C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Tribe 
         Caption         =   "Chumara"
         Height          =   495
         Index           =   2
         Left            =   1965
         Picture         =   "Main.frx":100E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Tribe 
         Caption         =   "Dakini"
         Height          =   495
         Index           =   1
         Left            =   1035
         Picture         =   "Main.frx":13F8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Landscape"
      Height          =   3615
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   2895
      Begin VB.CommandButton Command6 
         Caption         =   "NEXT LANDSCAPE"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Change colour scheme"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "PREVIOUS LANDSCAPE"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Change colour scheme"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "   NEXT   TREE STYLE"
         Height          =   615
         Left            =   1560
         TabIndex        =   10
         ToolTipText     =   "Change colour scheme"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "PREVIOUS TREE STYLE"
         Height          =   615
         Left            =   1560
         TabIndex        =   12
         ToolTipText     =   "Change colour scheme"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Mason"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Mason"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Image Landscape 
         Height          =   945
         Left            =   120
         Picture         =   "Main.frx":179A
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1185
      End
      Begin VB.Image Tree 
         Height          =   945
         Left            =   1560
         Picture         =   "Main.frx":8FB1
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1185
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FOG is OFF"
      Height          =   975
      Left            =   4080
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Huts"
      Height          =   3615
      Left            =   3120
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
      Begin VB.Image Image15 
         Height          =   735
         Left            =   960
         Picture         =   "Main.frx":15C38
         Tag             =   "2"
         ToolTipText     =   "Balloon Hut"
         Top             =   2760
         Width           =   660
      End
      Begin VB.Image Image18 
         Height          =   735
         Left            =   120
         Picture         =   "Main.frx":163C2
         Tag             =   "2"
         ToolTipText     =   "Boat Hut"
         Top             =   2760
         Width           =   660
      End
      Begin VB.Image Image19 
         Height          =   735
         Left            =   960
         Picture         =   "Main.frx":16B61
         Tag             =   "2"
         ToolTipText     =   "Firewarrior Hut"
         Top             =   1920
         Width           =   660
      End
      Begin VB.Image Image20 
         Height          =   735
         Left            =   120
         Picture         =   "Main.frx":17308
         Tag             =   "2"
         ToolTipText     =   "Spy training hut"
         Top             =   1920
         Width           =   660
      End
      Begin VB.Image Image21 
         Height          =   735
         Left            =   960
         Picture         =   "Main.frx":17A30
         Tag             =   "2"
         ToolTipText     =   "Temple"
         Top             =   1080
         Width           =   660
      End
      Begin VB.Image Image22 
         Height          =   735
         Left            =   120
         Picture         =   "Main.frx":18129
         Tag             =   "2"
         ToolTipText     =   "Warrior hut"
         Top             =   1080
         Width           =   660
      End
      Begin VB.Image Image23 
         Height          =   735
         Left            =   960
         Picture         =   "Main.frx":1889D
         Tag             =   "2"
         ToolTipText     =   "Guard Tower"
         Top             =   240
         Width           =   660
      End
      Begin VB.Image Image24 
         Height          =   735
         Left            =   120
         Picture         =   "Main.frx":1903F
         Tag             =   "2"
         ToolTipText     =   "Hut"
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Spells"
      Height          =   4815
      Left            =   5040
      TabIndex        =   13
      Top             =   1920
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   630
         Left            =   720
         Picture         =   "Main.frx":197F5
         Tag             =   "2"
         ToolTipText     =   "Angle of death"
         Top             =   360
         Width           =   450
      End
      Begin VB.Image Image2 
         Height          =   630
         Left            =   1320
         Picture         =   "Main.frx":19DAD
         Tag             =   "2"
         ToolTipText     =   "Volcano"
         Top             =   360
         Width           =   450
      End
      Begin VB.Image Image3 
         Height          =   630
         Left            =   720
         Picture         =   "Main.frx":1A393
         Tag             =   "2"
         ToolTipText     =   "Earthquake"
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image Image4 
         Height          =   630
         Left            =   1320
         Picture         =   "Main.frx":1A91D
         Tag             =   "2"
         ToolTipText     =   "Erode"
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image Image5 
         Height          =   630
         Left            =   720
         Picture         =   "Main.frx":1AEA9
         Tag             =   "2"
         ToolTipText     =   "Tornado"
         Top             =   1800
         Width           =   450
      End
      Begin VB.Image Image6 
         Height          =   630
         Left            =   1320
         Picture         =   "Main.frx":1B464
         Tag             =   "2"
         ToolTipText     =   "Swamp"
         Top             =   1800
         Width           =   450
      End
      Begin VB.Image Image7 
         Height          =   630
         Left            =   120
         Picture         =   "Main.frx":1BA17
         Tag             =   "2"
         ToolTipText     =   "Firestorm"
         Top             =   360
         Width           =   450
      End
      Begin VB.Image Image8 
         Height          =   630
         Left            =   120
         Picture         =   "Main.frx":1BFAB
         Tag             =   "2"
         ToolTipText     =   "Flatten"
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image Image9 
         Height          =   630
         Left            =   720
         Picture         =   "Main.frx":1C53A
         Tag             =   "2"
         ToolTipText     =   "Landbridge"
         Top             =   2520
         Width           =   450
      End
      Begin VB.Image Image10 
         Height          =   630
         Left            =   1320
         Picture         =   "Main.frx":1CAD4
         Tag             =   "2"
         ToolTipText     =   "Lightning"
         Top             =   2520
         Width           =   450
      End
      Begin VB.Image Image11 
         Height          =   630
         Left            =   720
         Picture         =   "Main.frx":1D07B
         Tag             =   "2"
         ToolTipText     =   "Swarm"
         Top             =   3240
         Width           =   450
      End
      Begin VB.Image Image12 
         Height          =   630
         Left            =   1320
         Picture         =   "Main.frx":1D62E
         Tag             =   "2"
         ToolTipText     =   "Invisibility"
         Top             =   3240
         Width           =   450
      End
      Begin VB.Image Image13 
         Height          =   630
         Left            =   120
         Picture         =   "Main.frx":1DBD6
         Tag             =   "2"
         ToolTipText     =   "Hypnotise"
         Top             =   1800
         Width           =   450
      End
      Begin VB.Image Image14 
         Height          =   630
         Left            =   120
         Picture         =   "Main.frx":1E1C4
         Tag             =   "2"
         ToolTipText     =   "Magical Shield"
         Top             =   2520
         Width           =   450
      End
      Begin VB.Image Nothing1 
         Height          =   630
         Left            =   120
         Picture         =   "Main.frx":1E7A9
         ToolTipText     =   "Guest Spells (Cannot be saved)"
         Top             =   3960
         Width           =   450
      End
      Begin VB.Image Image16 
         Height          =   630
         Left            =   720
         Picture         =   "Main.frx":1EB7A
         Tag             =   "2"
         ToolTipText     =   "Blast"
         Top             =   3960
         Width           =   450
      End
      Begin VB.Image Image17 
         Height          =   630
         Left            =   1320
         Picture         =   "Main.frx":1F14A
         Tag             =   "2"
         ToolTipText     =   "Convert"
         Top             =   3960
         Width           =   450
      End
      Begin VB.Image Image52 
         Height          =   630
         Left            =   120
         Picture         =   "Main.frx":1F753
         Tag             =   "2"
         ToolTipText     =   "Ghost Army"
         Top             =   3240
         Width           =   450
      End
      Begin VB.Image Nothing2 
         Height          =   630
         Left            =   120
         Picture         =   "Main.frx":1FC0C
         Top             =   3960
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Allies"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
      Begin VB.CheckBox RY 
         Caption         =   "Red + Yellow"
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   1400
      End
      Begin VB.CheckBox RG 
         Caption         =   "Red + Green"
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   540
         Width           =   1400
      End
      Begin VB.CheckBox YG 
         Caption         =   "Yellow + Green"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1400
      End
      Begin VB.CheckBox BG 
         Caption         =   "Blue + Green"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1300
      End
      Begin VB.CheckBox BY 
         Caption         =   "Blue + Yellow"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   1300
      End
      Begin VB.CheckBox BR 
         Caption         =   "Blue + Red"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1300
      End
   End
   Begin MSComDlg.CommonDialog Open1 
      Left            =   9840
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GOD is OFF"
      Height          =   975
      Left            =   3120
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.Image TribeImage 
      Height          =   180
      Index           =   4
      Left            =   16920
      Top             =   1560
      Width           =   345
   End
   Begin VB.Image TribeImage 
      Height          =   180
      Index           =   3
      Left            =   16920
      Picture         =   "Main.frx":1FFDD
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image TribeImage 
      Height          =   195
      Index           =   2
      Left            =   16920
      Picture         =   "Main.frx":2037F
      Top             =   840
      Width           =   360
   End
   Begin VB.Image TribeImage 
      Height          =   180
      Index           =   1
      Left            =   16920
      Picture         =   "Main.frx":20769
      Top             =   480
      Width           =   360
   End
   Begin VB.Image TribeImage 
      Height          =   180
      Index           =   0
      Left            =   16920
      Picture         =   "Main.frx":20B0B
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Tribes 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Mason"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16920
      TabIndex        =   23
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image Image54 
      Height          =   780
      Left            =   4320
      Picture         =   "Main.frx":20EAD
      Top             =   360
      Width           =   660
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   35
      Left            =   15480
      Picture         =   "Main.frx":2244D
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   34
      Left            =   15480
      Picture         =   "Main.frx":2C2EF
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1020
   End
   Begin VB.Image Image53 
      Height          =   630
      Left            =   7320
      Picture         =   "Main.frx":36346
      ToolTipText     =   "Firestorm"
      Top             =   4440
      Width           =   450
   End
   Begin VB.Image Treepic 
      Height          =   945
      Index           =   6
      Left            =   17400
      Picture         =   "Main.frx":367FF
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Image Treepic 
      Height          =   945
      Index           =   5
      Left            =   16080
      Picture         =   "Main.frx":42E9C
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Image Treepic 
      Height          =   945
      Index           =   4
      Left            =   14760
      Picture         =   "Main.frx":50073
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Image Treepic 
      Height          =   945
      Index           =   3
      Left            =   13440
      Picture         =   "Main.frx":5CC63
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Image Treepic 
      Height          =   945
      Index           =   2
      Left            =   12120
      Picture         =   "Main.frx":69D29
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Image Treepic 
      Height          =   945
      Index           =   1
      Left            =   10800
      Picture         =   "Main.frx":769B0
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Image Treepic 
      Height          =   945
      Index           =   0
      Left            =   9480
      Picture         =   "Main.frx":83637
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Image Image51 
      Height          =   630
      Left            =   1560
      Picture         =   "Main.frx":902BE
      Top             =   480
      Width           =   420
   End
   Begin VB.Image Image50 
      BorderStyle     =   1  'Fixed Single
      Height          =   1755
      Left            =   120
      Picture         =   "Main.frx":90FA7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   33
      Left            =   15480
      Picture         =   "Main.frx":B1DEB
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   32
      Left            =   15480
      Picture         =   "Main.frx":BBC45
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   31
      Left            =   15480
      Picture         =   "Main.frx":C573D
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   30
      Left            =   15480
      Picture         =   "Main.frx":CF513
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   29
      Left            =   14280
      Picture         =   "Main.frx":D77AA
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   28
      Left            =   14280
      Picture         =   "Main.frx":E0D22
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   27
      Left            =   14280
      Picture         =   "Main.frx":E8D57
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   26
      Left            =   14280
      Picture         =   "Main.frx":F0F17
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   25
      Left            =   14280
      Picture         =   "Main.frx":F8C06
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   24
      Left            =   14280
      Picture         =   "Main.frx":10293C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   23
      Left            =   13080
      Picture         =   "Main.frx":10B412
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   22
      Left            =   13080
      Picture         =   "Main.frx":1133E0
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   21
      Left            =   13080
      Picture         =   "Main.frx":11B0AE
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   20
      Left            =   13080
      Picture         =   "Main.frx":1258A8
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   19
      Left            =   13080
      Picture         =   "Main.frx":130251
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   18
      Left            =   13080
      Picture         =   "Main.frx":139378
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   17
      Left            =   11880
      Picture         =   "Main.frx":140EFE
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   16
      Left            =   11880
      Picture         =   "Main.frx":14A452
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   15
      Left            =   11880
      Picture         =   "Main.frx":154BBF
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   14
      Left            =   11880
      Picture         =   "Main.frx":15D400
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   13
      Left            =   11880
      Picture         =   "Main.frx":167098
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   12
      Left            =   11880
      Picture         =   "Main.frx":17003E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   11
      Left            =   10680
      Picture         =   "Main.frx":17A05E
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   10
      Left            =   10680
      Picture         =   "Main.frx":184655
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   9
      Left            =   10680
      Picture         =   "Main.frx":18C40E
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   8
      Left            =   10680
      Picture         =   "Main.frx":194862
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   7
      Left            =   10680
      Picture         =   "Main.frx":19FA9B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   6
      Left            =   10680
      Picture         =   "Main.frx":1A6ED7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   5
      Left            =   9480
      Picture         =   "Main.frx":1B0287
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   4
      Left            =   9480
      Picture         =   "Main.frx":1BAFAA
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   3
      Left            =   9480
      Picture         =   "Main.frx":1C54ED
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   2
      Left            =   9480
      Picture         =   "Main.frx":1CDFAD
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   1
      Left            =   9480
      Picture         =   "Main.frx":1D938D
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Image Image49 
      Height          =   840
      Index           =   0
      Left            =   9480
      Picture         =   "Main.frx":1E2ABB
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1020
   End
   Begin VB.Image Nothing3 
      Height          =   735
      Left            =   9000
      Picture         =   "Main.frx":1EA2D2
      Top             =   7800
      Width           =   660
   End
   Begin VB.Image Image48 
      Height          =   630
      Left            =   7680
      Picture         =   "Main.frx":1EA8C2
      Top             =   840
      Width           =   450
   End
   Begin VB.Image Image47 
      Height          =   630
      Left            =   8280
      Picture         =   "Main.frx":1EAE7A
      Top             =   840
      Width           =   450
   End
   Begin VB.Image Image46 
      Height          =   630
      Left            =   7320
      Picture         =   "Main.frx":1EB460
      Top             =   1560
      Width           =   450
   End
   Begin VB.Image Image45 
      Height          =   630
      Left            =   7920
      Picture         =   "Main.frx":1EB9EA
      Top             =   1560
      Width           =   450
   End
   Begin VB.Image Image44 
      Height          =   630
      Left            =   7320
      Picture         =   "Main.frx":1EBF76
      Top             =   2280
      Width           =   450
   End
   Begin VB.Image Image43 
      Height          =   630
      Left            =   7920
      Picture         =   "Main.frx":1EC531
      Top             =   2280
      Width           =   450
   End
   Begin VB.Image Image42 
      Height          =   630
      Left            =   8520
      Picture         =   "Main.frx":1ECAE4
      Top             =   1560
      Width           =   450
   End
   Begin VB.Image Image41 
      Height          =   630
      Left            =   8520
      Picture         =   "Main.frx":1ED078
      Top             =   2280
      Width           =   450
   End
   Begin VB.Image Image40 
      Height          =   630
      Left            =   7320
      Picture         =   "Main.frx":1ED607
      Top             =   3000
      Width           =   450
   End
   Begin VB.Image Image39 
      Height          =   630
      Left            =   7920
      Picture         =   "Main.frx":1EDBA1
      Top             =   3000
      Width           =   450
   End
   Begin VB.Image Image38 
      Height          =   630
      Left            =   7320
      Picture         =   "Main.frx":1EE148
      Top             =   3720
      Width           =   450
   End
   Begin VB.Image Image37 
      Height          =   630
      Left            =   7920
      Picture         =   "Main.frx":1EE6FB
      Top             =   3720
      Width           =   450
   End
   Begin VB.Image Image36 
      Height          =   630
      Left            =   8520
      Picture         =   "Main.frx":1EECA3
      Top             =   3000
      Width           =   450
   End
   Begin VB.Image Image35 
      Height          =   630
      Left            =   8520
      Picture         =   "Main.frx":1EF291
      Top             =   3720
      Width           =   450
   End
   Begin VB.Image Image34 
      Height          =   630
      Left            =   7920
      Picture         =   "Main.frx":1EF876
      Top             =   4440
      Width           =   450
   End
   Begin VB.Image Image33 
      Height          =   630
      Left            =   8520
      Picture         =   "Main.frx":1EFE46
      Top             =   4440
      Width           =   450
   End
   Begin VB.Image Image32 
      Height          =   735
      Left            =   8280
      Picture         =   "Main.frx":1F044F
      Top             =   7800
      Width           =   660
   End
   Begin VB.Image Image31 
      Height          =   735
      Left            =   7320
      Picture         =   "Main.frx":1F0BD9
      Top             =   7800
      Width           =   660
   End
   Begin VB.Image Image30 
      Height          =   735
      Left            =   8280
      Picture         =   "Main.frx":1F1378
      Top             =   6960
      Width           =   660
   End
   Begin VB.Image Image29 
      Height          =   735
      Left            =   7320
      Picture         =   "Main.frx":1F1B1F
      Top             =   6960
      Width           =   660
   End
   Begin VB.Image Image28 
      Height          =   735
      Left            =   8280
      Picture         =   "Main.frx":1F2247
      Top             =   6120
      Width           =   660
   End
   Begin VB.Image Image27 
      Height          =   735
      Left            =   7320
      Picture         =   "Main.frx":1F2940
      Top             =   6120
      Width           =   660
   End
   Begin VB.Image Image26 
      Height          =   735
      Left            =   8280
      Picture         =   "Main.frx":1F30B4
      Top             =   5280
      Width           =   660
   End
   Begin VB.Image Image25 
      Height          =   735
      Left            =   7320
      Picture         =   "Main.frx":1F3856
      Top             =   5280
      Width           =   660
   End
   Begin VB.Image Teleport 
      Height          =   630
      Left            =   8400
      Picture         =   "Main.frx":1F400C
      Top             =   120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Armageddon 
      Height          =   630
      Left            =   7800
      Picture         =   "Main.frx":1F4509
      Top             =   120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Bloodlust 
      Height          =   630
      Left            =   7200
      Picture         =   "Main.frx":1F49C0
      Top             =   120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Menu File1 
      Caption         =   "&File"
      Begin VB.Menu Save1 
         Caption         =   "Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu Load1 
         Caption         =   "Load"
      End
      Begin VB.Menu Line 
         Caption         =   "-"
      End
      Begin VB.Menu Exit1 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help1 
      Caption         =   "&Help"
      Begin VB.Menu Help2 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu About1 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public hfile As String


Private Sub about1_Click()
frmAbout.Show
End Sub

Private Sub BG_Click()
If Not BG.Value = 0 Then
  Mid(hfile, 93, 1) = Chr(Asc(Mid(hfile, 93, 1)) Or 8)
  Mid(hfile, 96, 1) = Chr(Asc(Mid(hfile, 96, 1)) Or 1)
Else
  Mid(hfile, 93, 1) = Chr(Asc(Mid(hfile, 93, 1)) And Not 8)
  Mid(hfile, 96, 1) = Chr(Asc(Mid(hfile, 96, 1)) And Not 1)
End If
End Sub

Private Sub BR_Click()
If Not BR.Value = 0 Then
  Mid(hfile, 93, 1) = Chr(Asc(Mid(hfile, 93, 1)) Or 2)
  Mid(hfile, 94, 1) = Chr(Asc(Mid(hfile, 94, 1)) Or 1)
Else
  Mid(hfile, 93, 1) = Chr(Asc(Mid(hfile, 93, 1)) And Not 2)
  Mid(hfile, 94, 1) = Chr(Asc(Mid(hfile, 94, 1)) And Not 1)
End If
End Sub

Private Sub BY_Click()
If Not BY.Value = 0 Then
  Mid(hfile, 93, 1) = Chr(Asc(Mid(hfile, 93, 1)) Or 4)
  Mid(hfile, 95, 1) = Chr(Asc(Mid(hfile, 95, 1)) Or 1)
Else
  Mid(hfile, 93, 1) = Chr(Asc(Mid(hfile, 93, 1)) And Not 4)
  Mid(hfile, 95, 1) = Chr(Asc(Mid(hfile, 95, 1)) And Not 1)
End If
End Sub

Private Sub MaskEdBox1_Change()
If MaskEdBox1.Text > 255 Then MaskEdBox1.Text = "000"
Mid(hfile, 90, 1) = Chr(MaskEdBox1.Text)
End Sub

Private Sub MaskEdBox2_Change()
If MaskEdBox2.Text > 255 Then MaskEdBox2.Text = "000"
Mid(hfile, 91, 1) = Chr(MaskEdBox2.Text)
End Sub

Private Sub MaskEdBox3_Change()
If MaskEdBox3.Text > 255 Then MaskEdBox3.Text = "000"
Mid(hfile, 92, 1) = Chr(MaskEdBox3.Text)
End Sub

Private Sub Tribe_Click(Index As Integer)
Select Case Index
Case 1
Tribe(1).Picture = TribeImage(1).Picture
Tribe(2).Picture = TribeImage(4).Picture
Tribe(3).Picture = TribeImage(4).Picture
Tribes.Caption = 2
Case 2
Tribe(1).Picture = TribeImage(1).Picture
Tribe(2).Picture = TribeImage(2).Picture
Tribe(3).Picture = TribeImage(4).Picture
Tribes.Caption = 3
Case 3
Tribe(1).Picture = TribeImage(1).Picture
Tribe(2).Picture = TribeImage(2).Picture
Tribe(3).Picture = TribeImage(3).Picture
Tribes.Caption = 4
Case Else
Tribe(1).Picture = TribeImage(4).Picture
Tribe(2).Picture = TribeImage(4).Picture
Tribe(3).Picture = TribeImage(4).Picture
Tribes.Caption = 1
End Select
Mid(hfile, 89, 1) = Chr(Asc(Tribes.Caption) - 48)
End Sub

Private Sub YG_Click()
If Not YG.Value = 0 Then
  Mid(hfile, 95, 1) = Chr(Asc(Mid(hfile, 95, 1)) Or 8)
  Mid(hfile, 96, 1) = Chr(Asc(Mid(hfile, 96, 1)) Or 4)
Else
  Mid(hfile, 95, 1) = Chr(Asc(Mid(hfile, 95, 1)) And Not 8)
  Mid(hfile, 96, 1) = Chr(Asc(Mid(hfile, 96, 1)) And Not 4)
End If
End Sub
Private Sub RG_Click()
If Not RG.Value = 0 Then
  Mid(hfile, 94, 1) = Chr(Asc(Mid(hfile, 94, 1)) Or 8)
  Mid(hfile, 96, 1) = Chr(Asc(Mid(hfile, 96, 1)) Or 2)
Else
  Mid(hfile, 94, 1) = Chr(Asc(Mid(hfile, 94, 1)) And Not 8)
  Mid(hfile, 96, 1) = Chr(Asc(Mid(hfile, 96, 1)) And Not 2)
End If
End Sub

Private Sub RY_Click()
If Not RY.Value = 0 Then
  Mid(hfile, 94, 1) = Chr(Asc(Mid(hfile, 94, 1)) Or 4)
  Mid(hfile, 95, 1) = Chr(Asc(Mid(hfile, 95, 1)) Or 2)
Else
  Mid(hfile, 94, 1) = Chr(Asc(Mid(hfile, 94, 1)) And Not 4)
  Mid(hfile, 95, 1) = Chr(Asc(Mid(hfile, 95, 1)) And Not 2)
End If
End Sub
Private Sub Command1_Click()
If Command1.Caption = "FOG is OFF" Then
  Command1.Caption = "FOG is ON"
  Mid(hfile, 99, 1) = Chr(Asc(Mid(hfile, 99, 1)) Or 1)
Else
  Command1.Caption = "FOG is OFF"
  Mid(hfile, 99, 1) = Chr(Asc(Mid(hfile, 99, 1)) Xor 1)
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "GOD is OFF" Then
  Command2.Caption = "GOD is ON"
  Mid(hfile, 99, 1) = Chr(Asc(Mid(hfile, 99, 1)) Or 2)
Else
  Command2.Caption = "GOD is OFF"
  Mid(hfile, 99, 1) = Chr(Asc(Mid(hfile, 99, 1)) Xor 2)
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
If Asc(Mid(hfile, 98, 1)) < 6 Then
  If Asc(Mid(hfile, 98, 1)) = 0 Then
    Mid(hfile, 98, 1) = Chr(Asc(Mid(hfile, 98, 1)) + 3)
  Else
    Mid(hfile, 98, 1) = Chr(Asc(Mid(hfile, 98, 1)) + 1)
  End If
Else
  Mid(hfile, 98, 1) = Chr(0)
End If
If Asc(Mid(hfile, 98, 1)) <= 6 Then 'Display image if in range
  Tree.Picture = Treepic(Asc(Mid(hfile, 98, 1))).Picture
Else
  Tree.Picture = Treepic(0).Picture
End If
Label3.Caption = Asc(Mid(hfile, 98, 1))
End Sub

Private Sub Command5_Click()
If Asc(Mid(hfile, 98, 1)) > 3 Then
  Mid(hfile, 98, 1) = Chr(Asc(Mid(hfile, 98, 1)) - 1)
  Else
    If Asc(Mid(hfile, 98, 1)) > 1 Then
      Mid(hfile, 98, 1) = Chr(0)
   Else
      Mid(hfile, 98, 1) = Chr(6)
   End If
End If
If Asc(Mid(hfile, 98, 1)) <= 6 Then 'Display image if in range
  Tree.Picture = Treepic(Asc(Mid(hfile, 98, 1))).Picture
Else
  Tree.Picture = Treepic(0).Picture
End If
Label3.Caption = Asc(Mid(hfile, 98, 1))
End Sub

Private Sub Command6_click()
If Asc(Mid(hfile, 97, 1)) < 35 Then
  Mid(hfile, 97, 1) = Chr(Asc(Mid(hfile, 97, 1)) + 1)
  Label1.Caption = Asc(Mid(hfile, 97, 1))
  If Label1.Caption > 9 Then Label1.Caption = Chr(Asc(Mid(hfile, 97, 1)) + 87)
  Landscape.Picture = Image49(Asc(Mid(hfile, 97, 1))).Picture
Else
  Mid(hfile, 97, 1) = Chr(0)
  Label1.Caption = Asc(Mid(hfile, 97, 1))
  If Label1.Caption > 9 Then Label1.Caption = Chr(Asc(Mid(hfile, 97, 1)) + 87)
  Landscape.Picture = Image49(Asc(Mid(hfile, 97, 1))).Picture
End If
End Sub

Private Sub Command7_click()
If Asc(Mid(hfile, 97, 1)) > 0 Then
  Mid(hfile, 97, 1) = Chr(Asc(Mid(hfile, 97, 1)) - 1)
  Label1.Caption = Asc(Mid(hfile, 97, 1))
  If Label1.Caption > 9 Then Label1.Caption = Chr(Asc(Mid(hfile, 97, 1)) + 87)
  If Asc(Mid(hfile, 97, 1)) < 35 Then 'Display image if in range
    Landscape.Picture = Image49(Asc(Mid(hfile, 97, 1))).Picture
  Else
    Landscape.Picture = Image49(0).Picture
  End If
Else
  Mid(hfile, 97, 1) = Chr(35)
  Label1.Caption = Asc(Mid(hfile, 97, 1))
  If Label1.Caption > 9 Then Label1.Caption = Chr(Asc(Mid(hfile, 97, 1)) + 87)
  Landscape.Picture = Image49(Asc(Mid(hfile, 97, 1))).Picture
End If
End Sub

Private Sub exit1_Click()
End
End Sub

Private Sub Form_Load()
hfile = String(616, 0)
Open1.InitDir = "c:\Program Files\Bullfrog\Populous\levels"
End Sub

Private Sub Helpb_Click()
frmHelp.Show
End Sub

Private Sub Help2_Click()
RetVal = Shell("Notepad.exe " + App.Path + "\Spelleditor.txt", 1) ' Run Notepad
End Sub

Private Sub Image1_Click()
If Image1.Tag = 1 Then
  Image1.Picture = Image48.Picture
  Image1.Tag = 2
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Or 32)
Else
  Image1.Picture = Nothing2.Picture
  Image1.Tag = 1
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Xor 32)
End If
End Sub

Private Sub Image10_Click()
If Image10.Tag = 1 Then
  Image10.Picture = Image39.Picture
  Image10.Tag = 2
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Or 8)
Else
  Image10.Picture = Nothing2.Picture
  Image10.Tag = 1
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Xor 8)
End If
End Sub

Private Sub Image11_Click()
If Image11.Tag = 1 Then
  Image11.Picture = Image38.Picture
  Image11.Tag = 2
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Or 32)
Else
  Image11.Picture = Nothing2.Picture
  Image11.Tag = 1
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Xor 32)
End If
End Sub

Private Sub Image12_Click()
If Image12.Tag = 1 Then
  Image12.Picture = Image37.Picture
  Image12.Tag = 2
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Or 64)
Else
  Image12.Picture = Nothing2.Picture
  Image12.Tag = 1
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Xor 64)
End If
End Sub

Private Sub Image13_Click()
If Image13.Tag = 1 Then
  Image13.Picture = Image36.Picture
  Image13.Tag = 2
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Or 128)
Else
  Image13.Picture = Nothing2.Picture
  Image13.Tag = 1
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Xor 128)
End If
End Sub

Private Sub Image14_Click()
If Image14.Tag = 1 Then
  Image14.Picture = Image35.Picture
  Image14.Tag = 2
  Mid(hfile, 3, 1) = Chr(Asc(Mid(hfile, 3, 1)) Or 8)
Else
  Image14.Picture = Nothing2.Picture
  Image14.Tag = 1
  Mid(hfile, 3, 1) = Chr(Asc(Mid(hfile, 3, 1)) Xor 8)
End If
End Sub

Private Sub Image15_Click()
If Image15.Tag = 1 Then
  Image15.Picture = Image32.Picture
  Image15.Tag = 2
  Mid(hfile, 6, 1) = Chr(Asc(Mid(hfile, 6, 1)) Or 128)
Else
  Image15.Picture = Nothing3.Picture
  Image15.Tag = 1
  Mid(hfile, 6, 1) = Chr(Asc(Mid(hfile, 6, 1)) Xor 128)
End If
End Sub

Private Sub Image16_Click()
If Image16.Tag = 1 Then
  Image16.Picture = Image34.Picture
  Image16.Tag = 2
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Or 4)
Else
  Image16.Picture = Nothing2.Picture
  Image16.Tag = 1
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Xor 4)
End If
End Sub

Private Sub Image17_Click()
If Image17.Tag = 1 Then
  Image17.Picture = Image33.Picture
  Image17.Tag = 2
  Mid(hfile, 3, 1) = Chr(Asc(Mid(hfile, 3, 1)) Or 2)
Else
  Image17.Picture = Nothing2.Picture
  Image17.Tag = 1
  Mid(hfile, 3, 1) = Chr(Asc(Mid(hfile, 3, 1)) Xor 2)
End If
End Sub

Private Sub Image18_Click()
If Image18.Tag = 1 Then
  Image18.Picture = Image31.Picture
  Image18.Tag = 2
  Mid(hfile, 6, 1) = Chr(Asc(Mid(hfile, 6, 1)) Or 32)
Else
  Image18.Picture = Nothing3.Picture
  Image18.Tag = 1
  Mid(hfile, 6, 1) = Chr(Asc(Mid(hfile, 6, 1)) Xor 32)
End If
End Sub

Private Sub Image19_Click()
If Image19.Tag = 1 Then
  Image19.Picture = Image30.Picture
  Image19.Tag = 2
  Mid(hfile, 6, 1) = Chr(Asc(Mid(hfile, 6, 1)) Or 1)
Else
  Image19.Picture = Nothing3.Picture
  Image19.Tag = 1
  Mid(hfile, 6, 1) = Chr(Asc(Mid(hfile, 6, 1)) Xor 1)
End If
End Sub

Private Sub Image2_Click()
If Image2.Tag = 1 Then
  Image2.Picture = Image47.Picture
  Image2.Tag = 2
  Mid(hfile, 3, 1) = Chr(Asc(Mid(hfile, 3, 1)) Or 1)
Else
  Image2.Picture = Nothing2.Picture
  Image2.Tag = 1
  Mid(hfile, 3, 1) = Chr(Asc(Mid(hfile, 3, 1)) Xor 1)
End If
End Sub

Private Sub Image20_Click()
If Image20.Tag = 1 Then
  Image20.Picture = Image29.Picture
  Image20.Tag = 2
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Or 64)
Else
  Image20.Picture = Nothing3.Picture
  Image20.Tag = 1
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Xor 64)
End If
End Sub

Private Sub Image21_Click()
If Image21.Tag = 1 Then
  Image21.Picture = Image28.Picture
  Image21.Tag = 2
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Or 32)
Else
  Image21.Picture = Nothing3.Picture
  Image21.Tag = 1
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Xor 32)
End If
End Sub

Private Sub Image22_Click()
If Image22.Tag = 1 Then
  Image22.Picture = Image27.Picture
  Image22.Tag = 2
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Or 128)
Else
  Image22.Picture = Nothing3.Picture
  Image22.Tag = 1
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Xor 128)
End If
End Sub

Private Sub Image23_Click()
If Image23.Tag = 1 Then
  Image23.Picture = Image26.Picture
  Image23.Tag = 2
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Or 16)
Else
  Image23.Picture = Nothing3.Picture
  Image23.Tag = 1
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Xor 16)
End If
End Sub

Private Sub Image24_Click()
If Image24.Tag = 1 Then
  Image24.Picture = Image25.Picture
  Image24.Tag = 2
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Or 2)
Else
  Image24.Picture = Nothing3.Picture
  Image24.Tag = 1
  Mid(hfile, 5, 1) = Chr(Asc(Mid(hfile, 5, 1)) Xor 2)
End If
End Sub

Private Sub Image3_Click()
If Image3.Tag = 1 Then
  Image3.Picture = Image46.Picture
  Image3.Tag = 2
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Or 64)
Else
  Image3.Picture = Nothing2.Picture
  Image3.Tag = 1
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Xor 64)
End If
End Sub

Private Sub Image4_Click()
If Image4.Tag = 1 Then
  Image4.Picture = Image45.Picture
  Image4.Tag = 2
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Or 4)
Else
  Image4.Picture = Nothing2.Picture
  Image4.Tag = 1
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Xor 4)
End If
End Sub

Private Sub Image5_Click()
If Image5.Tag = 1 Then
  Image5.Picture = Image44.Picture
  Image5.Tag = 2
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Or 16)
Else
  Image5.Picture = Nothing2.Picture
  Image5.Tag = 1
  Mid(hfile, 1, 1) = Chr(Asc(Mid(hfile, 1, 1)) Xor 16)
End If
End Sub

Private Sub Image52_Click()
If Image52.Tag = 1 Then
  Image52.Picture = Image53.Picture
  Image52.Tag = 2
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Or 2)
Else
  Image52.Picture = Nothing2.Picture
  Image52.Tag = 1
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Xor 2)
End If
End Sub

Private Sub Image6_Click()
If Image6.Tag = 1 Then
  Image6.Picture = Image43.Picture
  Image6.Tag = 2
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Or 8)
Else
  Image6.Picture = Nothing2.Picture
  Image6.Tag = 1
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Xor 8)
End If
End Sub

Private Sub Image7_Click()
If Image7.Tag = 1 Then
  Image7.Picture = Image42.Picture
  Image7.Tag = 2
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Or 1)
Else
  Image7.Picture = Nothing2.Picture
  Image7.Tag = 1
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Xor 1)
End If
End Sub

Private Sub Image8_Click()
If Image8.Tag = 1 Then
  Image8.Picture = Image41.Picture
  Image8.Tag = 2
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Or 128)
Else
  Image8.Picture = Nothing2.Picture
  Image8.Tag = 1
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Xor 128)
End If
End Sub

Private Sub Image9_Click()
If Image9.Tag = 1 Then
  Image9.Picture = Image40.Picture
  Image9.Tag = 2
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Or 16)
Else
  Image9.Picture = Nothing2.Picture
  Image9.Tag = 1
  Mid(hfile, 2, 1) = Chr(Asc(Mid(hfile, 2, 1)) Xor 16)
End If
End Sub


Private Sub load1_Click()
On Error GoTo endload
Open1.Filter = "Populous hdr Files (LEVL*.HDR)|LEVL*.HDR"
Open1.ShowOpen
Open Open1.FileName For Binary As #1
Main.Caption = Open1.FileTitle
hfile = String(616, 0)
Get #1, , hfile
Close #1
On Error GoTo 0
If Asc(Mid(hfile, 99, 1)) And 1 Then 'Fog
  Command1.Caption = "FOG is ON"
Else
  Command1.Caption = "FOG is OFF"
End If
If Asc(Mid(hfile, 99, 1)) And 2 Then 'God mode
  Command2.Caption = "GOD is ON"
Else
  Command2.Caption = "GOD is OFF"
End If
If Asc(Mid(hfile, 1, 1)) And 4 Then 'Blast
  Image16.Tag = 2
  Image16.Picture = Image34.Picture
Else
  Image16.Tag = 1
  Image16.Picture = Nothing2.Picture
End If

If Asc(Mid(hfile, 1, 1)) And 8 Then 'Lightning
  Image10.Tag = 2
  Image10.Picture = Image39.Picture
Else
  Image10.Tag = 1
  Image10.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 1, 1)) And 16 Then 'Tornado
  Image5.Tag = 2
  Image5.Picture = Image44.Picture
Else
  Image5.Tag = 1
  Image5.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 1, 1)) And 32 Then 'Swarm
  Image11.Tag = 2
  Image11.Picture = Image38.Picture
Else
  Image11.Tag = 1
  Image11.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 1, 1)) And 64 Then 'Invisibility
  Image12.Tag = 2
  Image12.Picture = Image37.Picture
Else
  Image12.Tag = 1
  Image12.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 1, 1)) And 128 Then 'Hypnotise
  Image13.Tag = 2
  Image13.Picture = Image36.Picture
Else
  Image13.Tag = 1
  Image13.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 2, 1)) And 1 Then 'Firestorm
  Image7.Tag = 2
  Image7.Picture = Image42.Picture
Else
  Image7.Tag = 1
  Image7.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 2, 1)) And 2 Then 'Ghost Army
  Image52.Tag = 2
  Image52.Picture = Image53.Picture
Else
  Image52.Tag = 1
  Image52.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 2, 1)) And 4 Then 'Erode
  Image4.Tag = 2
  Image4.Picture = Image45.Picture
Else
  Image4.Tag = 1
  Image4.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 2, 1)) And 8 Then 'Swamp
  Image6.Tag = 2
  Image6.Picture = Image43.Picture
Else
  Image6.Tag = 1
  Image6.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 2, 1)) And 16 Then 'Landbridge
  Image9.Tag = 2
  Image9.Picture = Image40.Picture
Else
  Image9.Tag = 1
  Image9.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 2, 1)) And 32 Then 'Angel of Death
  Image1.Tag = 2
  Image1.Picture = Image48.Picture
Else
  Image1.Tag = 1
  Image1.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 2, 1)) And 64 Then 'Earthquake
  Image3.Tag = 2
  Image3.Picture = Image46.Picture
Else
  Image3.Tag = 1
  Image3.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 2, 1)) And 128 Then 'Flatten
  Image8.Tag = 2
  Image8.Picture = Image41.Picture
Else
  Image8.Tag = 1
  Image8.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 3, 1)) And 1 Then 'Volcano
  Image2.Tag = 2
  Image2.Picture = Image47.Picture
Else
  Image2.Tag = 1
  Image2.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 3, 1)) And 2 Then 'Convert
  Image17.Tag = 2
  Image17.Picture = Image33.Picture
Else
  Image17.Tag = 1
  Image17.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 3, 1)) And 8 Then 'Magic Shield
  Image14.Tag = 2
  Image14.Picture = Image35.Picture
Else
  Image14.Tag = 1
  Image14.Picture = Nothing2.Picture
End If
If Asc(Mid(hfile, 5, 1)) And 2 Then 'Hut
  Image24.Tag = 2
  Image24.Picture = Image25.Picture
Else
  Image24.Tag = 1
  Image24.Picture = Nothing3.Picture
End If
If Asc(Mid(hfile, 5, 1)) And 16 Then 'Guard Tower
  Image23.Tag = 2
  Image23.Picture = Image26.Picture
Else
  Image23.Tag = 1
  Image23.Picture = Nothing3.Picture
End If
If Asc(Mid(hfile, 5, 1)) And 32 Then 'Temple
  Image21.Tag = 2
  Image21.Picture = Image28.Picture
Else
  Image21.Tag = 1
  Image21.Picture = Nothing3.Picture
End If
If Asc(Mid(hfile, 5, 1)) And 64 Then 'Spy Hut
  Image20.Tag = 2
  Image20.Picture = Image29.Picture
Else
  Image20.Tag = 1
  Image20.Picture = Nothing3.Picture
End If
If Asc(Mid(hfile, 5, 1)) And 128 Then 'Warrior Hut
  Image22.Tag = 2
  Image22.Picture = Image27.Picture
Else
  Image22.Tag = 1
  Image22.Picture = Nothing3.Picture
End If
If Asc(Mid(hfile, 6, 1)) And 1 Then 'Fire Warrior Hut
  Image19.Tag = 2
  Image19.Picture = Image30.Picture
Else
  Image19.Tag = 1
  Image19.Picture = Nothing3.Picture
End If
If Asc(Mid(hfile, 6, 1)) And 32 Then 'Boat Hut
  Image18.Tag = 2
  Image18.Picture = Image31.Picture
Else
  Image18.Tag = 1
  Image18.Picture = Nothing3.Picture
End If
If Asc(Mid(hfile, 6, 1)) And 128 Then 'Balloon Hut
  Image15.Tag = 2
  Image15.Picture = Image32.Picture
Else
  Image15.Tag = 1
  Image15.Picture = Nothing3.Picture
End If
Nothing1.Picture = Nothing2.Picture 'Clear guest spells
Nothing1.ToolTipText = "Guest Spells (Cannot be saved)"
Label1.Caption = Asc(Mid(hfile, 97, 1)) 'Landscape
If Label1.Caption > 9 Then Label1.Caption = Chr(Asc(Mid(hfile, 97, 1)) + 87)
If Asc(Mid(hfile, 97, 1)) < 35 Then
  Landscape.Picture = Image49(Asc(Mid(hfile, 97, 1))).Picture
Else
  Landscape.Picture = Image49(0).Picture
End If
'Tree Style
If Asc(Mid(hfile, 98, 1)) < 6 Then
  Tree.Picture = Treepic(Asc(Mid(hfile, 98, 1))).Picture 'Tree Style
Else
  Tree.Picture = Treepic(0).Picture
End If
Label3.Caption = Asc(Mid(hfile, 98, 1))
'Allies
If Asc(Mid(hfile, 93, 1)) And 2 Then
  BR.Value = 1
Else
  BR.Value = 0
End If
If Asc(Mid(hfile, 93, 1)) And 4 Then
  BY.Value = 1
Else
  BY.Value = 0
End If
If Asc(Mid(hfile, 93, 1)) And 8 Then
  BG.Value = 1
Else
  BG.Value = 0
End If
If Asc(Mid(hfile, 94, 1)) And 4 Then
  RY.Value = 1
Else
  RY.Value = 0
End If
If Asc(Mid(hfile, 94, 1)) And 8 Then
  RG.Value = 1
Else
  RG.Value = 0
End If
If Asc(Mid(hfile, 95, 1)) And 8 Then
  YG.Value = 1
Else
  YG.Value = 0
End If
'Script files
MaskEdBox1.Text = Format(Asc(Mid(hfile, 90, 1)), "000")
MaskEdBox2.Text = Format(Asc(Mid(hfile, 91, 1)), "000")
MaskEdBox3.Text = Format(Asc(Mid(hfile, 92, 1)), "000")
'Number of Tribes
Tribes.Caption = Chr(Asc(Mid(hfile, 89, 1)) + 48)
Call Tribe_Click(Tribes.Caption - 1)
Save1.Enabled = True
On Error GoTo 0
Exit Sub
endload:
On Error GoTo 0
Msgval = MsgBox("Error loading " & Open1.FileName, 48, "Populous Spell Editor")
End Sub

Private Sub Nothing1_Click()
If Nothing1.ToolTipText = "Guest Spells (Cannot be saved)" Then
  Nothing1.Picture = Bloodlust.Picture
  Nothing1.ToolTipText = "Bloodlust (Cannot be saved)"
  Else
  If Nothing1.ToolTipText = "Bloodlust (Cannot be saved)" Then
    Nothing1.Picture = Teleport.Picture
    Nothing1.ToolTipText = "Teleport (Cannot be saved)"
    Else
    If Nothing1.ToolTipText = "Teleport (Cannot be saved)" Then
      Nothing1.Picture = Armageddon.Picture
      Nothing1.ToolTipText = "Armageddon (Cannot be saved)"
      Else
      If Nothing1.ToolTipText = "Armageddon (Cannot be saved)" Then
        Nothing1.Picture = Nothing2.Picture
        Nothing1.ToolTipText = "Guest Spells (Cannot be saved)"
      End If
    End If
  End If
End If
End Sub


Private Sub save1_Click()
Dim Msgval
On Error GoTo endsave
Open Open1.FileName For Binary As #1
Put #1, 1, hfile
Close #1
Msgval = MsgBox(Open1.FileName & " saved", 64, "Populous Spell Editor")
On Error GoTo 0
Exit Sub
endsave:
On Error GoTo 0
Close #1
Msgval = MsgBox("Error saving " & Open1.FileName & ". Check that file properties are not set to read-only.", 48, "Populous Spell Editor")
End Sub
Private Sub Errors()

End Sub


