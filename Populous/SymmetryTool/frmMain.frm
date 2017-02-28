VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Populous Symmetry Tool"
   ClientHeight    =   10050
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   12075
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   670
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRotate 
      Caption         =   "Rotate 90° Clockwise"
      Height          =   495
      Left            =   3960
      TabIndex        =   44
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Move map"
      Height          =   2415
      Left            =   4200
      TabIndex        =   35
      Top             =   600
      Width           =   1695
      Begin VB.CommandButton MoveDown 
         Caption         =   "Down"
         Height          =   495
         Left            =   480
         TabIndex        =   40
         Top             =   1280
         Width           =   735
      End
      Begin VB.CommandButton MoveRight 
         Caption         =   "Right"
         Height          =   495
         Left            =   860
         TabIndex        =   43
         Top             =   750
         Width           =   735
      End
      Begin VB.TextBox txtRotate 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   540
         TabIndex        =   37
         Top             =   1800
         Width           =   615
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   540
         Max             =   6
         TabIndex        =   36
         Tag             =   "64"
         Top             =   2040
         Value           =   4
         Width           =   615
      End
      Begin VB.CommandButton MoveUp 
         Caption         =   "Up"
         Height          =   495
         Left            =   480
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton MoveLeft 
         Caption         =   "Left"
         Height          =   495
         Left            =   100
         TabIndex        =   41
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Map"
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Top             =   1845
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "By"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1845
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdSymmetry 
      Caption         =   "Mirror"
      Height          =   495
      Left            =   2160
      TabIndex        =   29
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame Frame7 
      Caption         =   "Replicate blue objects as"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      Begin VB.OptionButton optBlue 
         Caption         =   "B"
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "R"
         Height          =   200
         Index           =   1
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "Y"
         Height          =   200
         Index           =   2
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "G"
         Height          =   200
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Replicate red objects as"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1230
      Width           =   2055
      Begin VB.OptionButton optRed 
         Caption         =   "B"
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optRed 
         Caption         =   "R"
         Height          =   200
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optRed 
         Caption         =   "Y"
         Height          =   200
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton optRed 
         Caption         =   "G"
         Height          =   200
         Index           =   3
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Replicate yellow objects as"
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   1870
      Width           =   2055
      Begin VB.OptionButton optYellow 
         Caption         =   "B"
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optYellow 
         Caption         =   "R"
         Height          =   200
         Index           =   1
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optYellow 
         Caption         =   "Y"
         Height          =   200
         Index           =   2
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optYellow 
         Caption         =   "G"
         Height          =   200
         Index           =   3
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Replicate green objects as"
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
      Begin VB.OptionButton optGreen 
         Caption         =   "G"
         Height          =   200
         Index           =   3
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optGreen 
         Caption         =   "Y"
         Height          =   200
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optGreen 
         Caption         =   "R"
         Height          =   200
         Index           =   1
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optGreen 
         Caption         =   "B"
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   5760
      Left            =   6120
      ScaleHeight     =   384
      ScaleMode       =   0  'User
      ScaleWidth      =   384
      TabIndex        =   32
      Top             =   3960
      Width           =   5760
   End
   Begin VB.CheckBox chkMirrorLand 
      Caption         =   "Mirror / Rotate Land"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkMirrorObjects 
      Caption         =   "Mirror / Rotate Objects"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Symmetry"
      Height          =   2415
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   1815
      Begin VB.OptionButton optAxis 
         Caption         =   "X Both Diagonals"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "/ Diagonal Axis"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "\ Diagonal Axis"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "+ Both Axes"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "- Horizontal Axis"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "| Vertical Axis"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   13440
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   13440
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CCA
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DDC
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   5760
      Left            =   120
      ScaleHeight     =   384
      ScaleMode       =   0  'User
      ScaleWidth      =   384
      TabIndex        =   31
      Top             =   3960
      Width           =   5760
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "North Pole View"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2865
      Left            =   6120
      Picture         =   "frmMain.frx":1EEE
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5850
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "South Pole View"
      Height          =   255
      Left            =   6120
      TabIndex        =   33
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      X1              =   0
      X2              =   808
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSavepic 
         Caption         =   "Sa&ve Image"
         Begin VB.Menu mnuNorth 
            Caption         =   "North Pole"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuSouth 
            Caption         =   "South Pole"
         End
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "O&ptions"
      Begin VB.Menu mnuOptionsShowLines 
         Caption         =   "Sho&w Symmetry Lines"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsShowObjects 
         Caption         =   "Show O&bjects"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsShowSouthPole 
         Caption         =   "Show South Pole &View"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShading 
         Caption         =   "Hi&ghlight Source"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpFile 
         Caption         =   "He&lp"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public map As String
Private mapsave(0 To 9) As String
Public CheckpointNumber As Integer
Public CheckpointCount As Integer
Public UndoCount As Integer
Public LastCheckpointType As String
Public sFile As String, iFile As String
Public objstart As Long
Public scrollinc As Integer
Public SymmetryType As Integer
Private Sub cmdRotate_Click()
'Rotate 90 degrees clockwise
Dim i As Long, j As Long, k As Long, X As Long, Y As Long, obj As String, obyte As Integer
Call checkpoint("Rotate") 'take checkpoint for undo
LastCheckpointNumber = (CheckpointNumber + 9) Mod 10
If chkMirrorLand.Value = 1 Then 'Rotate Land?
  'copy top LHS to top RHS
  For Y = 0 To 127
    For X = 0 To 127
     Mid(map, (Y * 256) + (X * 2) + 1, 2) = Mid(mapsave(LastCheckpointNumber), (X * 256) + ((127 - Y) * 2) + 1, 2)
    Next
  Next
End If
If chkMirrorObjects.Value = 1 Then 'Rotate Objects?
  For i = 0 To 110000 Step 55
    If Mid(map, objstart + i, 2) <> zeroword Then
      'process each object
      obj = Mid(map, objstart + i + 6, 1)
      Mid(map, objstart + i + 6, 1) = Chr(255 - Asc(Mid(map, objstart + i + 4, 1))) 'vertical position of object
      Mid(map, objstart + i + 4, 1) = obj 'horizontal position of object
      'Change orientation for 02-Buildings and 05-Scenery
      If Asc(Mid(map, objstart + i + 1, 1)) = 2 Or Asc(Mid(map, objstart + i + 1, 1)) = 5 Then
        If Asc(Mid(map, objstart + i + 1, 1)) = 2 Then
          obyte = 8 'Orientation is in byte 9 for object group 02
        Else
          obyte = 11 'Orientation is in byte 12 for object group 05
        End If
        Orientation = Asc(Mid(map, objstart + i + obyte, 1))
        Select Case Orientation
        Case 0 'S
          Mid(map, objstart + i + obyte, 1) = Chr(2) 'W
        Case 1 'SW
          Mid(map, objstart + i + obyte, 1) = Chr(3) 'NW
        Case 2 'W
          Mid(map, objstart + i + obyte, 1) = Chr(4) 'N
        Case 3 'NW
          Mid(map, objstart + i + obyte, 1) = Chr(5) 'NE
        Case 4 'N
          Mid(map, objstart + i + obyte, 1) = Chr(6) 'E
        Case 5 'NE
          Mid(map, objstart + i + obyte, 1) = Chr(7) 'SE
        Case 6 'E
          Mid(map, objstart + i + obyte, 1) = Chr(0) 'S
        Case 7 'SE
          Mid(map, objstart + i + obyte, 1) = Chr(1) 'SW
        Case Else
        End Select
      End If
    End If
  Next
End If
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub cmdSymmetry_Click()
Dim i As Long, j As Long, k As Long, X As Long, Y As Long
Call checkpoint("Symmetry") 'take checkpoint for undo
If chkMirrorLand.Value = 1 Then
If optAxis(0).Value = True Then
  'Vertical Axis symmetry - set RHS of map to mirror of LHS
  For j = 0 To 32512 Step 256
    For i = 1 To 127 Step 2
     Mid(map, j + 256 - i, 2) = Mid(map, j + i, 2)
    Next
  Next
Else
  If optAxis(1).Value = True Then
  'Horizontal Axis symmetry - set top of map to mirror of bottom
  For j = 0 To 16128 Step 256
    For i = 1 To 255 Step 2
      Mid(map, 32512 - j + i, 2) = Mid(map, j + i, 2)
    Next
  Next
  Else
    If optAxis(2).Value = True Then
      'Vertical & Horizontal Axis symmetry - set top RHS of map to mirror of top LHS
      For j = 0 To 16128 Step 256
        For i = 1 To 127 Step 2
         Mid(map, j + 256 - i, 2) = Mid(map, j + i, 2)
        Next
      Next
      'set top of map to mirror of bottom
      For j = 0 To 16128 Step 256
        For i = 1 To 255 Step 2
          Mid(map, 32512 - j + i, 2) = Mid(map, j + i, 2)
        Next
      Next
    Else
      If optAxis(3).Value = True Then
        'Diagonal TL-BR Axis - set top right triangle of map to mirror of bottom left
         For Y = 0 To 127
          For X = 0 To 127
           If X + Y < 128 Then Mid(map, ((127 - X) * 256) + ((127 - Y) * 2) + 1, 2) = Mid(map, (X * 2) + (Y * 256) + 1, 2)
          Next
        Next
      Else
        If optAxis(4).Value = True Then
        'Diagonal BL-TR Axis - set top right triangle of map to mirror of bottom left
        For Y = 0 To 127
          For X = 0 To Y
            Mid(map, (X * 256) + (Y * 2) + 1, 2) = Mid(map, (X * 2) + (Y * 256) + 1, 2)
          Next
        Next
        Else
        'Diagonal Both Axes - 4 way diagonal symmetry through centre point
        'Diagonal TL-BR Axis - set top right triangle of map to mirror of bottom left
         For Y = 0 To 127
          For X = 0 To 127
           If X + Y < 128 Then Mid(map, ((127 - X) * 256) + ((127 - Y) * 2) + 1, 2) = Mid(map, (X * 2) + (Y * 256) + 1, 2)
          Next
         Next
        'Diagonal BL-TR Axis - set top right triangle of map to mirror of bottom left
         For Y = 0 To 127
          For X = 0 To Y
            Mid(map, (X * 256) + (Y * 2) + 1, 2) = Mid(map, (X * 2) + (Y * 256) + 1, 2)
          Next
         Next
        End If
      End If
    End If
  End If
End If
End If
If chkMirrorObjects.Value = 1 Then Call MirrorObjects
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub
Sub MirrorObjects()
Dim zeroword As String, i As Long, j As Long, k As Long, jmax As Long, objbyte As Integer, Orientation As Integer
Dim mapobj(2000) As String
Dim mapobjindx(110000) As Integer
zeroword = Chr$(0) + Chr$(0)
mapobjindx(0) = 0
  
  If optAxis(0) = True Then
  'Two way mirror (Vertical)
  '--------------------------
  j = 1
  k = 1
  jmax = 0
  'Keep objects in copy sector
  For i = 0 To 110000 Step 55
    If Mid(map, objstart + i, 2) <> zeroword And Asc(Mid(map, objstart + i + 4, 1)) < 128 Then
      mapobj(j) = Mid(map, objstart + i, 55)
      mapobjindx(k) = j 'Keep new record number for each object stored - for trigger adjustments
      jmax = j
      j = j + 1
    End If
    k = k + 1
  Next
  'Adjust triggers
  For j = 1 To jmax
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(j), 14 + k, 2) = WordChr(mapobjindx(ChrWord(Mid(mapobj(j), 14 + k, 2))))
        End If
      Next
    End If
  Next
  'Replicate objects in each sector
  For j = 1 To jmax
    mapobj(jmax + j) = mapobj(j)
    Mid(mapobj(jmax + j), 5, 1) = Chr(255 - Asc(Mid(mapobj(j), 5, 1)))
    'Change orientation for 02-Buildings and 05-Scenery
    If Asc(Mid(mapobj(j), 2, 1)) = 2 Or Asc(Mid(mapobj(j), 2, 1)) = 5 Then
      If Asc(Mid(mapobj(j), 2, 1)) = 2 Then
        obyte = 9 'Orientation is in byte 9 for object group 02
      Else
        obyte = 12 'Orientation is in byte 12 for object group 05
      End If
      Orientation = Asc(Mid(mapobj(jmax + j), obyte, 1))
      Select Case Orientation
      Case 2 'W=E
        Mid(mapobj(jmax + j), obyte, 1) = Chr(6)
      Case 6 'E=W
        Mid(mapobj(jmax + j), obyte, 1) = Chr(2)
      Case 1 'SW=SE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(7)
      Case 7 'SE=SW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(1)
      Case 3 'NW=NE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(5)
      Case 5 'NE=NW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(3)
      Case Else
      End Select
    End If
    'Adjust triggers
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(jmax + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax + j), 14 + k, 2)) + jmax)
        End If
      Next
    End If
    'Convert Blue tribe objects to Red, Yellow, Green
    objbyte = Asc(Mid(mapobj(j), 3, 1))
    If objbyte <> 255 And Mid(mapobj(j), 1, 2) <> Chr(6) + Chr(6) Then 'Ignore Triggers (0606) and Wild colours
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax + j), 3, 1) = Chr(objbyte)
    End If
  Next
  jmax = jmax * 2
  'Remove all objects from map
  For i = objstart To 110000 + objstart Step 2
    Mid(map, i, 2) = zeroword
  Next
  'Add all objects to map
  i = objstart
  For j = 1 To jmax
    Mid(map, i, 55) = mapobj(j)
    i = i + 55
  Next
  End If
  If optAxis(1) = True Then
  'Two way mirror - horizontal
  '----------------------------
  j = 1
  k = 1
  jmax = 0
  'Keep objects in copy sector
  For i = 0 To 110000 Step 55
    If Mid(map, objstart + i, 2) <> zeroword And Asc(Mid(map, objstart + i + 6, 1)) < 128 Then
      mapobj(j) = Mid(map, objstart + i, 55)
      mapobjindx(k) = j 'Keep new record number for each object stored - for trigger adjustments
      jmax = j
      j = j + 1
    End If
    k = k + 1
  Next
  'Adjust triggers
  For j = 1 To jmax
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(j), 14 + k, 2) = WordChr(mapobjindx(ChrWord(Mid(mapobj(j), 14 + k, 2))))
        End If
      Next
    End If
  Next
  'Replicate objects in each sector
  For j = 1 To jmax
    mapobj(jmax + j) = mapobj(j)
    Mid(mapobj(jmax + j), 7, 1) = Chr(255 - Asc(Mid(mapobj(j), 7, 1)))
    'Change orientation for 02-Buildings and 05-Scenery
    If Asc(Mid(mapobj(j), 2, 1)) = 2 Or Asc(Mid(mapobj(j), 2, 1)) = 5 Then
      If Asc(Mid(mapobj(j), 2, 1)) = 2 Then
        obyte = 9 'Orientation is in byte 9 for object group 02
      Else
        obyte = 12 'Orientation is in byte 12 for object group 05
      End If
      Orientation = Asc(Mid(mapobj(jmax + j), obyte, 1))
      Select Case Orientation
      Case 0 'S=N
        Mid(mapobj(jmax + j), obyte, 1) = Chr(4)
      Case 4 'N=S
        Mid(mapobj(jmax + j), obyte, 1) = Chr(0)
      Case 1 'SW=NW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(3)
      Case 3 'NW=SW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(1)
      Case 5 'NE=SE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(7)
      Case 7 'SE=NE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(5)
      Case Else
      End Select
    End If
    'Adjust triggers
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(jmax + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax + j), 14 + k, 2)) + jmax)
        End If
      Next
    End If
    'Convert Blue tribe objects to Red, Yellow, Green
    objbyte = Asc(Mid(mapobj(j), 3, 1))
    If objbyte <> 255 And Mid(mapobj(j), 1, 2) <> Chr(6) + Chr(6) Then 'Ignore Triggers (0606) and Wild colours
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax + j), 3, 1) = Chr(objbyte)
    End If
  Next
  jmax = jmax * 2
  'Remove all objects from map
  For i = objstart To 110000 + objstart Step 2
    Mid(map, i, 2) = zeroword
  Next
  'Add all objects to map
  i = objstart
  For j = 1 To jmax
    Mid(map, i, 55) = mapobj(j)
    i = i + 55
  Next
  End If
  If optAxis(2) = True Then
  'Four way mirror
  '--------------------------
  j = 1
  k = 1
  jmax = 0
  'Keep objects in copy sector
  For i = 0 To 110000 Step 55
    If Mid(map, objstart + i, 2) <> zeroword And Asc(Mid(map, objstart + i + 4, 1)) < 128 And Asc(Mid(map, objstart + i + 6, 1)) < 128 Then
      mapobj(j) = Mid(map, objstart + i, 55)
      mapobjindx(k) = j 'Keep new record number for each object stored - for trigger adjustments
      jmax = j
      j = j + 1
    End If
    k = k + 1
  Next
  'Adjust triggers
  For j = 1 To jmax
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(j), 14 + k, 2) = WordChr(mapobjindx(ChrWord(Mid(mapobj(j), 14 + k, 2))))
        End If
      Next
    End If
  Next
  'Replicate objects in each sector
  For j = 1 To jmax
    mapobj(jmax + j) = mapobj(j)
    Mid(mapobj(jmax + j), 7, 1) = Chr(255 - Asc(Mid(mapobj(j), 7, 1)))
    mapobj(jmax * 2 + j) = mapobj(j)
    Mid(mapobj(jmax * 2 + j), 5, 1) = Chr(255 - Asc(Mid(mapobj(j), 5, 1)))
    Mid(mapobj(jmax * 2 + j), 7, 1) = Chr(255 - Asc(Mid(mapobj(j), 7, 1)))
    mapobj(jmax * 3 + j) = mapobj(j)
    Mid(mapobj(jmax * 3 + j), 5, 1) = Chr(255 - Asc(Mid(mapobj(j), 5, 1)))
    'Change orientation for 02-Buildings and 05-Scenery
    If Asc(Mid(mapobj(j), 2, 1)) = 2 Or Asc(Mid(mapobj(j), 2, 1)) = 5 Then
      If Asc(Mid(mapobj(j), 2, 1)) = 2 Then
        obyte = 9 'Orientation is in byte 9 for object group 02
      Else
        obyte = 12 'Orientation is in byte 12 for object group 05
      End If
      Orientation = Asc(Mid(mapobj(jmax + j), obyte, 1))
      Select Case Orientation
      Case 0 'S
        Mid(mapobj(jmax + j), obyte, 1) = Chr(4) 'N
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(4) 'N
      Case 1 'SW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(3) 'NW
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(5) 'NE
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(7) 'SE
      Case 2 'W
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(6) 'E
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(6) 'E
      Case 3 'NW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(1) 'SW
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(7) 'SE
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(5) 'NE
      Case 4 'N
        Mid(mapobj(jmax + j), obyte, 1) = Chr(0) 'S
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(0) 'S
      Case 5 'NE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(7) 'SE
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(1) 'SW
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(3) 'NW
      Case 6 'E
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(2) 'W
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(2) 'W
      Case 7 'SE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(5) 'NE
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(3) 'NW
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(1) 'SW
      Case Else
      End Select
    End If
    'Adjust triggers
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(jmax + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax + j), 14 + k, 2)) + jmax)
          Mid(mapobj(jmax * 2 + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax * 2 + j), 14 + k, 2)) + jmax * 2)
          Mid(mapobj(jmax * 3 + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax * 3 + j), 14 + k, 2)) + jmax * 3)
        End If
      Next
    End If
    'Convert Blue tribe objects to Red, Yellow, Green
    objbyte = Asc(Mid(mapobj(j), 3, 1))
    If objbyte <> 255 And Mid(mapobj(j), 1, 2) <> Chr(6) + Chr(6) Then 'Ignore Triggers (0606) and Wild colours
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax + j), 3, 1) = Chr(objbyte)
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax * 2 + j), 3, 1) = Chr(objbyte)
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax * 3 + j), 3, 1) = Chr(objbyte)
    End If
  Next
  jmax = jmax * 4
  'Remove all objects from map
  For i = objstart To 110000 + objstart Step 2
    Mid(map, i, 2) = zeroword
  Next
  'Add all objects to map
  i = objstart
  For j = 1 To jmax
    Mid(map, i, 55) = mapobj(j)
    i = i + 55
  Next
  End If
  If optAxis(3) = True Then
  'Two way mirror (Diagonal TL-BR)
  '-------------------------------
  j = 1
  k = 1
  jmax = 0
  'Keep objects in copy sector
  For i = 0 To 110000 Step 55
    If Mid(map, objstart + i, 2) <> zeroword And Asc(Mid(map, objstart + i + 4, 1)) + Asc(Mid(map, objstart + i + 6, 1)) < 256 Then
      mapobj(j) = Mid(map, objstart + i, 55)
      mapobjindx(k) = j 'Keep new record number for each object stored - for trigger adjustments
      jmax = j
      j = j + 1
    End If
    k = k + 1
  Next
   'Adjust triggers
  For j = 1 To jmax
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(j), 14 + k, 2) = WordChr(mapobjindx(ChrWord(Mid(mapobj(j), 14 + k, 2))))
        End If
      Next
    End If
  Next
  'Replicate objects in each sector
  For j = 1 To jmax
    mapobj(jmax + j) = mapobj(j)
    Mid(mapobj(jmax + j), 5, 1) = Chr(255 - Asc(Mid(mapobj(j), 7, 1)))
    Mid(mapobj(jmax + j), 7, 1) = Chr(255 - Asc(Mid(mapobj(j), 5, 1)))
    'Change orientation for 02-Buildings and 05-Scenery
    If Asc(Mid(mapobj(j), 2, 1)) = 2 Or Asc(Mid(mapobj(j), 2, 1)) = 5 Then
      If Asc(Mid(mapobj(j), 2, 1)) = 2 Then
        obyte = 9 'Orientation is in byte 9 for object group 02
      Else
        obyte = 12 'Orientation is in byte 12 for object group 05
      End If
      Orientation = Asc(Mid(mapobj(jmax + j), obyte, 1))
      Select Case Orientation
      Case 0 'S
        Mid(mapobj(jmax + j), obyte, 1) = Chr(6) 'E
      Case 1 'SW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(5) 'NE
      Case 2 'W
        Mid(mapobj(jmax + j), obyte, 1) = Chr(4) 'N
      Case 3 'NW
      Case 4 'N
        Mid(mapobj(jmax + j), obyte, 1) = Chr(2) 'W
      Case 5 'NE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(1) 'SW
      Case 6 'E
        Mid(mapobj(jmax + j), obyte, 1) = Chr(0) 'S
      Case 7 'SE
      Case Else
      End Select
    End If
    'Adjust triggers
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(jmax + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax + j), 14 + k, 2)) + jmax)
        End If
      Next
    End If
    'Convert Blue tribe objects to Red, Yellow, Green
    objbyte = Asc(Mid(mapobj(j), 3, 1))
    If objbyte <> 255 And Mid(mapobj(j), 1, 2) <> Chr(6) + Chr(6) Then 'Ignore Triggers (0606) and Wild colours
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax + j), 3, 1) = Chr(objbyte)
    End If
  Next
  jmax = jmax * 2
  'Remove all objects from map
  For i = objstart To 110000 + objstart Step 2
    Mid(map, i, 2) = zeroword
  Next
  'Add all objects to map
  i = objstart
  For j = 1 To jmax
    Mid(map, i, 55) = mapobj(j)
    i = i + 55
  Next
  End If
  If optAxis(4) = True Then
  'Two way mirror (Diagonal BL-TR)
  '-------------------------------
  j = 1
  k = 1
  jmax = 0
  'Keep objects in copy sector
  For i = 0 To 110000 Step 55
    If Mid(map, objstart + i, 2) <> zeroword And Asc(Mid(map, objstart + i + 4, 1)) < Asc(Mid(map, objstart + i + 6, 1)) Then
      mapobj(j) = Mid(map, objstart + i, 55)
      mapobjindx(k) = j 'Keep new record number for each object stored - for trigger adjustments
      jmax = j
      j = j + 1
    End If
    k = k + 1
  Next
  'Adjust triggers
  For j = 1 To jmax
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(j), 14 + k, 2) = WordChr(mapobjindx(ChrWord(Mid(mapobj(j), 14 + k, 2))))
        End If
      Next
    End If
  Next
  'Replicate objects in each sector
  For j = 1 To jmax
    mapobj(jmax + j) = mapobj(j)
    Mid(mapobj(jmax + j), 5, 1) = Chr(Asc(Mid(mapobj(j), 7, 1)))
    Mid(mapobj(jmax + j), 7, 1) = Chr(Asc(Mid(mapobj(j), 5, 1)))
    'Change orientation for 02-Buildings and 05-Scenery
    If Asc(Mid(mapobj(j), 2, 1)) = 2 Or Asc(Mid(mapobj(j), 2, 1)) = 5 Then
      If Asc(Mid(mapobj(j), 2, 1)) = 2 Then
        obyte = 9 'Orientation is in byte 9 for object group 02
      Else
        obyte = 12 'Orientation is in byte 12 for object group 05
      End If
      Orientation = Asc(Mid(mapobj(jmax + j), obyte, 1))
      Select Case Orientation
      Case 0 'S
        Mid(mapobj(jmax + j), obyte, 1) = Chr(2) 'W
      Case 1 'SW
      Case 2 'W
        Mid(mapobj(jmax + j), obyte, 1) = Chr(0) 'S
      Case 3 'NW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(7) 'SE
      Case 4 'N
        Mid(mapobj(jmax + j), obyte, 1) = Chr(6) 'E
      Case 5 'NE
      Case 6 'E
        Mid(mapobj(jmax + j), obyte, 1) = Chr(4) 'N
      Case 7 'SE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(3) 'NW
      Case Else
      End Select
    End If
    'Adjust triggers
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(jmax + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax + j), 14 + k, 2)) + jmax)
        End If
      Next
    End If
    'Convert Blue tribe objects to Red, Yellow, Green
    objbyte = Asc(Mid(mapobj(j), 3, 1))
    If objbyte <> 255 And Mid(mapobj(j), 1, 2) <> Chr(6) + Chr(6) Then 'Ignore Triggers (0606) and Wild colours
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax + j), 3, 1) = Chr(objbyte)
    End If
  Next
  jmax = jmax * 2
  'Remove all objects from map
  For i = objstart To 110000 + objstart Step 2
    Mid(map, i, 2) = zeroword
  Next
  'Add all objects to map
  i = objstart
  For j = 1 To jmax
    Mid(map, i, 55) = mapobj(j)
    i = i + 55
  Next
  End If
  If optAxis(5) = True Then
  'Four way mirror (both diagonals)
  '-------------------------------
  j = 1
  k = 1
  jmax = 0
  'Keep objects in copy sector
  For i = 0 To 110000 Step 55
    If Mid(map, objstart + i, 2) <> zeroword And Asc(Mid(map, objstart + i + 4, 1)) + Asc(Mid(map, objstart + i + 6, 1)) < 256 And Asc(Mid(map, objstart + i + 4, 1)) < Asc(Mid(map, objstart + i + 6, 1)) Then
      mapobj(j) = Mid(map, objstart + i, 55)
      mapobjindx(k) = j 'Keep new record number for each object stored - for trigger adjustments
      jmax = j
      j = j + 1
    End If
    k = k + 1
  Next
  'Adjust triggers
  For j = 1 To jmax
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(j), 14 + k, 2) = WordChr(mapobjindx(ChrWord(Mid(mapobj(j), 14 + k, 2))))
        End If
      Next
    End If
  Next
  'Replicate objects in each sector
  For j = 1 To jmax
    mapobj(jmax + j) = mapobj(j)
    Mid(mapobj(jmax + j), 5, 1) = Chr(255 - Asc(Mid(mapobj(j), 7, 1)))
    Mid(mapobj(jmax + j), 7, 1) = Chr(255 - Asc(Mid(mapobj(j), 5, 1)))
    mapobj(jmax * 2 + j) = mapobj(j)
    Mid(mapobj(jmax * 2 + j), 5, 1) = Chr(255 - Asc(Mid(mapobj(j), 5, 1)))
    Mid(mapobj(jmax * 2 + j), 7, 1) = Chr(255 - Asc(Mid(mapobj(j), 7, 1)))
    mapobj(jmax * 3 + j) = mapobj(j)
    Mid(mapobj(jmax * 3 + j), 5, 1) = Chr(Asc(Mid(mapobj(j), 7, 1)))
    Mid(mapobj(jmax * 3 + j), 7, 1) = Chr(Asc(Mid(mapobj(j), 5, 1)))
    'Change orientation for 02-Buildings and 05-Scenery
    If Asc(Mid(mapobj(j), 2, 1)) = 2 Or Asc(Mid(mapobj(j), 2, 1)) = 5 Then
      If Asc(Mid(mapobj(j), 2, 1)) = 2 Then
        obyte = 9 'Orientation is in byte 9 for object group 02
      Else
        obyte = 12 'Orientation is in byte 12 for object group 05
      End If
      Orientation = Asc(Mid(mapobj(jmax + j), obyte, 1))
      Select Case Orientation
      Case 0 'S
        Mid(mapobj(jmax + j), obyte, 1) = Chr(6) 'E
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(4) 'N
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(2) 'W
      Case 1 'SW
        Mid(mapobj(jmax + j), obyte, 1) = Chr(5) 'NE
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(5) 'NE
      Case 2 'W
        Mid(mapobj(jmax + j), obyte, 1) = Chr(4) 'N
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(6) 'E
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(0) 'S
      Case 3 'NW
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(7) 'SE
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(7) 'SE
      Case 4 'N
        Mid(mapobj(jmax + j), obyte, 1) = Chr(2) 'W
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(0) 'S
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(6) 'E
      Case 5 'NE
        Mid(mapobj(jmax + j), obyte, 1) = Chr(1) 'SW
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(1) 'SW
      Case 6 'E
        Mid(mapobj(jmax + j), obyte, 1) = Chr(0) 'S
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(2) 'W
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(4) 'N
      Case 7 'SE
        Mid(mapobj(jmax * 2 + j), obyte, 1) = Chr(3) 'NW
        Mid(mapobj(jmax * 3 + j), obyte, 1) = Chr(3) 'NW
      Case Else
      End Select
    End If    'Convert Blue tribe objects to Red, Yellow, Green
    'Adjust triggers
    If Asc(Mid(mapobj(j), 1, 1)) = 6 And Asc(Mid(mapobj(j), 2, 1)) = 6 Then 'Trigger
      For k = 0 To 18 Step 2 'up to 10 record references may be held in a trigger
        If Mid(mapobj(j), 14 + k, 2) <> zeroword Then
          Mid(mapobj(jmax + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax + j), 14 + k, 2)) + jmax)
          Mid(mapobj(jmax * 2 + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax * 2 + j), 14 + k, 2)) + jmax * 2)
          Mid(mapobj(jmax * 3 + j), 14 + k, 2) = WordChr(ChrWord(Mid(mapobj(jmax * 3 + j), 14 + k, 2)) + jmax * 3)
        End If
      Next
    End If
    objbyte = Asc(Mid(mapobj(j), 3, 1))
    If objbyte <> 255 And Mid(mapobj(j), 1, 2) <> Chr(6) + Chr(6) Then 'Ignore Triggers (0606) and Wild colours
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax + j), 3, 1) = Chr(objbyte)
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax * 2 + j), 3, 1) = Chr(objbyte)
      objbyte = NextObjectColour(objbyte)
      Mid(mapobj(jmax * 3 + j), 3, 1) = Chr(objbyte)
    End If
  Next
  jmax = jmax * 4
  'Remove all objects from map
  For i = objstart To 110000 + objstart Step 2
    Mid(map, i, 2) = zeroword
  Next
  'Add all objects to map
  i = objstart
  For j = 1 To jmax
    Mid(map, i, 55) = mapobj(j)
    i = i + 55
  Next
  End If
End Sub
Private Function NextObjectColour(objcolour As Integer) As Integer
Select Case objcolour
Case 0
If optBlue(0).Value = True Then objcolour = 0
If optBlue(1).Value = True Then objcolour = 1
If optBlue(2).Value = True Then objcolour = 2
If optBlue(3).Value = True Then objcolour = 3
Case 1
If optRed(0).Value = True Then objcolour = 0
If optRed(1).Value = True Then objcolour = 1
If optRed(2).Value = True Then objcolour = 2
If optRed(3).Value = True Then objcolour = 3
Case 2
If optYellow(0).Value = True Then objcolour = 0
If optYellow(1).Value = True Then objcolour = 1
If optYellow(2).Value = True Then objcolour = 2
If optYellow(3).Value = True Then objcolour = 3
Case 3
If optGreen(0).Value = True Then objcolour = 0
If optGreen(1).Value = True Then objcolour = 1
If optGreen(2).Value = True Then objcolour = 2
If optGreen(3).Value = True Then objcolour = 3
Case Else
End Select
NextObjectColour = objcolour
End Function

Private Sub mnuNorth_Click()
On Error GoTo errorhandler
With dlgCommonDialog
 .CancelError = True
 .DialogTitle = "Save As"
 .Filter = "Bitmap (*.BMP)|*.bmp"
 .ShowSave
 If Len(.FileName) = 0 Then
   Exit Sub
 End If
 iFile = .FileName
 If Left(Right(iFile, 4), 1) = "." Then iFile = Left(iFile, Len(iFile) - 4) + ".BMP"
 SavePicture Picture1.Image, iFile
 msgval = MsgBox(iFile & " image saved", 64, "Populous Symmetry Tool")
End With
Exit Sub
errorhandler:
msgval = MsgBox("Error saving image " & iFile, 48, "Populous Symmetry Tool")
On Error GoTo 0
End Sub

Private Sub mnuShading_Click()
If mnuShading.Checked = True Then
mnuShading.Checked = False
Else
mnuShading.Checked = True
End If
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub mnuSouth_Click()
On Error GoTo errorhandler
With dlgCommonDialog
 .CancelError = True
 .DialogTitle = "Save As"
 .Filter = "Bitmap (*.BMP)|*.bmp"
 .ShowSave
 If Len(.FileName) = 0 Then
   Exit Sub
 End If
 iFile = .FileName
 If Left(Right(iFile, 4), 1) = "." Then iFile = Left(iFile, Len(iFile) - 4) + ".BMP"
 SavePicture Picture2.Image, iFile
 msgval = MsgBox(iFile & " image saved", 64, "Populous Symmetry Tool")
End With
Exit Sub
errorhandler:
msgval = MsgBox("Error saving image " & iFile, 48, "Populous Symmetry Tool")
On Error GoTo 0
End Sub

Private Sub MoveDown_Click()
Dim mapline(128) As String, X As Long, Y As Long, i As Long, scrollinc2 As Long
Call checkpoint("Move") 'take checkpoint for undo
'Move map down-------------------------
For Y = 0 To 127
  scrollinc2 = (Y + scrollinc) Mod 128
  mapline(Y) = Mid(map, scrollinc2 * 256 + 1, 256)
Next
For Y = 0 To 127
  Mid(map, Y * 256 + 1, 256) = mapline(Y)
Next
'adjust object locations down
For i = 0 To 110000 Step 55
  If Mid(map, objstart + i, 2) <> zeroword Then
    Mid(map, objstart + i + 6, 1) = Chr((Asc(Mid(map, objstart + i + 6, 1)) + (128 - scrollinc) * 2) Mod 256)
  End If
Next
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub MoveLeft_Click()
Dim mapline(128) As String, X As Long, Y As Long, i As Long, scrollinc2 As Long
Call checkpoint("Move") 'take checkpoint for undo
'Move map left-------------------------
scrollinc2 = scrollinc * 2
For Y = 0 To 127
  mapline(0) = Mid(map, Y * 256 + 1, 256)
  Mid(map, Y * 256 + 1, 256) = Mid(mapline(0), scrollinc2 + 1) + Mid(mapline(0), 1, scrollinc2)
Next
'adjust object locations left
For i = 0 To 110000 Step 55
  If Mid(map, objstart + i, 2) <> zeroword Then
    Mid(map, objstart + i + 4, 1) = Chr((Asc(Mid(map, objstart + i + 4, 1)) + (128 - scrollinc) * 2) Mod 256)
  End If
Next
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub MoveUp_Click()
Dim mapline(128) As String, X As Long, Y As Long, i As Long, scrollinc2 As Long
Call checkpoint("Move") 'take checkpoint for undo
'Move map up-------------------------
For Y = 0 To 127
  scrollinc2 = (Y + 128 - scrollinc) Mod 128
  mapline(Y) = Mid(map, scrollinc2 * 256 + 1, 256)
Next
For Y = 0 To 127
  Mid(map, Y * 256 + 1, 256) = mapline(Y)
Next
'adjust object locations up
For i = 0 To 110000 Step 55
  If Mid(map, objstart + i, 2) <> zeroword Then
    Mid(map, objstart + i + 6, 1) = Chr((Asc(Mid(map, objstart + i + 6, 1)) + scrollinc * 2) Mod 256)
  End If
Next
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub MoveRight_Click()
Dim mapline(128) As String, X As Long, Y As Long, i As Long, scrollinc2 As Long
Call checkpoint("Move") 'take checkpoint for undo
'Move map right-------------------------
scrollinc2 = (128 - scrollinc) * 2
For Y = 0 To 127
  mapline(0) = Mid(map, Y * 256 + 1, 256)
  Mid(map, Y * 256 + 1, 256) = Mid(mapline(0), scrollinc2 + 1) + Mid(mapline(0), 1, scrollinc2)
Next
'adjust object locations right
For i = 0 To 110000 Step 55
  If Mid(map, objstart + i, 2) <> zeroword Then
    Mid(map, objstart + i + 4, 1) = Chr((Asc(Mid(map, objstart + i + 4, 1)) + scrollinc * 2) Mod 256)
  End If
Next
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub Form_Load()
Dim j As Integer
map = String(192137, Chr(0))
For j = 0 To 9
   mapsave(j) = String(192137, Chr(0))
Next
objstart = 81988
scrollinc = 16 'default scroll is 1/8th of the map
txtRotate = "1/8"
SymmetryType = 2
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub Displaymap()
Dim X As Long, Y As Long, mapspot As Long, pixcolour As Long
Dim zeroword As String, tribe As Integer, Shade As Integer
Picture1.AutoRedraw = True
zeroword = Chr(0) + Chr(0)
'Display land
Picture1.FillStyle = 0
For X = 0 To 127
For Y = 0 To 127
'Shade is the difference added to each colour to create the source area shading
If mnuShading.Checked = True Then
  Shade = 50
  Select Case SymmetryType
  Case 0
    If X > 63 Then Shade = 0
  Case 1
    If Y < 64 Then Shade = 0
  Case 2
    If X > 63 Then Shade = 0
    If Y < 64 Then Shade = 0
  Case 3
    If X >= Y Then Shade = 0
  Case 4
    If X + Y > 126 Then Shade = 0
  Case 5
    If X >= Y Then Shade = 0
    If X + Y > 126 Then Shade = 0
  End Select
Else
  Shade = 0 'No Shading highlight
End If
mapspot = Asc(Mid(map, (X * 2) + ((127 - Y) * 256) + 2, 1))
mapspot = mapspot * 256
mapspot = mapspot + Asc(Mid(map, (X * 2) + ((127 - Y) * 256) + 1, 1))
Select Case mapspot
Case 0
pixcolour = RGB(135 + Shade, 156 + Shade, 227 + Shade)
Case Is < 8
pixcolour = RGB(HSLtoRGB(24, 224, 200).Red + Shade, HSLtoRGB(24, 224, 200).Green + Shade, HSLtoRGB(24, 224, 200).Blue + Shade)
Case Is > 1023
pixcolour = RGB(HSLtoRGB(24, 224, 40).Red + Shade, HSLtoRGB(24, 224, 40).Green + Shade, HSLtoRGB(24, 224, 40).Blue + Shade)
Case Else
mapspot = 178 - (mapspot / 8)
pixcolour = RGB(HSLtoRGB(24, 224, mapspot).Red + Shade, HSLtoRGB(24, 224, mapspot).Green + Shade, HSLtoRGB(24, 224, mapspot).Blue + Shade)
End Select
Picture1.FillColor = pixcolour
Picture1.Line (X * 3, Y * 3)-(X * 3 + 3, Y * 3 + 3), pixcolour, B
Next
Next
If mnuOptionsShowObjects.Checked = True Then
'Display objects
 For i = 0 To 110000 Step 55
   If Mid(map, objstart + i, 2) <> zeroword Then
     X = Asc(Mid(map, objstart + i + 4, 1)) * 3 / 2
     Y = (255 - Asc(Mid(map, objstart + i + 6, 1))) * 3 / 2
     tribe = Asc(Mid(map, objstart + i + 2, 1))
     Select Case tribe
     Case 0
       pixcolour = RGB(0, 0, 255)
     Case 1
       pixcolour = RGB(255, 0, 0)
     Case 2
       pixcolour = RGB(255, 255, 0)
     Case 3
       pixcolour = RGB(0, 255, 0)
     Case Else
       pixcolour = RGB(255, 255, 255)
     End Select
      Picture1.FillColor = pixcolour
      Picture1.Line (X, Y)-(X + 3, Y + 3), 0, B
    End If
  Next
End If
End Sub
Private Sub DisplaySouthPole()
Picture2.AutoRedraw = True
Picture2.PaintPicture Picture1.Image, 0, 0, 192, 192, 192, 192, 192, 192
Picture2.PaintPicture Picture1.Image, 0, 192, 192, 192, 192, 0, 192, 192
Picture2.PaintPicture Picture1.Image, 192, 0, 192, 192, 0, 192, 192, 192
Picture2.PaintPicture Picture1.Image, 192, 192, 192, 192, 0, 0, 192, 192
End Sub

Private Sub HScroll1_Change()
Select Case HScroll1.Value
Case 0
txtRotate.Text = "1/128"
scrollinc = 1
Case 1
txtRotate.Text = "1/64"
scrollinc = 2
Case 2
txtRotate.Text = "1/32"
scrollinc = 4
Case 3
txtRotate.Text = "1/16"
scrollinc = 8
Case 4
txtRotate.Text = "1/8"
scrollinc = 16
Case 5
txtRotate.Text = "1/4"
scrollinc = 32
Case Else
txtRotate.Text = "1/2"
scrollinc = 64
End Select
End Sub

Private Sub mnuHelpFile_Click()
'RetVal = Shell("winhlp32.exe " + App.Path + "\Pes_help.HLP", 1) 'Run help KHICK'S ONE
RetVal = Shell("notepad.exe " + App.Path + "\Readme-Help.txt", 1) 'Run help
End Sub

Private Sub mnuOptionsShowLines_Click()
If mnuOptionsShowLines.Checked = True Then
mnuOptionsShowLines.Checked = False
Else
mnuOptionsShowLines.Checked = True
End If
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub mnuOptionsShowObjects_Click()
If mnuOptionsShowObjects.Checked = True Then
mnuOptionsShowObjects.Checked = False
Else
mnuOptionsShowObjects.Checked = True
End If
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub mnuOptionsShowSouthPole_Click()
If mnuOptionsShowSouthPole.Checked = True Then
mnuOptionsShowSouthPole.Checked = False
Width = 6150
Else
mnuOptionsShowSouthPole.Checked = True
Width = 12165
Call DisplaySouthPole
End If
End Sub

Private Sub optAxis_Click(Index As Integer)
SymmetryType = Index 'Store Symmetry Type
Call Displaymap 'remove any previous lines
If mnuOptionsShowSouthPole.Checked = True Then Call DisplaySouthPole
If mnuOptionsShowLines.Checked = True Then
  Select Case Index
  Case 0  'Vertical Line
    Picture1.Line (0, 0)-(0, 383), 0
    Picture1.Line (191, 0)-(191, 383), 0
    Picture1.Line (383, 0)-(383, 383), 0
  Case 1  'Horizontal Line
    Picture1.Line (0, 0)-(383, 0), 0
    Picture1.Line (0, 191)-(383, 191), 0
    Picture1.Line (0, 383)-(383, 383), 0
  Case 2  'Both vert & Horiz
    Picture1.Line (0, 0)-(0, 383), 0
    Picture1.Line (191, 0)-(191, 383), 0
    Picture1.Line (383, 0)-(383, 383), 0
    Picture1.Line (0, 0)-(383, 0), 0
    Picture1.Line (0, 191)-(383, 191), 0
    Picture1.Line (0, 383)-(383, 383), 0
  Case 3  'TL-BR diagonal
    Picture1.Line (0, 0)-(383, 383), 0
  Case 4  'BL-TR diagonal
    Picture1.Line (0, 383)-(383, 0), 0
  Case 5  'Both Diagonal
    Picture1.Line (0, 0)-(383, 383), 0
    Picture1.Line (0, 383)-(383, 0), 0
  Case Else
  End Select
  If mnuOptionsShowSouthPole.Checked = True Then
  Select Case Index
  Case 0  'Vertical Line
    Picture2.Line (0, 0)-(0, 383), 0
    Picture2.Line (191, 0)-(191, 383), 0
    Picture2.Line (383, 0)-(383, 383), 0
  Case 1  'Horizontal Line
    Picture2.Line (0, 0)-(383, 0), 0
    Picture2.Line (0, 191)-(383, 191), 0
    Picture2.Line (0, 383)-(383, 383), 0
  Case 2  'Both vert & Horiz
    Picture2.Line (0, 0)-(0, 383), 0
    Picture2.Line (191, 0)-(191, 383), 0
    Picture2.Line (383, 0)-(383, 383), 0
    Picture2.Line (0, 0)-(383, 0), 0
    Picture2.Line (0, 191)-(383, 191), 0
    Picture2.Line (0, 383)-(383, 383), 0
  Case 3  'TL-BR diagonal
    Picture2.Line (0, 0)-(383, 383), 0
  Case 4  'BL-TR diagonal
    Picture2.Line (0, 383)-(383, 0), 0
  Case 5  'Both Diagonal
    Picture2.Line (0, 0)-(383, 383), 0
    Picture2.Line (0, 383)-(383, 0), 0
  Case Else
  End Select
  End If
End If

End Sub

Private Sub optDirection_Click(Index As Integer)
Dim mapline(128) As String, X As Long, Y As Long, i As Long, scrollinc2 As Long
'Move map right-------------------------
  scrollinc2 = (128 - scrollinc) * 2
  For Y = 0 To 127
    mapline(0) = Mid(map, Y * 256 + 1, 256)
    Mid(map, Y * 256 + 1, 256) = Mid(mapline(0), scrollinc2 + 1) + Mid(mapline(0), 1, scrollinc2)
  Next
  'adjust object locations right
  For i = 0 To 110000 Step 55
    If Mid(map, objstart + i, 2) <> zeroword Then
      Mid(map, objstart + i + 4, 1) = Chr((Asc(Mid(map, objstart + i + 4, 1)) + scrollinc * 2) Mod 256)
    End If
  Next
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    End

End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo errorhandler
With dlgCommonDialog
 .CancelError = True
 .DialogTitle = "Save As"
 .Filter = "PopulousLevelFiles (LEVL*.DAT)|LEVL*.DAT"
 .ShowSave
 If Len(.FileName) = 0 Then
   Exit Sub
 End If
 sFile = .FileName
 Caption = .FileTitle + " - Populous Symmetry Tool"
 Call mnuFileSave_Click
End With
Exit Sub
errorhandler:
On Error GoTo 0
Close
End Sub

Private Sub mnuFileSave_Click()
Dim msgval
On Error GoTo errorhandler
Open sFile For Binary As #1
Put #1, 1, map
Close #1
msgval = MsgBox(sFile & " saved", 64, "Populous Symmetry Tool")
Exit Sub
errorhandler:
On Error GoTo 0
msgval = MsgBox("Error saving " & sFile & ". Check that file properties are not set to read-only.", 48, "Populous Symmetry Tool")
Close
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo Openerror
Call checkpoint("Open") 'take checkpoint for undo
    dlgCommonDialog.InitDir = "c:\Program Files\Bullfrog\Populous\levels"
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Populous Level Files (LEVL*.DAT)|LEVL*.DAT|Level Files (*.DAT)|*.DAT"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        Caption = .FileTitle + " - Populous Symmetry Tool"
     End With
Open sFile For Binary As #1
map = String(192137, 0)
Get #1, , map
Close #1
mnuFileSave.Enabled = True
mnuFileSaveAs.Enabled = True
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
Openerror: 'Cancel pressed
End Sub

Private Sub checkpoint(checkpointtype As String)
If checkpointtype = "Move" And LastCheckpointType = "Move" Then Exit Sub 'only checkpoint first of a move sequence
LastCheckpointType = checkpointtype
mapsave(CheckpointNumber) = map
If checkpointtype <> "Undo" Then
  CheckpointNumber = (CheckpointNumber + 1) Mod 10
  If CheckpointCount < 9 Then CheckpointCount = CheckpointCount + 1
  mnuUndo.Enabled = True
  MnuRedo.Enabled = False
  UndoCount = 0
End If
End Sub

Private Sub mnuUndo_Click()
If UndoCount = 0 Then Call checkpoint("Undo")  'checkpoint in case of Redo
CheckpointNumber = (CheckpointNumber - 1)
If CheckpointNumber = -1 Then CheckpointNumber = 9
map = mapsave(CheckpointNumber)
CheckpointCount = CheckpointCount - 1
UndoCount = UndoCount + 1
MnuRedo.Enabled = True
If CheckpointCount = 0 Then mnuUndo.Enabled = False
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub

Private Sub mnuRedo_Click()
CheckpointNumber = (CheckpointNumber + 1) Mod 10
If CheckpointCount < 9 Then CheckpointCount = CheckpointCount + 1
map = mapsave(CheckpointNumber)
UndoCount = UndoCount - 1
If UndoCount = 0 Then MnuRedo.Enabled = False
mnuUndo.Enabled = True
'Display map with grid according to symmetry option
optAxis_Click (SymmetryType)
End Sub
