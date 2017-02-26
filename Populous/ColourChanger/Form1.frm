VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Populous Color changer"
   ClientHeight    =   11055
   ClientLeft      =   7005
   ClientTop       =   4080
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   7095
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "Form1.frx":0000
      Top             =   12720
      Width           =   6855
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8415
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   14843
      _Version        =   393216
      Rows            =   30
      Cols            =   6
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "(none)"
      Top             =   2760
      Width           =   6855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Default"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove"
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ListBox List3 
      Height          =   645
      ItemData        =   "Form1.frx":0012
      Left            =   6000
      List            =   "Form1.frx":0022
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   645
      ItemData        =   "Form1.frx":0040
      Left            =   2520
      List            =   "Form1.frx":0050
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "Form1.frx":006E
      Left            =   960
      List            =   "Form1.frx":007E
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog open1 
      Left            =   5280
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Swap"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   7080
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label7 
      Caption         =   "Graham@brambles.org"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "Ted@brambles.org"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Designed by Edward and Graham John"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7080
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Remove"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "With"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Swap"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Populous Color changer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public map As String
Dim objcount(100, 4)
Dim objname(100)
Public objcoulour As Integer
Public objstart As Long
Private Sub Check1_Click()
If Check1.Value = 1 Then
Combo1.Enabled = True
Combo2.Enabled = True
Combo1.Text = "Blue"
Combo2.Text = "Red"
Else: Combo1.Enabled = False
Combo2.Enabled = False
Combo1.Text = "-"
Combo2.Text = "-"
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Combo3.Enabled = True
Combo3.Text = "Yellow"
Else: Combo3.Enabled = False
Combo3.Text = "-"
End If
End Sub

Private Sub Command1_Click()
On Error GoTo endload
Text2.Text = "Unknown objects: "
open1.Filter = "PopulousLevelFiles (LEVL*.DAT)|LEVL*.DAT"
open1.ShowOpen
Open open1.FileName For Binary As #1
Text1 = open1.FileName
map = String(192137, 0)
Get #1, , map
Close #1
teststring = Chr$(1) + Chr$(0) + Chr$(1) + String(13, Chr$(0)) + Chr$(2) + Chr$(0) + Chr$(2) + String(13, Chr$(0)) + Chr$(3) + Chr$(0) + Chr$(3) + String(13, Chr$(0)) + Chr$(4) + Chr$(0) + Chr$(4) + String(13, Chr$(0))
For i = 1 To 192137
  If Mid(map, i, 64) = teststring Then GoTo search
Next
MsgBox ("Start of map object data not found")
Exit Sub
search:
objstart = i + 67
'MsgBox ("Start of object data is at byte " & objstart)
Grid1.TextMatrix(0, 5) = "Wild"
Grid1.TextMatrix(0, 4) = "Green"
Grid1.TextMatrix(0, 3) = "Yellow"
Grid1.TextMatrix(0, 2) = "Red"
Grid1.TextMatrix(0, 1) = "Blue"
Grid1.TextMatrix(0, 0) = "Object"
Grid1.TextMatrix(1, 0) = "Wildman"
Grid1.TextMatrix(2, 0) = "Brave"
Grid1.TextMatrix(3, 0) = "Warrior"
Grid1.TextMatrix(4, 0) = "Preacher"
Grid1.TextMatrix(5, 0) = "Spy"
Grid1.TextMatrix(6, 0) = "Fire Warrior"
Grid1.TextMatrix(7, 0) = "Shaman"
Grid1.TextMatrix(8, 0) = "Angel of Death"
Grid1.TextMatrix(9, 0) = "Small Hut"
Grid1.TextMatrix(10, 0) = "Medium Hut"
Grid1.TextMatrix(11, 0) = "Large Hut"
Grid1.TextMatrix(12, 0) = "Guard Tower"
Grid1.TextMatrix(13, 0) = "Temple"
Grid1.TextMatrix(14, 0) = "Spy Hut"
Grid1.TextMatrix(15, 0) = "Warrior Hut"
Grid1.TextMatrix(16, 0) = "Fire Warrior Hut"
Grid1.TextMatrix(17, 0) = "Boat Hut"
Grid1.TextMatrix(18, 0) = "Balloon Hut"
Grid1.TextMatrix(19, 0) = "Prison"
Grid1.TextMatrix(20, 0) = "Site of Worship"
Grid1.TextMatrix(21, 0) = "Gift of Worship"
Grid1.TextMatrix(22, 0) = "Stone Head/Totem Pole"
Grid1.TextMatrix(23, 0) = "Vault of Knowledge"
Grid1.TextMatrix(24, 0) = "Small Tree"
Grid1.TextMatrix(25, 0) = "Large Tree"
Grid1.TextMatrix(26, 0) = "Crooked Tree"
Grid1.TextMatrix(27, 0) = "Other object"
Grid1.TextMatrix(28, 0) = "Totals"
For i = 1 To 28
For j = 1 To 5
 Grid1.TextMatrix(i, j) = 0
Next
Next
For j = objstart To 192137 Step 55
Select Case Mid(map, j + 2, 1)
Case Chr$(0)
  objcolour = 1
Case Chr$(1)
  objcolour = 2
Case Chr$(2)
  objcolour = 3
Case Chr$(3)
  objcolour = 4
Case Else
  objcolour = 5
End Select
If Mid(map, j, 2) <> Chr$(0) & Chr$(0) Then
Select Case Mid(map, j, 2)
Case Chr$(1) & Chr$(1)
Grid1.TextMatrix(1, objcolour) = Grid1.TextMatrix(1, objcolour) + 1
Case Chr$(2) & Chr$(1)
Grid1.TextMatrix(2, objcolour) = Grid1.TextMatrix(2, objcolour) + 1
Case Chr$(3) & Chr$(1)
Grid1.TextMatrix(3, objcolour) = Grid1.TextMatrix(3, objcolour) + 1
Case Chr$(4) & Chr$(1)
Grid1.TextMatrix(4, objcolour) = Grid1.TextMatrix(4, objcolour) + 1
Case Chr$(5) & Chr$(1)
Grid1.TextMatrix(5, objcolour) = Grid1.TextMatrix(5, objcolour) + 1
Case Chr$(6) & Chr$(1)
Grid1.TextMatrix(6, objcolour) = Grid1.TextMatrix(6, objcolour) + 1
Case Chr$(7) & Chr$(1)
Grid1.TextMatrix(7, objcolour) = Grid1.TextMatrix(7, objcolour) + 1
Case Chr$(8) & Chr$(1)
Grid1.TextMatrix(8, objcolour) = Grid1.TextMatrix(8, objcolour) + 1
Case Chr$(1) & Chr$(2)
Grid1.TextMatrix(9, objcolour) = Grid1.TextMatrix(9, objcolour) + 1
Case Chr$(2) & Chr$(2)
Grid1.TextMatrix(10, objcolour) = Grid1.TextMatrix(10, objcolour) + 1
Case Chr$(3) & Chr$(2)
Grid1.TextMatrix(11, objcolour) = Grid1.TextMatrix(11, objcolour) + 1
Case Chr$(4) & Chr$(2)
Grid1.TextMatrix(12, objcolour) = Grid1.TextMatrix(12, objcolour) + 1
Case Chr$(5) & Chr$(2)
Grid1.TextMatrix(13, objcolour) = Grid1.TextMatrix(13, objcolour) + 1
Case Chr$(6) & Chr$(2)
Grid1.TextMatrix(14, objcolour) = Grid1.TextMatrix(14, objcolour) + 1
Case Chr$(7) & Chr$(2)
Grid1.TextMatrix(15, objcolour) = Grid1.TextMatrix(15, objcolour) + 1
Case Chr$(8) & Chr$(2)
Grid1.TextMatrix(16, objcolour) = Grid1.TextMatrix(16, objcolour) + 1
Case Chr$(13) & Chr$(2)
Grid1.TextMatrix(17, objcolour) = Grid1.TextMatrix(17, objcolour) + 1
Case Chr$(15) & Chr$(2)
Grid1.TextMatrix(18, objcolour) = Grid1.TextMatrix(18, objcolour) + 1
Case Chr$(19) & Chr$(2)
Grid1.TextMatrix(19, objcolour) = Grid1.TextMatrix(19, objcolour) + 1
Case Chr$(6) & Chr$(6)
Grid1.TextMatrix(20, objcolour) = Grid1.TextMatrix(20, objcolour) + 1
Case Chr$(2) & Chr$(6)
Grid1.TextMatrix(21, objcolour) = Grid1.TextMatrix(21, objcolour) + 1
Case Chr$(9) & Chr$(5)
Grid1.TextMatrix(22, objcolour) = Grid1.TextMatrix(22, objcolour) + 1
Case Chr$(18) & Chr$(2)
Grid1.TextMatrix(23, objcolour) = Grid1.TextMatrix(23, objcolour) + 1
Case Chr$(1) & Chr$(5)
Grid1.TextMatrix(24, objcolour) = Grid1.TextMatrix(24, objcolour) + 1
Case Chr$(2) & Chr$(5)
Grid1.TextMatrix(25, objcolour) = Grid1.TextMatrix(25, objcolour) + 1
Case Chr$(3) & Chr$(5)
Grid1.TextMatrix(26, objcolour) = Grid1.TextMatrix(26, objcolour) + 1
Case Chr$(5) & Chr$(5) 'This object is grouped in with large trees because there is no discernable difference
Grid1.TextMatrix(26, objcolour) = Grid1.TextMatrix(26, objcolour) + 1
Case Else
Grid1.TextMatrix(27, objcolour) = Grid1.TextMatrix(27, objcolour) + 1
'List unknown objects in the text window
hexchr1 = Hex(Asc(Mid(map, j, 1)))
If Len(hexchr1) = 1 Then hexchr1 = "0" & hexchr1
hexchr2 = Hex(Asc(Mid(map, j + 1, 1)))
If Len(hexchr2) = 1 Then hexchr2 = "0" & hexchr2
Text2.Text = Text2.Text & " " & hexchr1 & hexchr2 & ", "
End Select
Grid1.TextMatrix(28, objcolour) = Grid1.TextMatrix(28, objcolour) + 1
End If
Next
endload:
End Sub

Private Sub Command2_Click()
changes = 0
objects0 = 0
objects1 = 0
objects2 = 0
objects3 = 0
objects9 = 0
For j = objstart + 2 To 192317 Step 55
If Mid(map, j - 2, 2) <> Chr$(0) & Chr$(0) Then
  If Mid(map, j, 1) = Chr$(List1.ItemData(List1.ListIndex)) Then
    Mid(map, j, 1) = Chr$(List2.ItemData(List2.ListIndex))
   changes = changes + 1
  Else
    If Mid(map, j, 1) = Chr$(List2.ItemData(List2.ListIndex)) Then
    Mid(map, j, 1) = Chr$(List1.ItemData(List1.ListIndex))
    changes = changes + 1
    End If
  End If
Select Case (Mid(map, j, 1))
 Case Chr$(0)
 objects0 = objects0 + 1
  Case Chr$(1)
 objects1 = objects1 + 1
 Case Chr$(2)
 objects2 = objects2 + 1
 Case Chr$(3)
 objects3 = objects3 + 1
 Case Else
 objects9 = objects9 + 1
 End Select
End If
Next
MsgBox (changes & " changes made. Object counts are Blue=" & objects0 & " Red=" & objects1 & " Yellow=" & objects2 & " Green=" & objects3 & " Other=" & objects9)
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command7_Click()
MsgBox (open1.FileName)
Open open1.FileName For Binary As #1
Put #1, 1, map
Close #1
End Sub

Private Sub Command8_Click()
Form1.Height = 4635
Command7.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Command9_Click()
Form1.Height = 3675
Command7.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Form_Load()
open1.InitDir = "c:\Program Files\Bullfrog\Populous\levels"
End Sub
