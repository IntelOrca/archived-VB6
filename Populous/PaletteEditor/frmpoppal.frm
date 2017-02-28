VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Populous Palette Editor"
   ClientHeight    =   7455
   ClientLeft      =   6435
   ClientTop       =   3405
   ClientWidth     =   10545
   Icon            =   "frmpoppal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   13150
      _Version        =   393216
      Rows            =   257
      Cols            =   9
      FixedCols       =   3
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
   End
   Begin MSComDlg.CommonDialog open1 
      Left            =   10920
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu Saveas 
         Caption         =   "Save &As"
         Shortcut        =   ^S
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu undo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "Co&py"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "Pa&ste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu Replace 
         Caption         =   "&Replace"
         Shortcut        =   ^R
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu editcol 
         Caption         =   "E&dit Colour"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTribes 
         Caption         =   "Edit &Tribes"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu help2 
         Caption         =   "He&lp"
         Shortcut        =   {F1}
      End
      Begin VB.Menu About 
         Caption         =   "&About Populous Palette Editor"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Red, Green, Blue As Long
Public SaveRed, SaveGreen, SaveBlue As Integer
Public UndoColour As Long
Public UndoRow As Integer
Public PaletteFileName As String

Private Sub About_Click()
frmAbout.Show
End Sub
Private Sub Copy_Click()
SaveRed = Grid1.TextMatrix(Grid1.Row, 3)
SaveGreen = Grid1.TextMatrix(Grid1.Row, 4)
SaveBlue = Grid1.TextMatrix(Grid1.Row, 5)
Paste.Enabled = True
End Sub

Private Sub Cut_Click()
Call Set_UnDo
SaveRed = Grid1.TextMatrix(Grid1.Row, 3)
SaveGreen = Grid1.TextMatrix(Grid1.Row, 4)
SaveBlue = Grid1.TextMatrix(Grid1.Row, 5)
Paste.Enabled = True
Grid1.TextMatrix(Grid1.Row, 3) = 0
Grid1.TextMatrix(Grid1.Row, 4) = 0
Grid1.TextMatrix(Grid1.Row, 5) = 0
Grid1.TextMatrix(Grid1.Row, 6) = 0
Grid1.TextMatrix(Grid1.Row, 7) = 0
Grid1.TextMatrix(Grid1.Row, 8) = 0
Call Colour_Changed
End Sub

Private Sub editcol_Click()
Const L65536 As Long = 65536
Const L256 As Long = 256
Call Set_UnDo
open1.ShowColor
If open1.Color = 0 Then Exit Sub
Grid1.Col = 2
Grid1.CellBackColor = open1.Color
If Grid1.CellBackColor = 0 Then
 Grid1.CellBackColor = 1
End If
colour = open1.Color
Blue = colour \ L65536
Green = (colour - (Blue * L65536)) \ L256
Red = (colour - (Blue * L65536) - (Green * L256))
Grid1.TextMatrix(Grid1.Row, 3) = Red
Grid1.TextMatrix(Grid1.Row, 4) = Green
Grid1.TextMatrix(Grid1.Row, 5) = Blue
Grid1.TextMatrix(Grid1.Row, 6) = RGBtoHSL(Red, Green, Blue).Hue
Grid1.TextMatrix(Grid1.Row, 7) = RGBtoHSL(Red, Green, Blue).Saturation
Grid1.TextMatrix(Grid1.Row, 8) = RGBtoHSL(Red, Green, Blue).Luminance
Grid1.Col = 3
End Sub

Private Sub Form_Resize()
On Error Resume Next
Grid1.Height = Me.ScaleHeight
Grid1.Width = 703
Me.Width = 10650
End Sub

Private Sub mnuTribes_Click()
frmTribes.Show 1
End Sub

Private Sub new_Click()
Save.Enabled = False
For i = 1 To 256
Grid1.TextMatrix(i, 0) = i
Grid1.TextMatrix(i, 3) = 0
Grid1.TextMatrix(i, 4) = 0
Grid1.TextMatrix(i, 5) = 0
Grid1.TextMatrix(i, 6) = 0
Grid1.TextMatrix(i, 7) = 0
Grid1.TextMatrix(i, 8) = 0
Grid1.Col = 2
Grid1.Row = i
Grid1.CellBackColor = 1
Next
Grid1.Col = 3
Grid1.Row = 1
undo.Enabled = False
End Sub

Private Sub Paste_Click()
Grid1.TextMatrix(Grid1.Row, 3) = SaveRed
Grid1.TextMatrix(Grid1.Row, 4) = SaveGreen
Grid1.TextMatrix(Grid1.Row, 5) = SaveBlue
Call Colour_Changed
End Sub

Private Sub Replace_Click()
frmReplace.Show
End Sub

Private Sub Save_Click()
On Error GoTo errorhandler
poppal = ""
For i = 1 To 256
  poppal = poppal + Chr$(Grid1.TextMatrix(i, 3)) + Chr$(Grid1.TextMatrix(i, 4)) + Chr$(Grid1.TextMatrix(i, 5)) + Chr$(0)
Next
Open PaletteFileName For Binary As #1
Put #1, 1, poppal
Close #1
Msgval = MsgBox(PaletteFileName & " saved", 64, "Populous Palette Editor")
Exit Sub
errorhandler:
On Error GoTo 0
Msgval = MsgBox("Error saving " & PaletteFileName & ". Check that file properties are not set to read-only.", 48, "Populous Palette Editor")
Close
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
Dim textline As String
open1.InitDir = "c:\Program Files\Bullfrog\Populous\data"
For i = 0 To 255
Grid1.TextMatrix(i + 1, 0) = i
Grid1.TextMatrix(i + 1, 3) = 0
Grid1.TextMatrix(i + 1, 4) = 0
Grid1.TextMatrix(i + 1, 5) = 0
Grid1.TextMatrix(i + 1, 6) = 0
Grid1.TextMatrix(i + 1, 7) = 0
Grid1.TextMatrix(i + 1, 8) = 0
Next
Grid1.ColWidth(0) = 600
Grid1.ColAlignment(0) = 3
Grid1.TextMatrix(0, 0) = "Index"
Grid1.ColWidth(1) = 5000
Grid1.ColAlignment(1) = 0
Grid1.TextMatrix(0, 1) = "Use in Populous"
Grid1.ColWidth(2) = 1000
Grid1.ColAlignment(2) = 3
Grid1.TextMatrix(0, 2) = "Colour"
Grid1.ColWidth(3) = 600
Grid1.ColAlignment(3) = 3
Grid1.TextMatrix(0, 3) = "Red"
Grid1.ColWidth(4) = 600
Grid1.ColAlignment(4) = 3
Grid1.TextMatrix(0, 4) = "Green"
Grid1.ColWidth(5) = 600
Grid1.ColAlignment(5) = 3
Grid1.TextMatrix(0, 5) = "Blue"
Grid1.ColWidth(6) = 600
Grid1.ColAlignment(6) = 3
Grid1.TextMatrix(0, 6) = "Hue"
Grid1.ColWidth(7) = 600
Grid1.ColAlignment(7) = 3
Grid1.TextMatrix(0, 7) = "Sat."
Grid1.ColWidth(8) = 600
Grid1.ColAlignment(8) = 3
Grid1.TextMatrix(0, 8) = "Lum."
On Error GoTo errorhandler
Open App.Path + "\Descriptions.txt" For Input As #2
Do While Not EOF(2)   ' Loop until end of file.
  Line Input #2, textline
  lineno = Left$(textline, 3)
  If lineno >= 0 And lineno <= 255 Then
    Grid1.Row = lineno + 1
    Grid1.Col = 1
    Grid1.Text = Mid$(textline, 5)
    Grid1.Col = 2
    Grid1.CellBackColor = 1
  End If
Loop
Close #2
Grid1.Row = 1
Grid1.Col = 3
undo.Enabled = False
Exit Sub
errorhandler:
On Error GoTo 0
Msgval = MsgBox("Error reading " & App.Path + "\Descriptions.txt", 48, "Populous Palette Editor")
Close
End Sub

Private Sub Grid1_DblClick()
Call editcol_Click
End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
With Grid1
Select Case KeyCode
Case vbKeyDelete
.Text = 0
Call Colour_Changed
End Select
End With
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
Dim chkint As Integer
   With Grid1
        Select Case KeyAscii
                
            Case 8: 'IF KEY IS BACKSPACE THEN
                If .Text <> "" Then
                 .Text = Left$(.Text, (Len(.Text) - 1))
                 If .Text = "" Then
                 .Text = 0
                 End If
                 Call Colour_Changed
                End If
            Case 13: 'IF KEY IS ENTER THEN
                Select Case .Col
                    Case Is < (.Cols - 1):
                        SendKeys "{right}"
                    Case (.Cols - 1):
                        If (.Row + 1) = .Rows Then
                            .Rows = .Rows + 1
                        End If
                        SendKeys "{home}" + "{down}"
                End Select
            Case Else
                Select Case .Col
                    Case 3, 4, 5:
                        If Chr$(KeyAscii) < "0" Or Chr$(KeyAscii) > "9" Then
                        Else
                          .Text = .Text + Chr$(KeyAscii)
                          'remove leading zeros and check for max value
                          chkint = .Text
                          If chkint > 255 Then
                            .Text = 255
                          Else
                            .Text = chkint
                          End If
                          Call Colour_Changed
                        End If
                    Case 6:
                        If Chr$(KeyAscii) < "0" Or Chr$(KeyAscii) > "9" Then
                        Else
                          .Text = .Text + Chr$(KeyAscii)
                          'remove leading zeros and check for max value
                          chkint = .Text
                          If chkint > 240 Then
                            .Text = 240
                          Else
                            .Text = chkint
                          End If
                          Call Colour_Changed
                        End If
                    Case 7, 8:
                        If Chr$(KeyAscii) < "0" Or Chr$(KeyAscii) > "9" Then
                        Else
                          .Text = .Text + Chr$(KeyAscii)
                          'remove leading zeros and check for max value
                          chkint = .Text
                          If chkint > 239 Then
                            .Text = 239
                          Else
                            .Text = chkint
                          End If
                          Call Colour_Changed
                        End If
                    Case Else:
                End Select
        End Select
    End With

End Sub
Public Sub Colour_Changed()
Dim savecol As Integer
Call Set_UnDo
With Grid1
  Select Case .Col
    Case 3, 4, 5
      .TextMatrix(.Row, 6) = RGBtoHSL(Grid1.TextMatrix(.Row, 3), Grid1.TextMatrix(.Row, 4), Grid1.TextMatrix(.Row, 5)).Hue
      .TextMatrix(.Row, 7) = RGBtoHSL(Grid1.TextMatrix(.Row, 3), Grid1.TextMatrix(.Row, 4), Grid1.TextMatrix(.Row, 5)).Saturation
      .TextMatrix(.Row, 8) = RGBtoHSL(Grid1.TextMatrix(.Row, 3), Grid1.TextMatrix(.Row, 4), Grid1.TextMatrix(.Row, 5)).Luminance
    Case Else
      .TextMatrix(.Row, 3) = HSLtoRGB(Grid1.TextMatrix(.Row, 6), Grid1.TextMatrix(.Row, 7), Grid1.TextMatrix(.Row, 8)).Red
      .TextMatrix(.Row, 4) = HSLtoRGB(Grid1.TextMatrix(.Row, 6), Grid1.TextMatrix(.Row, 7), Grid1.TextMatrix(.Row, 8)).Green
      .TextMatrix(.Row, 5) = HSLtoRGB(Grid1.TextMatrix(.Row, 6), Grid1.TextMatrix(.Row, 7), Grid1.TextMatrix(.Row, 8)).Blue
  End Select
savecol = .Col
.Col = 2
.CellBackColor = RGB(Grid1.TextMatrix(.Row, 3), Grid1.TextMatrix(.Row, 4), Grid1.TextMatrix(.Row, 5))
 If .CellBackColor = 0 Then
 .CellBackColor = 1
 End If
.Col = savecol
End With
End Sub

Private Sub help2_Click()
RetVal = Shell("Notepad.exe " + App.Path + "\Readme-help.txt", 1) ' Run Notepad
End Sub

Public Sub Open_Click()
On Error GoTo errorhandler
open1.Filter = "PopulousPaletteFiles (PAL*.DAT)|PAL*.DAT"
open1.ShowOpen
palfilesize = 0
If open1.FileName <> "" Then
palfilesize = FileLen(open1.FileName)
If palfilesize = 1024 Then
  Open open1.FileName For Binary As #1
  PaletteFileName = open1.FileName
  poppal = String(1024, 0)
  Get #1, , poppal
  Close #1
Else
  Msgval = MsgBox("The file length of " + open1.FileName + " is " + Format(palfilesize, "General Number") + " bytes. This is not a 1024 byte palette file.", 48, "Populous Palette Editor")
  Exit Sub
End If
For i = 0 To 255
 Red = Asc(Mid(poppal, i * 4 + 1, 1))
 Green = Asc(Mid(poppal, i * 4 + 2, 1))
 Blue = Asc(Mid(poppal, i * 4 + 3, 1))
 Grid1.TextMatrix(i + 1, 3) = Red
 Grid1.TextMatrix(i + 1, 4) = Green
 Grid1.TextMatrix(i + 1, 5) = Blue
 Grid1.Row = i + 1
 Grid1.Col = 2
 Grid1.CellBackColor = RGB(Grid1.TextMatrix(i + 1, 3), Grid1.TextMatrix(i + 1, 4), Grid1.TextMatrix(i + 1, 5))
 If Grid1.CellBackColor = 0 Then
 Grid1.CellBackColor = 1
 End If
 Grid1.Col = 6
 Grid1.Text = RGBtoHSL(Red, Green, Blue).Hue
 Grid1.Col = 7
 Grid1.Text = RGBtoHSL(Red, Green, Blue).Saturation
 Grid1.Col = 8
 Grid1.Text = RGBtoHSL(Red, Green, Blue).Luminance
Next
Grid1.Row = 1
Grid1.Col = 3
Call Set_UnDo
frmMain.Caption = "Populous Palette Editor - " + PaletteFileName
Save.Enabled = True
Call Set_UnDo
End If
Exit Sub
errorhandler:
On Error GoTo 0
Msgval = MsgBox("Error reading " & open1.FileName, 48, "Populous Palette Editor.")
Close
End Sub

Public Sub Saveas_Click()
open1.Filter = "PopulousPaletteFiles (PAL*.DAT)|PAL*.DAT"
open1.FileName = ""
open1.ShowSave
If open1.FileName <> "" Then
 PaletteFileName = open1.FileName
 frmMain.Caption = "Populous Palette Editor - " + PaletteFileName
 Call Save_Click
 Grid1.Row = 1
 Grid1.Col = 3
End If
End Sub

Private Sub Undo_Click()
Const L65536 As Long = 65536
Const L256 As Long = 256
Grid1.Row = UndoRow
Grid1.Col = 2
Grid1.CellBackColor = UndoColour
Blue = UndoColour \ L65536
Green = (UndoColour - (Blue * L65536)) \ L256
Red = (UndoColour - (Blue * L65536) - (Green * L256))
Grid1.TextMatrix(Grid1.Row, 3) = Red
Grid1.TextMatrix(Grid1.Row, 4) = Green
Grid1.TextMatrix(Grid1.Row, 5) = Blue
Grid1.TextMatrix(Grid1.Row, 6) = RGBtoHSL(Red, Green, Blue).Hue
Grid1.TextMatrix(Grid1.Row, 7) = RGBtoHSL(Red, Green, Blue).Saturation
Grid1.TextMatrix(Grid1.Row, 8) = RGBtoHSL(Red, Green, Blue).Luminance
Grid1.Col = 3
End Sub
Public Sub Set_UnDo()
Dim savecol As Integer
If Grid1.Row <> UndoRow Then
  savecol = Grid1.Col
  Grid1.Col = 2
  UndoColour = Grid1.CellBackColor
  Grid1.Col = savecol
  UndoRow = Grid1.Row
  undo.Enabled = True
End If
End Sub

