VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Replace"
   ClientHeight    =   2040
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox To_Row 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Text            =   "255"
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox From_Row 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Text            =   "0"
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox With_Lum 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   11
      Text            =   "?"
      Top             =   1200
      Width           =   500
   End
   Begin VB.TextBox With_Sat 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   10
      Text            =   "?"
      Top             =   1200
      Width           =   500
   End
   Begin VB.TextBox With_Hue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Text            =   "?"
      Top             =   1200
      Width           =   500
   End
   Begin VB.TextBox With_Blue 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Text            =   "?"
      Top             =   1200
      Width           =   500
   End
   Begin VB.TextBox With_Green 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "?"
      Top             =   1200
      Width           =   500
   End
   Begin VB.TextBox With_Red 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "?"
      Top             =   1200
      Width           =   500
   End
   Begin VB.TextBox Replace_Lum 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Text            =   "?"
      Top             =   840
      Width           =   500
   End
   Begin VB.TextBox Replace_Sat 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Text            =   "?"
      Top             =   840
      Width           =   500
   End
   Begin VB.TextBox Replace_Hue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Text            =   "?"
      Top             =   840
      Width           =   500
   End
   Begin VB.TextBox Replace_Blue 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "?"
      Top             =   840
      Width           =   500
   End
   Begin VB.TextBox Replace_Green 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "?"
      Top             =   840
      Width           =   500
   End
   Begin VB.TextBox Replace_Red 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "?"
      Top             =   840
      Width           =   500
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Replace All"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame RGLorHSL 
      Caption         =   "Select RGB or HSL Replace"
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton Option_HSL 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option_RGB 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Lum."
         Height          =   255
         Left            =   4440
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Sat."
         Height          =   255
         Left            =   3840
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Hue"
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Blue"
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Green"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Red"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Use ? as a wildcard"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "To Row:"
      Height          =   255
      Left            =   1860
      TabIndex        =   19
      Top             =   1695
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "From Row:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1695
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Find What:"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Replace With:"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Replacements As Integer

Private Sub CancelButton_Click()
frmReplace.Hide
End Sub

Private Sub OKButton_Click()
Dim Msgval, i As Integer
With frmMain!Grid1
Replacements = 0
For i = From_Row To To_Row
.Row = i + 1
If Option_RGB(0).Value = True Then
  If Replace_Red = "?" Or .TextMatrix(.Row, 3) = Replace_Red Then
     If Replace_Green = "?" Or .TextMatrix(.Row, 4) = Replace_Green Then
       If Replace_Blue = "?" Or .TextMatrix(.Row, 5) = Replace_Blue Then
         'Row matches Replace template, so replace with clours which are not wild
         Replacements = Replacements + 1
         If With_Red <> "?" Then
           .TextMatrix(.Row, 3) = With_Red
         End If
         If With_Green <> "?" Then
           .TextMatrix(.Row, 4) = With_Green
         End If
         If With_Blue <> "?" Then
           .TextMatrix(.Row, 5) = With_Blue
         End If
         .TextMatrix(.Row, 6) = RGBtoHSL(.TextMatrix(.Row, 3), .TextMatrix(.Row, 4), .TextMatrix(.Row, 5)).Hue
         .TextMatrix(.Row, 7) = RGBtoHSL(.TextMatrix(.Row, 3), .TextMatrix(.Row, 4), .TextMatrix(.Row, 5)).Saturation
         .TextMatrix(.Row, 8) = RGBtoHSL(.TextMatrix(.Row, 3), .TextMatrix(.Row, 4), .TextMatrix(.Row, 5)).Luminance
         Call frmMain.Colour_Changed
       End If
     End If
  End If
Else
  If Replace_Hue = "?" Or .TextMatrix(.Row, 6) = Replace_Hue Then
    If Replace_Sat = "?" Or .TextMatrix(.Row, 7) = Replace_Sat Then
      If Replace_Lum = "?" Or .TextMatrix(.Row, 8) = Replace_Lum Then
        'Row matches Replace template, so replace with clours which are not wild
        Replacements = Replacements + 1
        If With_Hue <> "?" Then
          .TextMatrix(.Row, 6) = With_Hue
        End If
        If With_Sat <> "?" Then
          .TextMatrix(.Row, 7) = With_Sat
        End If
        If With_Lum <> "?" Then
          .TextMatrix(.Row, 8) = With_Lum
        End If
        .TextMatrix(.Row, 3) = HSLtoRGB(.TextMatrix(.Row, 6), .TextMatrix(.Row, 7), .TextMatrix(.Row, 8)).Red
        .TextMatrix(.Row, 4) = HSLtoRGB(.TextMatrix(.Row, 6), .TextMatrix(.Row, 7), .TextMatrix(.Row, 8)).Green
        .TextMatrix(.Row, 5) = HSLtoRGB(.TextMatrix(.Row, 6), .TextMatrix(.Row, 7), .TextMatrix(.Row, 8)).Blue
        Call frmMain.Colour_Changed
      End If
    End If
  End If
End If
Next
If Replacements = 0 Then
Msgval = MsgBox("No matching colours found in selected rows", 64, "Populous Palette Editor")
Else
Msgval = MsgBox(Replacements & " colours replaced", 64, "Populous Palette Editor")
End If
Call frmMain.Set_UnDo
End With
End Sub

Private Sub Option_RGB_Click(Index As Integer)
Replace_Red.Enabled = True
Replace_Green.Enabled = True
Replace_Blue.Enabled = True
With_Red.Enabled = True
With_Green.Enabled = True
With_Blue.Enabled = True
Replace_Hue.Enabled = False
Replace_Sat.Enabled = False
Replace_Lum.Enabled = False
With_Hue.Enabled = False
With_Sat.Enabled = False
With_Lum.Enabled = False
End Sub
Private Sub Option_HSL_Click(Index As Integer)
Replace_Red.Enabled = False
Replace_Green.Enabled = False
Replace_Blue.Enabled = False
With_Red.Enabled = False
With_Green.Enabled = False
With_Blue.Enabled = False
Replace_Hue.Enabled = True
Replace_Sat.Enabled = True
Replace_Lum.Enabled = True
With_Hue.Enabled = True
With_Sat.Enabled = True
With_Lum.Enabled = True
End Sub


Private Sub Replace_Red_LostFocus()
If IsNumeric(Replace_Red) Then
  If Replace_Red >= 0 And Replace_Red <= 255 Then
  Else
    Replace_Red = "255"
  End If
Else
  Replace_Red = "?"
End If
End Sub
Private Sub Replace_Green_LostFocus()
If IsNumeric(Replace_Green) Then
  If Replace_Green >= 0 And Replace_Green <= 255 Then
  Else
    Replace_Green = "255"
  End If
Else
  Replace_Green = "?"
End If
End Sub

Private Sub Replace_Blue_LostFocus()
If IsNumeric(Replace_Blue) Then
  If Replace_Blue >= 0 And Replace_Blue <= 255 Then
  Else
    Replace_Blue = "255"
  End If
Else
  Replace_Blue = "?"
End If
End Sub
Private Sub Replace_Hue_LostFocus()
If IsNumeric(Replace_Hue) Then
  If Replace_Hue >= 0 And Replace_Hue <= 240 Then
  Else
    Replace_Hue = "240"
  End If
Else
  Replace_Hue = "?"
End If
End Sub

Private Sub Replace_Sat_LostFocus()
If IsNumeric(Replace_Sat) Then
  If Replace_Sat >= 0 And Replace_Sat <= 239 Then
  Else
    Replace_Sat = "239"
  End If
Else
  Replace_Sat = "?"
End If
End Sub

Private Sub Replace_Lum_LostFocus()
If IsNumeric(Replace_Lum) Then
  If Replace_Lum >= 0 And Replace_Lum <= 239 Then
  Else
    Replace_Lum = "239"
  End If
Else
  Replace_Lum = "?"
End If
End Sub

Private Sub With_Red_LostFocus()
If IsNumeric(With_Red) Then
  If With_Red >= 0 And With_Red <= 255 Then
  Else
    With_Red = "255"
  End If
Else
  With_Red = "?"
End If
End Sub
Private Sub With_Green_LostFocus()
If IsNumeric(With_Green) Then
  If With_Green >= 0 And With_Green <= 255 Then
  Else
    With_Green = "255"
  End If
Else
  With_Green = "?"
End If
End Sub

Private Sub With_Blue_LostFocus()
If IsNumeric(With_Blue) Then
  If With_Blue >= 0 And With_Blue <= 255 Then
  Else
    With_Blue = "255"
  End If
Else
  With_Blue = "?"
End If
End Sub
Private Sub With_Hue_LostFocus()
If IsNumeric(With_Hue) Then
  If With_Hue >= 0 And With_Hue <= 240 Then
  Else
    With_Hue = "240"
  End If
Else
  With_Hue = "?"
End If
End Sub

Private Sub With_Sat_LostFocus()
If IsNumeric(With_Sat) Then
  If With_Sat >= 0 And With_Sat <= 239 Then
  Else
    With_Sat = "239"
  End If
Else
  With_Sat = "?"
End If
End Sub

Private Sub With_Lum_LostFocus()
If IsNumeric(With_Lum) Then
  If With_Lum >= 0 And With_Lum <= 239 Then
  Else
    With_Lum = "239"
  End If
Else
  With_Lum = "?"
End If
End Sub
Private Sub From_Row_LostFocus()
If IsNumeric(From_Row) Then
  If From_Row >= 0 And From_Row <= 255 Then
  Else
    From_Row = 0
  End If
Else
  From_Row = 0
End If
If Int(To_Row) < Int(From_Row) Then
  To_Row = From_Row
End If
End Sub
Private Sub To_Row_LostFocus()
If IsNumeric(To_Row) Then
  If To_Row >= 0 And To_Row <= 255 Then
  Else
    To_Row = 255
  End If
Else
  To_Row = 255
End If
If Int(To_Row) < Int(From_Row) Then
  To_Row = From_Row
End If
End Sub

Private Sub Replace_Red_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call Replace_Red_LostFocus
End If
End Sub
Private Sub Replace_Green_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call Replace_Green_LostFocus
End If
End Sub
Private Sub Replace_Blue_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call Replace_Blue_LostFocus
End If
End Sub
Private Sub Replace_Hue_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call Replace_Hue_LostFocus
End If
End Sub
Private Sub Replace_Sat_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call Replace_Sat_LostFocus
End If
End Sub
Private Sub Replace_Lum_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call Replace_Lum_LostFocus
End If
End Sub
Private Sub With_Red_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call With_Red_LostFocus
End If
End Sub
Private Sub With_Green_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call With_Green_LostFocus
End If
End Sub
Private Sub With_Blue_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call With_Blue_LostFocus
End If
End Sub
Private Sub With_Hue_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call With_Hue_LostFocus
End If
End Sub
Private Sub With_Sat_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call With_Sat_LostFocus
End If
End Sub

Private Sub With_Lum_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call With_Lum_LostFocus
End If
End Sub
Private Sub From_Row_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call From_Row_LostFocus
End If
End Sub
Private Sub To_Row_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call To_Row_LostFocus
End If
End Sub

