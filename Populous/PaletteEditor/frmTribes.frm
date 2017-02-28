VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{D632D0BF-C2F7-4C7B-B58B-6CCCE74622A4}#1.0#0"; "ColourControl.ocx"
Begin VB.Form frmTribes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Tribe Colours"
   ClientHeight    =   4200
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6015
   Begin VB.Frame Frame4 
      Caption         =   "Change Colour"
      Height          =   2175
      Left            =   3120
      TabIndex        =   25
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtRed 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "255"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtGreen 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   29
         Text            =   "255"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtBlue 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "255"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Custom Colours"
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   2505
      End
      Begin Project1.ControlColour imgColour 
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1296
      End
      Begin VB.Label lblColour 
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   32
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblColour 
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   31
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Original Populous Colours"
      Height          =   615
      Left            =   3120
      TabIndex        =   19
      Top             =   2400
      Width           =   2775
      Begin Project1.ControlColour imgOriginalColour 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   13319987
      End
      Begin Project1.ControlColour imgOriginalColour 
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   5027
      End
      Begin Project1.ControlColour imgOriginalColour 
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   1275811
      End
      Begin Project1.ControlColour imgOriginalColour 
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   6531887
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colour Palette"
      Height          =   975
      Left            =   3120
      TabIndex        =   6
      Top             =   3120
      Width           =   2775
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   7864320
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   16711680
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   16776960
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   65280
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   30720
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   26265
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   65535
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   14
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   37119
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   15
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   255
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   16
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   7864440
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   17
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   0
      End
      Begin Project1.ControlColour imgBasicColour 
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   18
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colour          =   16777215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Preview Image"
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2895
      Begin VB.PictureBox picSample 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1800
         Left            =   120
         Picture         =   "frmTribes.frx":0000
         ScaleHeight     =   120
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   176
         TabIndex        =   5
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   6960
      Picture         =   "frmTribes.frx":F7C2
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   176
      TabIndex        =   3
      Top             =   360
      Width           =   2640
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "Blue"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   6960
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape shpShade 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpShade 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   480
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpShade 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   840
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpShade 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   1200
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpShade 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   1560
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpShade 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   1920
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpShade 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   2280
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpShade 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   2640
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblShadesEvent 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "frmTribes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Private TempColour As Long
Private Tribes(3, 7, 2) As Byte
Private Tribecolours(4, 8) As Long
Private TribeColour As Integer
Private SuspendChange As Boolean
Private rgbShade(7) As RGB
Private hslShade(7) As HSL
Private colourid As Byte
Private TribesColour(3) As Long
Private DoNotChange As Boolean

Function LimitTextInput(Source) As String 'prevemts anything but integers from being entered into the script text boxes.
    Const Numbers$ = "0123456789"
    'backspace =8
    If Source <> 8 Then
      If InStr(Numbers, Chr(Source)) = 0 Then
            LimitTextInput = 0
            Exit Function
      End If
    End If
    LimitTextInput = Source
End Function

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cmdColour_Click()
If cmdColour.Caption = "Blue" Then
  cmdColour.Caption = "Red"
  colourid = 1
ElseIf cmdColour.Caption = "Red" Then
  cmdColour.Caption = "Yellow"
  colourid = 2
ElseIf cmdColour.Caption = "Yellow" Then
  cmdColour.Caption = "Green"
  colourid = 3
ElseIf cmdColour.Caption = "Green" Then
  cmdColour.Caption = "Blue"
  colourid = 0
End If

Dim i As Byte
For i = 0 To 7
  shpShade(i).BackColor = RGB(Tribes(colourid, i, 0), Tribes(colourid, i, 1), Tribes(colourid, i, 2))
Next

imgColour.colour = TribesColour(colourid)

DoNotChange = True
txtBlue.Text = imgColour.colour \ 65536
txtGreen.Text = (imgColour.colour \ 256) Mod 256
txtRed.Text = imgColour.colour Mod 256
DoNotChange = False

'Update image with shades
For Y = 0 To picSample.ScaleHeight
  For X = 0 To picSample.ScaleWidth
    PixelColour = GetPixel(picOriginal.hDC, X, Y)

    For z1 = 0 To 7
      For z2 = 0 To 3
        If Tribecolours(z2, z1) = PixelColour Then
          SetPixel picSample.hDC, X, Y, shpShade(z1).BackColor
        End If
      Next
    Next
  Next
Next
picSample.Refresh
End Sub

Private Sub Command2_Click()
'On Error GoTo cmdColorError
'    Dim c As New cCommonDialog
'    With c
'        .DialogTitle = "Choose a Color"
'        .Flags = CC_AnyColor Or CC_FullOpen
'        .CancelError = True
'        .hWnd = Me.hWnd
'        .Color = picColor.BackColor
'        .ShowColor
'        picColor.BackColor = .Color
'
'        pShowColorNumber
'    End With
'    Exit Sub
'cmdColorError:
'    If (Err.Number <> 20001) Then
'        MsgBox "Error: " & Err.Description
'    End If

SuspendChange = True
On Error GoTo ignorit
  With cmd
  .DialogTitle = "Custom Color"
  .ShowColor
  imgColour.colour = .Color
  End With
txtBlue.Text = imgColour.colour \ 65536
txtGreen.Text = (imgColour.colour \ 256) Mod 256
txtRed.Text = imgColour.colour Mod 256
TribeColour = 4
Call Calculate_Shade
ignorit:
On Error GoTo 0
SuspendChange = False
End Sub



Private Sub Form_Load()
Tribecolours(0, 0) = RGB(16, 33, 66) '(15, 15, 59)
Tribecolours(0, 1) = RGB(23, 35, 67)
Tribecolours(0, 2) = RGB(33, 49, 107) '(35, 51, 111)
Tribecolours(0, 3) = RGB(41, 57, 156) '(43, 59, 155)
Tribecolours(0, 4) = RGB(49, 57, 206) '(51, 63, 203)
Tribecolours(0, 5) = RGB(74, 82, 214) '(75, 83, 211)
Tribecolours(0, 6) = RGB(99, 107, 222) '(103, 111, 223)
Tribecolours(0, 7) = RGB(132, 140, 239) '(135, 139, 235)
Tribecolours(1, 0) = RGB(24, 16, 8) '(19, 7, 7)
Tribecolours(1, 1) = RGB(57, 0, 0) '(63, 7, 7)
Tribecolours(1, 2) = RGB(90, 0, 0) '(95, 7, 7)
Tribecolours(1, 3) = RGB(123, 0, 0) '(127, 7, 7)
Tribecolours(1, 4) = RGB(165, 16, 0) '(163, 19, 0)
Tribecolours(1, 5) = RGB(181, 57, 24) '(183, 59, 31)
Tribecolours(1, 6) = RGB(198, 115, 74) '(199, 115, 75)
Tribecolours(1, 7) = RGB(239, 173, 123) '(239, 171, 127)
Tribecolours(2, 0) = RGB(57, 0, 0) '(51, 23, 23)
Tribecolours(2, 1) = RGB(74, 33, 24) '(79, 35, 27)
Tribecolours(2, 2) = RGB(107, 49, 42) '(107, 55, 31)
Tribecolours(2, 3) = RGB(132, 82, 24) '(135, 83, 27)
Tribecolours(2, 4) = RGB(165, 115, 16) '(163, 119, 19)
Tribecolours(2, 5) = RGB(189, 148, 33) '(191, 147, 39)
Tribecolours(2, 6) = RGB(222, 181, 57) '(219, 179, 63)
Tribecolours(2, 7) = RGB(255, 214, 90) '(251, 215, 95)
Tribecolours(3, 0) = RGB(0, 33, 16) '(7, 39, 19)
Tribecolours(3, 1) = RGB(16, 99, 57) '(15, 71, 43)
Tribecolours(3, 2) = RGB(23, 103, 63)
Tribecolours(3, 3) = RGB(33, 140, 74) '(35, 139, 79)
Tribecolours(3, 4) = RGB(41, 173, 99) '(47, 171, 99)
Tribecolours(3, 5) = RGB(57, 206, 115) '(63, 207, 119)
Tribecolours(3, 6) = RGB(99, 222, 156) '(99, 223, 159)
Tribecolours(3, 7) = RGB(148, 247, 206) '(147, 243, 203)
Call imgOriginalColour_Click(0)

  Dim Count As Byte
  Count = 0
  For i = 216 To 223
   Tribes(0, Count, 0) = frmMain!Grid1.TextMatrix(i + 1, 3)
   Tribes(0, Count, 1) = frmMain!Grid1.TextMatrix(i + 1, 4)
   Tribes(0, Count, 2) = frmMain!Grid1.TextMatrix(i + 1, 5)
   If Count = 4 Then
     TribesColour(0) = RGB(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5))
   End If
   Count = Count + 1
  Next
  Count = 0
  For i = 240 To 247
   Tribes(1, Count, 0) = frmMain!Grid1.TextMatrix(i + 1, 3)
   Tribes(1, Count, 1) = frmMain!Grid1.TextMatrix(i + 1, 4)
   Tribes(1, Count, 2) = frmMain!Grid1.TextMatrix(i + 1, 5)
   If Count = 4 Then
     TribesColour(1) = RGB(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5))
   End If
   Count = Count + 1
  Next
  Count = 0
  For i = 232 To 239
   Tribes(2, Count, 0) = frmMain!Grid1.TextMatrix(i + 1, 3)
   Tribes(2, Count, 1) = frmMain!Grid1.TextMatrix(i + 1, 4)
   Tribes(2, Count, 2) = frmMain!Grid1.TextMatrix(i + 1, 5)
   If Count = 4 Then
     TribesColour(2) = RGB(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5))
   End If
   Count = Count + 1
  Next
  Count = 0
  For i = 224 To 231
   Tribes(3, Count, 0) = frmMain!Grid1.TextMatrix(i + 1, 3)
   Tribes(3, Count, 1) = frmMain!Grid1.TextMatrix(i + 1, 4)
   Tribes(3, Count, 2) = frmMain!Grid1.TextMatrix(i + 1, 5)
   If Count = 4 Then
     TribesColour(3) = RGB(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5))
   End If
   Count = Count + 1
  Next

For i = 0 To 7
  shpShade(i).BackColor = RGB(Tribes(colourid, i, 0), Tribes(colourid, i, 1), Tribes(colourid, i, 2))
Next

imgColour.colour = TribesColour(colourid)

'Update image with shades
For Y = 0 To picSample.ScaleHeight
  For X = 0 To picSample.ScaleWidth
    PixelColour = GetPixel(picOriginal.hDC, X, Y)

    For z1 = 0 To 7
      For z2 = 0 To 3
        If Tribecolours(z2, z1) = PixelColour Then
          SetPixel picSample.hDC, X, Y, shpShade(z1).BackColor
        End If
      Next
    Next
  Next
Next
picSample.Refresh
End Sub



Private Sub imgBasicColour_Click(Index As Integer)
SuspendChange = True
imgColour.colour = imgBasicColour(Index).colour

txtBlue.Text = imgColour.colour \ 65536
txtGreen.Text = (imgColour.colour \ 256) Mod 256
txtRed.Text = imgColour.colour Mod 256
TribeColour = 4
Call Calculate_Shade
SuspendChange = False
End Sub

Private Sub imgOriginalColour_Click(Index As Integer)
SuspendChange = True
imgColour.colour = imgOriginalColour(Index).colour

txtBlue.Text = imgColour.colour \ 65536
txtGreen.Text = (imgColour.colour \ 256) Mod 256
txtRed.Text = imgColour.colour Mod 256
TribeColour = Index
Call Calculate_Shade
SuspendChange = False
End Sub

Private Sub OKButton_Click()
  Dim Count As Byte
  Count = 0
  For i = 216 To 223
   frmMain!Grid1.TextMatrix(i + 1, 3) = Tribes(0, Count, 0)
   frmMain!Grid1.TextMatrix(i + 1, 4) = Tribes(0, Count, 1)
   frmMain!Grid1.TextMatrix(i + 1, 5) = Tribes(0, Count, 2)
   Count = Count + 1
  Next
  Count = 0
  For i = 240 To 247
   frmMain!Grid1.TextMatrix(i + 1, 3) = Tribes(1, Count, 0)
   frmMain!Grid1.TextMatrix(i + 1, 4) = Tribes(1, Count, 1)
   frmMain!Grid1.TextMatrix(i + 1, 5) = Tribes(1, Count, 2)
   Count = Count + 1
  Next
  Count = 0
  For i = 232 To 239
   frmMain!Grid1.TextMatrix(i + 1, 3) = Tribes(2, Count, 0)
   frmMain!Grid1.TextMatrix(i + 1, 4) = Tribes(2, Count, 1)
   frmMain!Grid1.TextMatrix(i + 1, 5) = Tribes(2, Count, 2)
   Count = Count + 1
  Next
  Count = 0
  For i = 224 To 231
   frmMain!Grid1.TextMatrix(i + 1, 3) = Tribes(3, Count, 0)
   frmMain!Grid1.TextMatrix(i + 1, 4) = Tribes(3, Count, 1)
   frmMain!Grid1.TextMatrix(i + 1, 5) = Tribes(3, Count, 2)
   Count = Count + 1
  Next

  For i = 216 To 247
     frmMain!Grid1.TextMatrix(i + 1, 6) = RGBtoHSL(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5)).Hue
     frmMain!Grid1.TextMatrix(i + 1, 7) = RGBtoHSL(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5)).Saturation
     frmMain!Grid1.TextMatrix(i + 1, 8) = RGBtoHSL(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5)).Luminance
  Next

  For i = 216 To 247
     frmMain!Grid1.Row = i + 1
     frmMain!Grid1.Col = 2
     If RGB(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5)) = 0 Then
       frmMain!Grid1.CellBackColor = 1
     Else
       frmMain!Grid1.CellBackColor = RGB(frmMain!Grid1.TextMatrix(i + 1, 3), frmMain!Grid1.TextMatrix(i + 1, 4), frmMain!Grid1.TextMatrix(i + 1, 5))
     End If
  Next
  Unload Me
End Sub

Private Sub txtBlue_Change()
If DoNotChange = True Then Exit Sub
'The SuspendChange flag prevents this text field from being changed recursively
If SuspendChange Then
Else
SuspendChange = True
    'Check Values
If txtBlue.Text = "" Then
  txtBlue.Text = "0"
End If
If CLng(txtBlue.Text) > 255 Then
  txtBlue.Text = "255"
End If
txtBlue.Text = CLng(txtBlue.Text)
imgColour.colour = RGB(CInt(txtRed.Text), CInt(txtGreen.Text), CInt(txtBlue.Text))

TribeColour = 4
Call Calculate_Shade
SuspendChange = False
End If
End Sub

Private Sub txtBlue_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii)
    If KeyAscii <> 8 Then
        If LimitTextInput(KeyAscii) = 0 Then
        
        End If
    End If

End Sub

Private Sub txtGreen_Change()
If DoNotChange = True Then Exit Sub
'The SuspendChange flag prevents this text field from being changed recursively
If SuspendChange Then
Else
SuspendChange = True
'Check Values
If txtGreen.Text = "" Then
  txtGreen.Text = "0"
End If
If CLng(txtGreen.Text) > 255 Then
  txtGreen.Text = "255"
End If
txtGreen.Text = CLng(txtGreen.Text)
imgColour.colour = RGB(CInt(txtRed.Text), CInt(txtGreen.Text), CInt(txtBlue.Text))

TribeColour = 4
Call Calculate_Shade
SuspendChange = False
End If
End Sub

Private Sub txtGreen_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii)
    If KeyAscii <> 8 Then
        If LimitTextInput(KeyAscii) = 0 Then
        
        End If
    End If
End Sub

Private Sub txtRed_Change()
If DoNotChange = True Then Exit Sub
'The SuspendChange flag prevents this text field from being changed recursively
If SuspendChange Then
Else
SuspendChange = True
    'Check Values
If txtRed.Text = "" Then
  txtRed.Text = "0"
End If
If CLng(txtRed.Text) > 255 Then
  txtRed.Text = "255"
End If
txtRed.Text = CLng(txtRed.Text)
imgColour.colour = RGB(CInt(txtRed.Text), CInt(txtGreen.Text), CInt(txtBlue.Text))

TribeColour = 4
Call Calculate_Shade
SuspendChange = False
End If
End Sub

Private Sub txtRed_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii)
    If KeyAscii <> 8 Then
        If LimitTextInput(KeyAscii) = 0 Then
        
        End If
    End If
End Sub

Public Sub Calculate_Shade()
Dim PixelColour As Long
Dim i As Integer, X As Integer, Y As Integer, z1 As Integer, z2 As Integer
For i = 0 To 7
  hslShade(i).Hue = RGBtoHSL(txtRed.Text, txtGreen.Text, txtBlue.Text).Hue
  hslShade(i).Saturation = RGBtoHSL(txtRed.Text, txtGreen.Text, txtBlue.Text).Saturation
Next
Select Case RGBtoHSL(txtRed.Text, txtGreen.Text, txtBlue.Text).Luminance
Case 0 'Black
For i = 1 To 7
  hslShade(i).Luminance = 0
Next
Case Is < 80 'Low Luminance colour
  hslShade(0).Luminance = 0
  hslShade(1).Luminance = 20
  hslShade(2).Luminance = 40
  hslShade(3).Luminance = 60
  hslShade(4).Luminance = 80
  hslShade(5).Luminance = 100
  hslShade(6).Luminance = 120
  hslShade(7).Luminance = 140
Case Is < 160 'Medium Luminance colour
  hslShade(0).Luminance = 20
  hslShade(1).Luminance = 40
  hslShade(2).Luminance = 60
  hslShade(3).Luminance = 80
  hslShade(4).Luminance = 100
  hslShade(5).Luminance = 120
  hslShade(6).Luminance = 140
  hslShade(7).Luminance = 170
Case Else 'High luminance colour
  hslShade(0).Luminance = 100
  hslShade(1).Luminance = 120
  hslShade(2).Luminance = 140
  hslShade(3).Luminance = 160
  hslShade(4).Luminance = 180
  hslShade(5).Luminance = 200
  hslShade(6).Luminance = 220
  hslShade(7).Luminance = 239
End Select
For i = 0 To 7
  rgbShade(i).Red = HSLtoRGB(hslShade(i).Hue, hslShade(i).Saturation, hslShade(i).Luminance).Red
  rgbShade(i).Green = HSLtoRGB(hslShade(i).Hue, hslShade(i).Saturation, hslShade(i).Luminance).Green
  rgbShade(i).Blue = HSLtoRGB(hslShade(i).Hue, hslShade(i).Saturation, hslShade(i).Luminance).Blue
'Display
Select Case TribeColour
Case 0
  shpShade(i).BackColor = Tribecolours(0, i)
Case 1
  shpShade(i).BackColor = Tribecolours(1, i)
Case 2
  shpShade(i).BackColor = Tribecolours(2, i)
Case 3
  shpShade(i).BackColor = Tribecolours(3, i)
Case Else
  shpShade(i).BackColor = RGB(rgbShade(i).Red, rgbShade(i).Green, rgbShade(i).Blue)
End Select
Next

'Update image with shades
For Y = 0 To picSample.ScaleHeight
  For X = 0 To picSample.ScaleWidth
    PixelColour = GetPixel(picOriginal.hDC, X, Y)

    For z1 = 0 To 7
      For z2 = 0 To 3
        If Tribecolours(z2, z1) = PixelColour Then
          SetPixel picSample.hDC, X, Y, shpShade(z1).BackColor
        End If
      Next
    Next
  Next
Next
picSample.Refresh

TribesColour(colourid) = imgColour.colour

For X = 0 To 7
  Tribes(colourid, X, 0) = rgbShade(X).Red
  Tribes(colourid, X, 1) = rgbShade(X).Green
  Tribes(colourid, X, 2) = rgbShade(X).Blue
Next
End Sub

