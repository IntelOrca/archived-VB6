VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMMPaletteChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright TedTycoon 2005, Must only be used under TedTycoon's Permission

Dim PalStatus As String
Dim ColourRed(255) As Integer
Dim ColourGreen(255) As Integer
Dim ColourBlue(255) As Integer

Public Sub Change_Colour(PalIndex As Integer, ColRed As Integer, ColGreen As Integer, ColBlue As Integer)
'Check Colours
If PalIndex > 255 Then
  PalIndex = 255
End If
If ColRed > 255 Then
  ColRed = 255
End If
If ColGreen > 255 Then
  ColGreen = 255
End If
If ColBlue > 255 Then
  ColBlue = 255
End If

'Change Colour
ColourRed(PalIndex) = ColRed
ColourGreen(PalIndex) = ColGreen
ColourBlue(PalIndex) = ColBlue

PalStatus = PalIndex & " = Red: " & ColRed & " Green: " & ColGreen & " Blue: " & ColBlue
End Sub

Public Sub Save_Palette(PalFile As String)
Dim poppal As String
On Error GoTo errorhandler
poppal = ""
For i = 0 To 255
  poppal = poppal + Chr$(ColourRed(i)) + Chr$(ColourGreen(i)) + Chr$(ColourBlue(i)) + Chr$(0)
Next
Open PalFile For Binary As #1
Put #1, 1, poppal
Close #1
PalStatus = "Saving Palette Complete!"
Exit Sub
errorhandler:
On Error GoTo 0
PalStatus = "Saving Palette Failed!"
Close
End Sub
