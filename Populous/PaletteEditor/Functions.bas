Attribute VB_Name = "Functions"
Public poppal As String

Public Type RGB
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Public Type HSL
    Hue As Integer
    Saturation As Integer
    Luminance As Integer
End Type

Public Function HSLtoRGB(ByVal Hue As Integer, _
                         ByVal Saturation As Integer, _
                         ByVal Luminance As Integer) As RGB

    Dim pHue As Single
    Dim pSat As Single
    Dim pLum As Single
    Dim RetVal As RGB
    Dim pRed As Single
    Dim pGreen As Single
    Dim pBlue As Single
    Dim temp2 As Single
    Dim temp3() As Single
    Dim temp1 As Single
    Dim n As Integer

   ReDim temp3(0 To 2)
   
   pHue = Hue / 239
   pSat = Saturation / 239
   pLum = Luminance / 239

   If pSat = 0 Then
      pRed = pLum!
      pGreen = pLum
      pBlue = pLum
   Else
      If pLum < 0.5 Then
         temp2 = pLum * (1 + pSat)
      Else
         temp2 = pLum + pSat - pLum * pSat
      End If
      temp1! = 2 * pLum! - temp2!
   
      temp3(0) = pHue + 1 / 3
      temp3(1) = pHue
      temp3(2) = pHue - 1 / 3
      
      For n = 0 To 2
         If temp3(n) < 0 Then temp3(n) = temp3(n) + 1
         If temp3(n) > 1 Then temp3(n) = temp3(n) - 1
      
         If 6 * temp3(n) < 1 Then
            temp3(n) = temp1 + (temp2 - temp1) * 6 * temp3(n)
         Else
            If 2 * temp3(n) < 1 Then
               temp3(n) = temp2
            Else
               If 3 * temp3(n%) < 2 Then
                  temp3(n%) = temp1 + (temp2 - temp1) _
                        * ((2 / 3) - temp3(n%)) * 6
               Else
                  temp3(n%) = temp1
                End If
             End If
          End If
       Next n%

       pRed = temp3(0)
       pGreen = temp3(1)
       pBlue = temp3(2)
    End If

    RetVal.Red = Int(pRed * 255)
    RetVal.Green = Int(pGreen * 255)
    RetVal.Blue = Int(pBlue * 255)
    
    HSLtoRGB = RetVal
End Function


Public Function RGBtoHSL(ByVal Red As Integer, _
                         ByVal Green As Integer, _
                         ByVal Blue As Integer) As HSL

    Dim pRed As Single
    Dim pGreen As Single
    Dim pBlue As Single
    Dim RetVal As HSL
    Dim pMax As Single
    Dim pMin As Single
    Dim pLum As Single
    Dim pSat As Single
    Dim pHue As Single
    
    pRed = Red / 255
    pGreen = Green / 255
    pBlue = Blue / 255
   
    If pRed > pGreen Then
       If pRed > pBlue Then
          pMax = pRed
       Else
          pMax = pBlue
       End If
    ElseIf pGreen > pBlue Then
        pMax = pGreen
    Else
        pMax = pBlue
    End If

    If pRed < pGreen Then
        If pRed < pBlue Then
            pMin = pRed
        Else
            pMin = pBlue
        End If
    ElseIf pGreen < pBlue Then
        pMin = pGreen
    Else
        pMin = pBlue
    End If

    pLum = (pMax + pMin) / 2
   
    If pMax = pMin Then
        pSat = 0
        pHue = 0
    Else
        If pLum < 0.5 Then
            pSat = (pMax - pMin) / (pMax + pMin)
        Else
            pSat = (pMax - pMin) / (2 - pMax - pMin)
        End If
        
        Select Case pMax!
        Case pRed
            pHue = (pGreen - pBlue) / (pMax - pMin)
        Case pGreen
            pHue = 2 + (pBlue - pRed) / (pMax - pMin)
        Case pBlue
            pHue = 4 + (pRed - pGreen) / (pMax - pMin)
        End Select
    End If

    RetVal.Hue = pHue * 239 \ 6
    If RetVal.Hue < 0 Then RetVal.Hue = RetVal.Hue + 240
    
    RetVal.Saturation = Int(pSat * 239)
    RetVal.Luminance = Int(pLum * 239)
    
    RGBtoHSL = RetVal
End Function



