Attribute VB_Name = "modAirbrush"
Option Explicit

Public Const CSolid As Long = 0
Public Const CShift As Long = 1
Public Const MM_invert As Long = 2

Public Type AirbrushType
 diameter As Single
 definition As Single
 Pressure As Byte
 Red As Byte
 Green As Byte
 Blue As Byte
 DoOutline As Boolean
 outline_fullPressure_extension As Single
 outline_cushion_inside As Single
 DoErase As Boolean
 Mode_0_to_2 As Long
End Type

'ColorShift sub for when AirbrushType member Mode_0_to_2 = CShift
Public rgb_shift_intensity! ' ! means  As Single
Public sR!
Public sG!
Public sB!
Dim maximu!
Dim minimu!
Dim iSubt!
Dim bytMaxMin_diff As Byte

Public Airbrush As AirbrushType

Public Sub BlitAirbrush(Surf As AnimSurf2D, ByVal px!, ByVal py!)
Dim DrawX As Long
Dim DrawY As Long
Dim DrawRight As Long
Dim DrawLeft As Long
Dim DrawTop As Long
Dim DrawBot As Long
Dim AddDrawHeight As Long
Dim AddDrawWidth As Long
Dim brush_radius As Single
Dim brush_height As Single
Dim brush_slope As Single
Dim height_Sq As Single
Dim delta_ySq As Single
Dim delta_y As Single
Dim delta_x As Single
Dim delta_left As Single
Dim deltas_xy_Sq As Single
Dim BGRed As Long
Dim BGGrn As Long
Dim BGBlu As Long
Dim FGRed As Long
Dim FGGrn As Long
Dim FGBlu As Long
Dim first_cone_slice As Single
Dim first_slice_Sq As Single
Dim second_cone_slice As Single
Dim second_slice_Sq As Single
Dim inv_cone_tip As Single
Dim inv_tip_Sq As Single
Dim TmpPressure As Byte
Dim cushion_inside_multiplier As Single
Dim FGColor&
Dim BGColor&
Dim BGR&
Dim sngAlpha!
Dim sngAlpha2!
Dim sPress!

'New variables for rounding operations.  No longer use
'Realround()
Dim px_m As Single
Dim px_p As Single
Dim py_m As Single
Dim py_p As Single

If Surf.TotalPixels > 0 Then

 If Airbrush.diameter > 0 Then

 If Airbrush.Mode_0_to_2 = MM_invert Then
 TmpPressure = Airbrush.Pressure
 Airbrush.Pressure = 255
 End If
 
 sPress = Airbrush.Pressure / 255

 brush_height = sPress * Airbrush.definition
 brush_radius = Airbrush.diameter / 2
 
 FGColor = RGB(Airbrush.Blue, Airbrush.Green, Airbrush.Red)

 If Airbrush.DoOutline Then
  first_cone_slice = brush_height - sPress
  If first_cone_slice < 0 Then first_cone_slice = 0
  first_slice_Sq = first_cone_slice * first_cone_slice
  second_cone_slice = first_cone_slice - sPress * Airbrush.outline_fullPressure_extension
  If second_cone_slice < 0 Then second_cone_slice = 0
  second_slice_Sq = second_cone_slice * second_cone_slice
 
  If Airbrush.outline_cushion_inside <= 1 Then
   If Airbrush.outline_cushion_inside < 0 Then
    cushion_inside_multiplier = 1
   Else
    cushion_inside_multiplier = 1 - Airbrush.outline_cushion_inside
   End If
   inv_cone_tip = second_cone_slice - (Airbrush.Pressure / cushion_inside_multiplier) / 255
  Else
   cushion_inside_multiplier = Airbrush.outline_cushion_inside
   inv_cone_tip = second_cone_slice - sPress * cushion_inside_multiplier
  End If
  
  If inv_cone_tip < 0 Then inv_cone_tip = 0
  inv_tip_Sq = inv_cone_tip * inv_cone_tip
 End If

 brush_slope = brush_height / brush_radius
 height_Sq = brush_height * brush_height

 px_m = px - brush_radius
 py_m = py - brush_radius
 px_p = px + brush_radius
 py_p = py + brush_radius

 'Round() DrawLeft, DrawBot, DrawRight, DrawTop
 'VB's round is not helpful here
 DrawLeft = Int(px_m)
 DrawBot = Int(py_m)
 DrawRight = Int(px_p)
 DrawTop = Int(py_p)
 'If (px_m - DrawLeft) >= 0.5 Then DrawLeft = DrawLeft + 1&
 'If (py_m - DrawBot) >= 0.5 Then DrawBot = DrawBot + 1&
 'If (px_p - DrawRight) >= 0.5 Then DrawRight = DrawRight + 1&
 'If (py_p - DrawTop) >= 0.5 Then DrawTop = DrawTop + 1&

 If DrawLeft < 0 Then DrawLeft = 0
 If DrawBot < 0 Then DrawBot = 0
 If DrawRight > Surf.Dims.WideM1 Then DrawRight = Surf.Dims.WideM1
 If DrawTop > Surf.Dims.HighM1 Then DrawTop = Surf.Dims.HighM1

 delta_left = (DrawLeft - px) * brush_slope
 delta_y = (DrawBot - py) * brush_slope

 AddDrawWidth = DrawRight - DrawLeft
 AddDrawHeight = DrawTop - DrawBot

 DrawBot = DrawBot * Surf.Dims.Wide + DrawLeft
 DrawTop = DrawBot + Surf.Dims.Wide * AddDrawHeight
 
 If Airbrush.DoErase Then
  If Airbrush.Mode_0_to_2 = CSolid Then

  If Airbrush.DoOutline Then

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     If deltas_xy_Sq > first_slice_Sq Then
      sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     Else
      If deltas_xy_Sq < second_slice_Sq Then
       If deltas_xy_Sq >= inv_tip_Sq Then
        sngAlpha = sPress - (second_cone_slice - Sqr#(deltas_xy_Sq)) * cushion_inside_multiplier
      Else
        sngAlpha = 0
       End If
      Else
       sngAlpha = sPress
      End If
     End If
     If sngAlpha > 0 Then
      BGColor = Surf.LDib1D(DrawX)
      BGGrn = BGColor And vbGreen
      FGGrn = FGColor And vbGreen
      BGBlu = BGColor And &HFF&
      FGBlu = FGColor And &HFF&
      Surf.LDib1D(DrawX) = _
         BGColor + sngAlpha * (FGColor - BGColor) And vbBlue Or _
         BGGrn + sngAlpha * (FGGrn - BGGrn) And vbGreen Or _
         BGBlu + sngAlpha * (FGBlu - BGBlu) And &HFF&
     End If
    End If
    delta_x = delta_x + brush_slope
   Next
   delta_y = delta_y + brush_slope
  Next DrawY

  Else 'filled circle

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     If sngAlpha > sPress Then sngAlpha = sPress
     BGColor = Surf.LDib1D(DrawX)
     BGGrn = BGColor And vbGreen
     FGGrn = FGColor And vbGreen
     BGBlu = BGColor And &HFF&
     FGBlu = FGColor And &HFF&
     Surf.LDib1D(DrawX) = _
        BGColor + sngAlpha * (FGColor - BGColor) And vbBlue Or _
        BGGrn + sngAlpha * (FGGrn - BGGrn) And vbGreen Or _
        BGBlu + sngAlpha * (FGBlu - BGBlu) And &HFF&
    End If
    delta_x = delta_x + brush_slope
   Next
   delta_y = delta_y + brush_slope
  Next

  End If 'Airbrush.DoOutline

  ElseIf Airbrush.Mode_0_to_2 = MM_invert Then

  If Airbrush.DoOutline Then

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     If deltas_xy_Sq > first_slice_Sq Then
      sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     Else
      If deltas_xy_Sq < second_slice_Sq Then
       If deltas_xy_Sq >= inv_tip_Sq Then
        sngAlpha = sPress - (second_cone_slice - Sqr#(deltas_xy_Sq)) * cushion_inside_multiplier
       Else
        sngAlpha = 0
       End If
      Else
       sngAlpha = sPress
      End If
     End If
     If sngAlpha > 0 Then
      BGBlu = Surf.Dib1D(DrawX).Blue
      BGGrn = Surf.Dib1D(DrawX).Green
      BGRed = Surf.Dib1D(DrawX).Red
      FGRed = 255 - BGRed
      FGGrn = 255 - BGGrn
      FGBlu = 255 - BGBlu
      Surf.Dib1D(DrawX).Blue = BGBlu + sngAlpha * (FGBlu - BGBlu)
      Surf.Dib1D(DrawX).Green = BGGrn + sngAlpha * (FGGrn - BGGrn)
      Surf.Dib1D(DrawX).Red = BGRed + sngAlpha * (FGRed - BGRed)
     End If
    End If
    delta_x = delta_x + brush_slope
   Next
   delta_y = delta_y + brush_slope
  Next DrawY

  Else 'filled circle

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     If sngAlpha > sPress Then sngAlpha = sPress
     BGBlu = Surf.Dib1D(DrawX).Blue
     BGGrn = Surf.Dib1D(DrawX).Green
     BGRed = Surf.Dib1D(DrawX).Red
     FGRed = 255 - BGRed
     FGGrn = 255 - BGGrn
     FGBlu = 255 - BGBlu
     Surf.Dib1D(DrawX).Blue = BGBlu + sngAlpha * (FGBlu - BGBlu)
     Surf.Dib1D(DrawX).Green = BGGrn + sngAlpha * (FGGrn - BGGrn)
     Surf.Dib1D(DrawX).Red = BGRed + sngAlpha * (FGRed - BGRed)
    End If
    delta_x = delta_x + brush_slope
   Next
   delta_y = delta_y + brush_slope
  Next

  End If 'Airbrush.DoOutline

  Airbrush.Pressure = TmpPressure

  ElseIf Airbrush.Mode_0_to_2 = CShift Then

  If Airbrush.DoOutline Then

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     If deltas_xy_Sq > first_slice_Sq Then
      sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     Else
      If deltas_xy_Sq < second_slice_Sq Then
       If deltas_xy_Sq >= inv_tip_Sq Then
        rgb_shift_intensity = sPress - (second_cone_slice - Sqr#(deltas_xy_Sq)) * cushion_inside_multiplier
       Else
        rgb_shift_intensity = 0
       End If
      Else
       rgb_shift_intensity = sPress
      End If
     End If
     If rgb_shift_intensity > 0 Then
      sB = Surf.Dib1D(DrawX).Blue
      sG = Surf.Dib1D(DrawX).Green
      sR = Surf.Dib1D(DrawX).Red
      ColorShift
      Surf.Dib1D(DrawX).Blue = sB
      Surf.Dib1D(DrawX).Green = sG
      Surf.Dib1D(DrawX).Red = sR
     End If
    End If
    delta_x = delta_x + brush_slope
   Next
   delta_y = delta_y + brush_slope
  Next DrawY

  Else 'filled circle

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     rgb_shift_intensity = brush_height - Sqr#(deltas_xy_Sq)
     If rgb_shift_intensity > sPress Then rgb_shift_intensity = sPress
     sB = Surf.Dib1D(DrawX).Blue
     sG = Surf.Dib1D(DrawX).Green
     sR = Surf.Dib1D(DrawX).Red
     ColorShift
     Surf.Dib1D(DrawX).Blue = sB
     Surf.Dib1D(DrawX).Green = sG
     Surf.Dib1D(DrawX).Red = sR
    End If
    delta_x = delta_x + brush_slope
   Next
   delta_y = delta_y + brush_slope
  Next

  End If 'Airbrush.DoOutline

  End If 'AirBrush.Mode_0_to_2 = %MM_CSolid


 Else 'not erasing


  If Airbrush.Mode_0_to_2 = CSolid Then

  If Airbrush.DoOutline Then

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     If deltas_xy_Sq > first_slice_Sq Then
      sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     Else
      If deltas_xy_Sq < second_slice_Sq Then
       If deltas_xy_Sq >= inv_tip_Sq Then
        sngAlpha = sPress - (second_cone_slice - Sqr#(deltas_xy_Sq)) * cushion_inside_multiplier
       Else
        sngAlpha = 0
       End If
      Else
       sngAlpha = sPress
      End If
     End If
     If sngAlpha > 0 Then
      sngAlpha = sngAlpha '/ 255
      BGColor = Surf.LDib1D(DrawX)
      BGGrn = BGColor And vbGreen
      FGGrn = FGColor And vbGreen
      BGBlu = BGColor And &HFF&
      FGBlu = FGColor And &HFF&
      BGR = _
         BGColor + sngAlpha * (FGColor - BGColor) And vbBlue Or _
         BGGrn + sngAlpha * (FGGrn - BGGrn) And vbGreen Or _
         BGBlu + sngAlpha * (FGBlu - BGBlu) And &HFF&
      Surf.LDib1D(DrawX) = BGR
      Surf.EraseDib(DrawX) = BGR
     End If
    End If
    delta_x = delta_x + brush_slope
   Next DrawX
   delta_y = delta_y + brush_slope
  Next DrawY

  Else 'filled circle

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     If sngAlpha > sPress Then sngAlpha = sPress
      'sngAlpha = sngAlpha '/ 25
      BGColor = Surf.LDib1D(DrawX)
      BGGrn = BGColor And vbGreen
      FGGrn = FGColor And vbGreen
      BGBlu = BGColor And &HFF&
      FGBlu = FGColor And &HFF&
      Surf.LDib1D(DrawX) = _
         BGColor + sngAlpha * (FGColor - BGColor) And vbBlue Or _
         BGGrn + sngAlpha * (FGGrn - BGGrn) And vbGreen Or _
         BGBlu + sngAlpha * (FGBlu - BGBlu) And &HFF&
      'Surf.LDib1D(DrawX) = BGR
      'Surf.EraseDib(DrawX) = BGR
    End If
    delta_x = delta_x + brush_slope
   Next
   delta_y = delta_y + brush_slope
  Next

  End If 'Airbrush.DoOutline

  ElseIf Airbrush.Mode_0_to_2 = MM_invert Then

  If Airbrush.DoOutline Then

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     If deltas_xy_Sq > first_slice_Sq Then
      sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     Else
      If deltas_xy_Sq < second_slice_Sq Then
       If deltas_xy_Sq >= inv_tip_Sq Then
        sngAlpha = sPress - second_cone_slice + Sqr#(deltas_xy_Sq) * cushion_inside_multiplier
       Else
        sngAlpha = 0
       End If
      Else
       sngAlpha = sPress
      End If
     End If
     If sngAlpha > 0 Then
      BGBlu = Surf.Dib(DrawX).Blue
      BGGrn = Surf.Dib1D(DrawX).Green
      BGRed = Surf.Dib1D(DrawX).Red
      FGRed = 255 - BGRed
      FGGrn = 255 - BGGrn
      FGBlu = 255 - BGBlu
      Surf.Dib1D(DrawX).Blue = BGBlu + sngAlpha * (FGBlu - BGBlu)
      Surf.Dib1D(DrawX).Green = BGGrn + sngAlpha * (FGGrn - BGGrn)
      Surf.Dib1D(DrawX).Red = BGRed + sngAlpha * (FGRed - BGRed)
      Surf.EraseDib(DrawX) = Surf.LDib1D(DrawX)
     End If
    End If
    delta_x = delta_x + brush_slope
   Next DrawX
   delta_y = delta_y + brush_slope
  Next DrawY

  Else 'filled circle

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     If sngAlpha > sPress Then sngAlpha = sPress
     BGBlu = Surf.Dib1D(DrawX).Blue
     BGGrn = Surf.Dib1D(DrawX).Green
     BGRed = Surf.Dib1D(DrawX).Red
     FGRed = 255 - BGRed
     FGGrn = 255 - BGGrn
     FGBlu = 255 - BGBlu
     Surf.Dib1D(DrawX).Blue = BGBlu + sngAlpha * (FGBlu - BGBlu) '/ 255
     Surf.Dib1D(DrawX).Green = BGGrn + sngAlpha * (FGGrn - BGGrn) '/ 255
     Surf.Dib1D(DrawX).Red = BGRed + sngAlpha * (FGRed - BGRed) '/ 255
     Surf.EraseDib(DrawX) = Surf.LDib1D(DrawX)
    End If
    delta_x = delta_x + brush_slope
   Next DrawX
   delta_y = delta_y + brush_slope
  Next DrawY

  End If 'Airbrush.DoOutline

  Airbrush.Pressure = TmpPressure

  ElseIf Airbrush.Mode_0_to_2 = CShift Then

  If Airbrush.DoOutline Then

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     If deltas_xy_Sq > first_slice_Sq Then
      sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     Else
      If deltas_xy_Sq < second_slice_Sq Then
       'If deltas_xy_Sq > inv_tip_Sq Then
       ' sngalpha = Sqr#(deltas_xy_Sq) - inv_cone_tip2
       If deltas_xy_Sq >= inv_tip_Sq Then
        sngAlpha = sPress - second_cone_slice + Sqr#(deltas_xy_Sq) * cushion_inside_multiplier
       Else
        sngAlpha = 0
       End If
      Else
       sngAlpha = sPress
      End If
     End If
     If sngAlpha > 0 Then
      rgb_shift_intensity! = sngAlpha
      sB = Surf.Dib1D(DrawX).Blue
      sG = Surf.Dib1D(DrawX).Green
      sR = Surf.Dib1D(DrawX).Red
      ColorShift
      Surf.Dib1D(DrawX).Blue = sB
      Surf.Dib1D(DrawX).Green = sG
      Surf.Dib1D(DrawX).Red = sR
      Surf.EraseDib(DrawX) = Surf.LDib1D(DrawX)
     End If
    End If
    delta_x = delta_x + brush_slope
   Next DrawX
   delta_y = delta_y + brush_slope
  Next DrawY

  Else 'filled circle

  For DrawY = DrawBot To DrawTop Step Surf.Dims.Wide
   delta_ySq = delta_y * delta_y
   delta_x = delta_left
   DrawRight = DrawY + AddDrawWidth
   For DrawX = DrawY To DrawRight
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
     sngAlpha = brush_height - Sqr#(deltas_xy_Sq)
     If sngAlpha > sPress Then sngAlpha = sPress
     rgb_shift_intensity! = sngAlpha
     sB = Surf.Dib1D(DrawX).Blue
     sG = Surf.Dib1D(DrawX).Green
     sR = Surf.Dib1D(DrawX).Red
     ColorShift
     Surf.Dib1D(DrawX).Blue = sB
     Surf.Dib1D(DrawX).Green = sG
     Surf.Dib1D(DrawX).Red = sR
     Surf.EraseDib(DrawX) = Surf.LDib1D(DrawX)
    End If
    delta_x = delta_x + brush_slope
   Next DrawX
   delta_y = delta_y + brush_slope
  Next DrawY

  End If 'Airbrush.DoOutline

  End If 'AirBrush.Mode_0_to_2

 End If 'erasing

 End If 'Airbrush.diameter > 0

End If 'surf.TotalPixels > 0

End Sub

Public Sub ColorShift()

 If sR < sB Then
  If sR < sG Then
   If sG < sB Then
    bytMaxMin_diff = sB - sR
    sG = sG - bytMaxMin_diff * rgb_shift_intensity
    If sG < sR Then
     iSubt = sR - sG
     sG = sR
     sR = sR + iSubt
    End If
   Else
    bytMaxMin_diff = sG - sR
    sB = sB + bytMaxMin_diff * rgb_shift_intensity
    If sB > sG Then
     iSubt = sB - sG
     sB = sG
     sG = sG - iSubt
    End If
   End If
  Else
   bytMaxMin_diff = sB - sG
   sR = sR + bytMaxMin_diff * rgb_shift_intensity
   If sR > sB Then
    iSubt = sR - sB
    sR = sB
    sB = sB - iSubt
   End If
  End If
 ElseIf sR > sG Then
  If sB < sG Then
   bytMaxMin_diff = sR - sB
   sG = sG + bytMaxMin_diff * rgb_shift_intensity
   If sG > sR Then
    iSubt = sG - sR
    sG = sR
    sR = sR - iSubt
   End If
  Else
   bytMaxMin_diff = sR - sG
   sB = sB - bytMaxMin_diff * rgb_shift_intensity
   If sB < sG Then
    iSubt = sG - sB
    sB = sG
    sG = sG + iSubt
   End If
  End If
 Else
  bytMaxMin_diff = sG - sB
  sR = sR - bytMaxMin_diff * rgb_shift_intensity
  If sR < sB Then
   iSubt = sB - sR
   sR = sB
   sB = sB + iSubt
  End If
 End If
 
End Sub


