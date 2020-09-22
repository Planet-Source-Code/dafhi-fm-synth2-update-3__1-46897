Attribute VB_Name = "mod_vControl"
Option Explicit

Private N&

Public vSelect As Long 'which control

Public blnSliderClicked As Boolean 'default = 0 = button clicked
Private rolo_previous_y As Single

Private Type PrecisionRGB
 sRed As Single
 sGrn As Single
 sBlu As Single
End Type

Public Type KnobAlpha1
 KnobColor As PrecisionRGB
 DotColor  As PrecisionRGB
 StrText   As String
 FontSize  As Long
 TextColor As Long
 DoMaxAdjustRate As Boolean
 'MarkColor As PrecisionRGB
 vMax      As Single
 vMin      As Single
 vVal      As Single
 vstep     As Single
 knobSize  As Single
 dotSize   As Single
 Index      As Long
 angle     As Single
 mouse_mult As Single
 posHalfW   As Single
 posHalfH   As Single
 CurrX     As Integer
 CurrY     As Integer
 FCG       As FCGRect
End Type

Private C_SlideAuto

Private mouseDown_angle!
Private mouseMove_angle!
Private angle_difference!

Private SlideAuto(1 To 200) As KnobAlpha1
Public ControlSurface(1 To 200) As AnimSurf2D

Public Sub InitializeSlideControl(SLC As KnobAlpha1, ByVal Left As Integer, ByVal Top As Integer, Width As Integer, Height As Integer, ByVal sngMinValue!, ByVal sngMaxValue!, ByVal Steps&, Optional ByVal MouseStepsPerIncr As Byte = 1, Optional knobSizePercent! = 0.8, Optional dotSizePercent! = 0.1, Optional DisableValSetToMin As Boolean = False, Optional StrText$, Optional FontSize% = 10, Optional TopLeftColor&, Optional TopRightColor&, Optional BottomLeftColor&, Optional BottomRightColor&, Optional KnobColor&, Optional DotColor&, Optional ByVal DoAdjustRateMax As Boolean = False)

If MouseStepsPerIncr = 0 Then MouseStepsPerIncr = 1

With SLC.FCG.IProcess.Dims
 .High = Height
 .Wide = Width
 .HighM1 = Height - 1
 .WideM1 = Width - 1
 .LowLeftPos.x = Left
 .TopRightPos.y = Top
 .LowLeftPos.y = Top + .HighM1
 .TopRightPos.x = Left + .WideM1
 SLC.posHalfW = (Left + .TopRightPos.x) / 2
 SLC.posHalfH = (Top + .LowLeftPos.y) / 2
End With

With SLC
 .vMax = sngMaxValue
 .vMin = sngMinValue
 If Not DisableValSetToMin Then
 .vVal = .vMin
 End If
 .vstep = (.vMax - .vMin) / Steps
 If Width > Height Then
  .knobSize = Height * knobSizePercent
 Else
  .knobSize = Width * knobSizePercent
 End If
 .dotSize = .knobSize * dotSizePercent / 2
 .angle = pi * 1.25 - pi * 1.5 * (.vVal - .vMin) / (.vMax - .vMin)
 .mouse_mult = 1 / MouseStepsPerIncr
End With

modFCG.SetCornerColor SLC.FCG.TopLeft, TopLeftColor
modFCG.SetCornerColor SLC.FCG.TopRight, TopRightColor
modFCG.SetCornerColor SLC.FCG.LowLeft, BottomLeftColor
modFCG.SetCornerColor SLC.FCG.LowRight, BottomRightColor

SRGBColor SLC.DotColor, DotColor
SRGBColor SLC.KnobColor, KnobColor

SLC.DoMaxAdjustRate = DoAdjustRateMax

C_SlideAuto = C_SlideAuto + 1
LSet SlideAuto(C_SlideAuto) = SLC
SLC.Index = C_SlideAuto

MakeSurf ControlSurface(C_SlideAuto), Width, Height
WrapFCG ControlSurface(C_SlideAuto), SLC.FCG
SLC.StrText = StrText
SLC.FontSize = FontSize
SLC.CurrX = 0
SLC.CurrY = 0 'Height - FontSize

DrawText ControlSurface(SLC.Index), SLC.CurrX, SLC.CurrY, SLC
DrawDots ControlSurface(C_SlideAuto), SLC, True, False, False

CopyMainToErase ControlSurface(C_SlideAuto)

DrawDots ControlSurface(C_SlideAuto), SLC

End Sub

Public Sub Clicked_Which_Slide_Ary(Obj As Object, Button As Integer, Shift As Integer, x As Single, y As Single)
 vSelect = 0 'Reset to client space
 blnSliderClicked = False 'False = "Button Clicked"
 
 
 If vSelect = 0 Then
  For N = 1 To C_SlideAuto
  With SlideAuto(N).FCG.IProcess.Dims
  If x < .TopRightPos.x And x > .LowLeftPos.x Then
   If y < .LowLeftPos.y And y > .TopRightPos.y Then
   
    'Default = False = "Button Clicked"
    blnSliderClicked = True
    
    'Form_MouseMove understands this
    vSelect = N
    
    rolo_previous_y = y
   
   End If
  End If
  End With
  Next N
    
 End If

End Sub
Public Sub CalcSlideVal(Obj As Object, VSL1 As KnobAlpha1, y As Single)

 If blnSliderClicked Then
  With VSL1.FCG.IProcess.Dims
   If y < .TopRightPos.y Then
   VSL1.vVal = VSL1.vMax
   ElseIf y > .LowLeftPos.y Then
   VSL1.vVal = VSL1.vMin
   Else 'caculate value based upon ratio of mouse Y
        'minus control top Y / control height
   VSL1.vVal = VSL1.vMin + (VSL1.vMax - VSL1.vMin) * (.LowLeftPos.y - y) / (.High)
   End If
   Obj.Caption = "Slider " & vSelect & "  " & VSL1.vVal
  End With
 End If

End Sub
Public Sub CalcSlideValRoloStyle(Obj As Object, RSLi As KnobAlpha1, y!, Optional yScale! = 1!)
  With RSLi
   .vVal = .vVal + yScale * (rolo_previous_y - y)
   If .vVal > .vMax Then
    .vVal = .vMax
   ElseIf .vVal < .vMin Then
    .vVal = .vMin
   End If
  End With
  rolo_previous_y = y
  DrawDots ControlSurface(vSelect), RSLi
  With SlideAuto(vSelect).FCG.IProcess.Dims
   BlitToDC Obj.hdc, ControlSurface(vSelect), .LowLeftPos.x, .TopRightPos.y
  End With
End Sub
Public Sub CalcKnobVal(Obj As Object, RSLi As KnobAlpha1, y!)
Dim tVal!

 tVal = RSLi.mouse_mult * (rolo_previous_y - y)
 
 If RSLi.DoMaxAdjustRate Then
  If tVal >= 3 Then
   tVal = 2
  ElseIf tVal <= -3 Then
   tVal = -2
  End If
 End If
 
 If tVal >= 1 Or tVal <= -1 Then
 
  With RSLi
   .vVal = .vVal + .vstep * Int(tVal + 0.5)
   If .vMax > .vMin Then
    If .vVal > .vMax Then
     .vVal = .vMax
    ElseIf .vVal < .vMin Then
     .vVal = .vMin
    End If
   Else
    If .vVal > .vMin Then
     .vVal = .vMin
    ElseIf .vVal < .vMax Then
     .vVal = .vMax
    End If
   End If
  End With
  rolo_previous_y = y 'rolo_previous_y + (y - rolo_previous_y) * 1
  RSLi.angle = pi * 1.25 - pi * 1.5 * (RSLi.vVal - RSLi.vMin) / (RSLi.vMax - RSLi.vMin)
  DrawDots ControlSurface(vSelect), RSLi
  With SlideAuto(vSelect).FCG.IProcess.Dims
   BlitToDC Obj.hdc, ControlSurface(vSelect), .LowLeftPos.x, .TopRightPos.y
  End With
  
 End If
 
End Sub
Private Sub DrawDots(Surf As AnimSurf2D, RSLi As KnobAlpha1, Optional ByVal DoKnob As Boolean = False, Optional ByVal DoDot As Boolean = True, Optional ByVal DoCopyFromErase As Boolean = True)
Dim radius!

 If DoCopyFromErase Then CopyEraseToMain Surf
 
 Airbrush.Blue = RSLi.KnobColor.sBlu
 Airbrush.Green = RSLi.KnobColor.sGrn
 Airbrush.Red = RSLi.KnobColor.sRed
 Airbrush.diameter = RSLi.knobSize
 Airbrush.definition = Airbrush.diameter / 2
 Airbrush.Pressure = 120
 
 If DoKnob Then
 BlitAirbrush Surf, Surf.halfW, Surf.halfH
 End If
 
 If DoDot Then
 radius = Airbrush.definition * 0.62
 
 Airbrush.Blue = RSLi.DotColor.sBlu
 Airbrush.Green = RSLi.DotColor.sGrn
 Airbrush.Red = RSLi.DotColor.sRed
 Airbrush.diameter = RSLi.dotSize
 Airbrush.definition = Airbrush.diameter / 2
 Airbrush.Pressure = 255
 
 BlitAirbrush Surf, Surf.halfW + radius * Cos(RSLi.angle), _
                    Surf.halfH + radius * Sin(RSLi.angle)
 End If
 
End Sub
Public Sub Draw_Slides(Obj As Object)

 For N = 1 To C_SlideAuto
 With SlideAuto(N)
  BlitToDC Obj.hdc, ControlSurface(N), .FCG.IProcess.Dims.LowLeftPos.x, .FCG.IProcess.Dims.TopRightPos.y
 End With
 Next N

End Sub
Public Sub ClearControlSurfaces()
 For N = 1 To C_SlideAuto
  ClearSurface ControlSurface(N)
 Next N
End Sub
Public Sub GetAngleFromDXDY(ByRef retAngle As Single, sngDX As Single, sngDY As Single)
 If sngDY = 0 Then
  If sngDX < 0 Then
   retAngle = pi * 3& / 2&
  ElseIf sngDX > 0 Then
   retAngle = pi / 2&
  End If
 Else
  If sngDY > 0 Then
   retAngle = pi - Atn(sngDX / sngDY)
  Else
   retAngle = Atn(sngDX / -sngDY)
  End If
 End If
End Sub
Public Sub ResetSlide()
 If vSelect <> 0 Then
  SlideAuto(vSelect).angle = angle_difference
  vSelect = 0
 End If
End Sub
Private Sub AngleModulus(ByRef retAngle As Single)
 retAngle = retAngle - TwoPi * Int(retAngle / TwoPi)
End Sub

Private Sub SRGBColor(SRGB1 As PrecisionRGB, BGR1&)
 SRGB1.sBlu = (BGR1 And vbBlue) / 65536
 SRGB1.sGrn = (BGR1 And &HFF00&) / 256
 SRGB1.sRed = BGR1 And &HFF&
End Sub
