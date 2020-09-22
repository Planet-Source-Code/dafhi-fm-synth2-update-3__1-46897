Attribute VB_Name = "modControls"
Option Explicit

Private N&

Public vSelect As Long 'which control

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

Private C_KnobAuto

Private mouseDown_angle!
Private mouseMove_angle!
Private angle_difference!

Private KnobAuto(1 To 200) As KnobAlpha1
Public ControlSurface(1 To 200) As AnimSurf2D

Public ShiftKeyDown As Long

Public Sub InitializeKnob(SLC As KnobAlpha1, ByVal Left As Integer, ByVal Top As Integer, Width As Integer, Height As Integer, ByVal sngMinValue!, ByVal sngMaxValue!, ByVal Steps&, Optional ByVal MouseStepsPerIncr As Byte = 1, Optional knobSizePercent! = 0.8, Optional dotSizePercent! = 0.1, Optional StrText$, Optional FontSize% = 10, Optional TopLeftColor&, Optional TopRightColor&, Optional BottomLeftColor&, Optional BottomRightColor&, Optional KnobColor&, Optional DotColor&, Optional ByVal DoAdjustRateMax As Boolean = False)

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
 .vstep = (.vMax - .vMin) / Steps
 If Width > Height Then
  .knobSize = Height * knobSizePercent
 Else
  .knobSize = Width * knobSizePercent
 End If
 .dotSize = .knobSize * dotSizePercent / 2
 .mouse_mult = 1 / MouseStepsPerIncr
End With

modFCG.SetCornerColor SLC.FCG.TopLeft, TopLeftColor
modFCG.SetCornerColor SLC.FCG.TopRight, TopRightColor
modFCG.SetCornerColor SLC.FCG.LowLeft, BottomLeftColor
modFCG.SetCornerColor SLC.FCG.LowRight, BottomRightColor

SRGBColor SLC.DotColor, DotColor
SRGBColor SLC.KnobColor, KnobColor

SLC.DoMaxAdjustRate = DoAdjustRateMax

C_KnobAuto = C_KnobAuto + 1
LSet KnobAuto(C_KnobAuto) = SLC
SLC.Index = C_KnobAuto

MakeSurf ControlSurface(C_KnobAuto), Width, Height
WrapFCG ControlSurface(C_KnobAuto), SLC.FCG
SLC.StrText = StrText
SLC.FontSize = FontSize
SLC.CurrX = 0
SLC.CurrY = 0 'Height - FontSize

DrawText ControlSurface(SLC.Index), SLC.CurrX, SLC.CurrY, SLC
DrawDots ControlSurface(C_KnobAuto), SLC, True, False, False

CopyMainToErase ControlSurface(C_KnobAuto)

End Sub

Public Sub Clicked_Which_Knob_Ary(Obj As Object, Button As Integer, Shift As Integer, x As Single, y As Single)

 vSelect = 0 'Reset to client space
 'blnSliderClicked = False 'False = "Button Clicked"
 
 
 If vSelect = 0 Then
  For N = 1 To C_KnobAuto
  With KnobAuto(N).FCG.IProcess.Dims
  If x < .TopRightPos.x And x > .LowLeftPos.x Then
   If y < .LowLeftPos.y And y > .TopRightPos.y Then
   
    'Default = False = "Button Clicked"
    'blnSliderClicked = True
    
    'Form_MouseMove understands this
    vSelect = N
    
    rolo_previous_y = y
   
   End If
  End If
  End With
  Next N
    
 End If

End Sub
Public Sub CalcKnobVal(Obj As Object, RSLi As KnobAlpha1, y!)
Dim tVal!

 tVal = RSLi.mouse_mult * (rolo_previous_y - y)
 'If ShiftKeyDown Then tVal = tVal / 10
 
 If RSLi.DoMaxAdjustRate Then
  If tVal >= 3 Then
   tVal = 2
  ElseIf tVal <= -3 Then
   tVal = -2
  End If
 End If
 
 If tVal >= 1 Or tVal <= -1 Then
 
  With RSLi
   If ShiftKeyDown Then
   .vVal = .vVal + .vstep * Int(tVal + 0.5) / 10
   Else
   .vVal = .vVal + .vstep * Int(tVal + 0.5)
   End If
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
  KnobAuto(vSelect).vVal = RSLi.vVal
  DrawDots ControlSurface(vSelect), KnobAuto(vSelect)
  With KnobAuto(vSelect).FCG.IProcess.Dims
   BlitToDC Obj.hdc, ControlSurface(vSelect), .LowLeftPos.x, .TopRightPos.y
  End With
  
 End If
 
End Sub
Public Sub Draw_Knobs(Obj As Object)

 For N = 1 To C_KnobAuto
 With KnobAuto(N)
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
  DrawDots ControlSurface(N), KnobAuto(N), False, True, True
  BlitToDC Obj.hdc, ControlSurface(N), .FCG.IProcess.Dims.LowLeftPos.x, .FCG.IProcess.Dims.TopRightPos.y
 End With
 Next N

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
 
 RSLi.angle = pi * 1.25 - pi * 1.5 * (RSLi.vVal - RSLi.vMin) / (RSLi.vMax - RSLi.vMin)
 BlitAirbrush Surf, Surf.halfW + radius * Cos(RSLi.angle), _
                    Surf.halfH + radius * Sin(RSLi.angle)
 End If
 
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
Public Sub ResetKnob()
 If vSelect <> 0 Then
  KnobAuto(vSelect).angle = angle_difference
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
Public Sub ClearControlSurfaces()
 For N = 1 To C_KnobAuto
  ClearSurface ControlSurface(N)
 Next N
End Sub

