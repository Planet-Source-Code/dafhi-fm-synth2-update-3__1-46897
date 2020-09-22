Attribute VB_Name = "modFCG"
Option Explicit

Private Type PointApi
 x As Long
 y As Long
End Type

Private Type RGBTripleInt
 Red As Integer
 Green As Integer
 Blue As Integer
End Type

Private Type RectDims
 Wide As Long
 High As Long
 WideM1 As Long
 HighM1 As Long
 LowLeftPos As PointApi
 TopRightPos As PointApi
End Type

Private Type RGBDiffAPI
 sRed As Single
 sGrn As Single
 sBlu As Single
End Type

Private Type FCG_System
 LeftRGBdelta As RGBDiffAPI
 TopRGBdelta As RGBDiffAPI
 BottomRGBdelta As RGBDiffAPI
 LeftRGBi As RGBDiffAPI
 TopRGBi As RGBDiffAPI
 BottomRGBi As RGBDiffAPI
 Dims As RectDims
End Type

Public Type FCGRect
 LowLeft As RGBTripleInt
 LowRight As RGBTripleInt
 TopLeft As RGBTripleInt
 TopRight As RGBTripleInt
 IProcess As FCG_System
End Type

Public FourColorGradient As FCGRect

'DrawFCG and WrapFCG
Dim iBlue!
Dim iGreen!
Dim iRed!
Dim iiBlu!
Dim iiGrn!
Dim iiRed!
Dim iiBlu2!
Dim iiGrn2!
Dim iiRed2!
Dim sngBlue!
Dim sngGreen!
Dim sngRed!
Dim vertBlue!
Dim vertGreen!
Dim vertRed!

Dim DrawTop&
Dim DrawBot&
Dim DrawLeft&
Dim DrawRight&

Private DrawX&
Private DrawY&

Public Sub FCG_ColorLowLeft(FCG As FCGRect, Red&, Green&, Blue&)
 FCG.LowLeft.Red = Red
 FCG.LowLeft.Green = Green
 FCG.LowLeft.Blue = Blue
End Sub
Public Sub FCG_ColorLowRight(FCG As FCGRect, Red&, Green&, Blue&)
 FCG.LowRight.Red = Red
 FCG.LowRight.Green = Green
 FCG.LowRight.Blue = Blue
End Sub
Public Sub FCG_ColorTopLeft(FCG As FCGRect, Red&, Green&, Blue&)
 FCG.TopLeft.Red = Red
 FCG.TopLeft.Green = Green
 FCG.TopLeft.Blue = Blue
End Sub
Public Sub FCG_ColorTopRight(FCG As FCGRect, Red&, Green&, Blue&)
 FCG.TopRight.Red = Red
 FCG.TopRight.Green = Green
 FCG.TopRight.Blue = Blue
End Sub

Private Sub SetLowLeftCorner(FCG As FCGRect, x&, y&)
 FCG.IProcess.Dims.LowLeftPos.x = x
 FCG.IProcess.Dims.LowLeftPos.y = y
End Sub
Private Sub SetFCGDims(FCG As FCGRect, Width&, Height&)
With FCG.IProcess.Dims
 .Wide = Width
 .High = Height
 .WideM1 = Width - 1
 .HighM1 = Height - 1
 .TopRightPos.x = .LowLeftPos.x + .WideM1
 .TopRightPos.y = .LowLeftPos.y + .HighM1
End With
End Sub

Public Sub WrapFCG(Surf As AnimSurf2D, FCG As FCGRect)
 
  SetLowLeftCorner FourColorGradient, 0, 0
  SetFCGDims FourColorGradient, Surf.Dims.Wide, Surf.Dims.High
  
  With FourColorGradient
  .IProcess.BottomRGBdelta.sBlu = FCG.LowRight.Blue - FCG.LowLeft.Blue
  .IProcess.BottomRGBdelta.sGrn = FCG.LowRight.Green - FCG.LowLeft.Green
  .IProcess.BottomRGBdelta.sRed = FCG.LowRight.Red - FCG.LowLeft.Red

  .IProcess.TopRGBdelta.sBlu = FCG.TopRight.Blue - FCG.TopLeft.Blue
  .IProcess.TopRGBdelta.sGrn = FCG.TopRight.Green - FCG.TopLeft.Green
  .IProcess.TopRGBdelta.sRed = FCG.TopRight.Red - FCG.TopLeft.Red
  
  .IProcess.LeftRGBdelta.sBlu = FCG.TopLeft.Blue - FCG.LowLeft.Blue
  .IProcess.LeftRGBdelta.sGrn = FCG.TopLeft.Green - FCG.LowLeft.Green
  .IProcess.LeftRGBdelta.sRed = FCG.TopLeft.Red - FCG.LowLeft.Red
  End With
  
  With FourColorGradient.IProcess
   
   If .Dims.WideM1 > 0 Then
   .TopRGBi.sBlu = .TopRGBdelta.sBlu / .Dims.WideM1
   .TopRGBi.sGrn = .TopRGBdelta.sGrn / .Dims.WideM1
   .TopRGBi.sRed = .TopRGBdelta.sRed / .Dims.WideM1
   .BottomRGBi.sBlu = .BottomRGBdelta.sBlu / .Dims.WideM1
   .BottomRGBi.sGrn = .BottomRGBdelta.sGrn / .Dims.WideM1
   .BottomRGBi.sRed = .BottomRGBdelta.sRed / .Dims.WideM1
   End If
   
   If .Dims.HighM1 > 0 Then
   .LeftRGBi.sBlu = .LeftRGBdelta.sBlu / .Dims.HighM1
   .LeftRGBi.sGrn = .LeftRGBdelta.sGrn / .Dims.HighM1
   .LeftRGBi.sRed = .LeftRGBdelta.sRed / .Dims.HighM1
    iiRed = (.TopRGBi.sRed - .BottomRGBi.sRed) / .Dims.HighM1
    iiGrn = (.TopRGBi.sGrn - .BottomRGBi.sGrn) / .Dims.HighM1
    iiBlu = (.TopRGBi.sBlu - .BottomRGBi.sBlu) / .Dims.HighM1
   End If
  
   iRed = .BottomRGBi.sRed
   iGreen = .BottomRGBi.sGrn
   iBlue = .BottomRGBi.sBlu
   
  End With
  
  vertBlue = FCG.LowLeft.Blue + 0.00001 'nullify most rounding errors
  vertGreen = FCG.LowLeft.Green + 0.00001
  vertRed = FCG.LowLeft.Red + 0.00001
      
  For DrawY = FourColorGradient.IProcess.Dims.LowLeftPos.y To FourColorGradient.IProcess.Dims.TopRightPos.y
   sngBlue = vertBlue
   sngGreen = vertGreen
   sngRed = vertRed
   vertBlue = vertBlue + FourColorGradient.IProcess.LeftRGBi.sBlu
   vertGreen! = vertGreen + FourColorGradient.IProcess.LeftRGBi.sGrn
   vertRed! = vertRed + FourColorGradient.IProcess.LeftRGBi.sRed
   For DrawX = FourColorGradient.IProcess.Dims.LowLeftPos.x To FourColorGradient.IProcess.Dims.TopRightPos.x
    Surf.Dib(DrawX, DrawY).Blue = sngBlue
    Surf.Dib(DrawX, DrawY).Green = sngGreen
    Surf.Dib(DrawX, DrawY).Red = sngRed
    sngBlue = sngBlue + iBlue
    sngGreen = sngGreen + iGreen
    sngRed = sngRed + iRed
   Next DrawX
   iBlue = iBlue + iiBlu
   iGreen = iGreen + iiGrn
   iRed = iRed + iiRed
  Next DrawY
  
End Sub

Public Sub DrawFCG(Surf As AnimSurf2D, FCG As FCGRect, ByVal Left&, ByVal Top&, Width&, Height&)
Dim ClipLeft&
Dim ClipBot&
  
  'SetLowLeftCorner Left, Surf.Dims.High - Top - Height
  'SetFCGDims Width, Height
  
  If Left < 0 Then
   ClipLeft = -Left
   DrawLeft = 0
  Else
   DrawLeft = Left
  End If
  
  If FCG.IProcess.Dims.LowLeftPos.y < 0 Then
   DrawBot = 0
   ClipBot = -FCG.IProcess.Dims.LowLeftPos.y
  Else
   DrawBot = FCG.IProcess.Dims.LowLeftPos.y
  End If
    
  If FCG.IProcess.Dims.TopRightPos.y > Surf.Dims.HighM1 Then
   DrawTop = Surf.Dims.HighM1
  Else
   DrawTop = FCG.IProcess.Dims.TopRightPos.y
  End If
  
  If FCG.IProcess.Dims.TopRightPos.x > Surf.Dims.WideM1 Then
   DrawRight = Surf.Dims.WideM1
  Else
   DrawRight = FCG.IProcess.Dims.TopRightPos.x
  End If
    
  With FCG
  .IProcess.BottomRGBdelta.sBlu = .LowRight.Blue - .LowLeft.Blue
  .IProcess.BottomRGBdelta.sGrn = .LowRight.Green - .LowLeft.Green
  .IProcess.BottomRGBdelta.sRed = .LowRight.Red - .LowLeft.Red

  .IProcess.TopRGBdelta.sBlu = .TopRight.Blue - .TopLeft.Blue
  .IProcess.TopRGBdelta.sGrn = .TopRight.Green - .TopLeft.Green
  .IProcess.TopRGBdelta.sRed = .TopRight.Red - .TopLeft.Red
  
  .IProcess.LeftRGBdelta.sBlu = .TopLeft.Blue - .LowLeft.Blue
  .IProcess.LeftRGBdelta.sGrn = .TopLeft.Green - .LowLeft.Green
  .IProcess.LeftRGBdelta.sRed = .TopLeft.Red - .LowLeft.Red
  End With
  
  If FCG.IProcess.Dims.HighM1 > 0 Then
  With FCG.IProcess
   
   If .Dims.WideM1 > 0 Then
   .TopRGBi.sBlu = .TopRGBdelta.sBlu / .Dims.WideM1
   .TopRGBi.sGrn = .TopRGBdelta.sGrn / .Dims.WideM1
   .TopRGBi.sRed = .TopRGBdelta.sRed / .Dims.WideM1
   .BottomRGBi.sBlu = .BottomRGBdelta.sBlu / .Dims.WideM1
   .BottomRGBi.sGrn = .BottomRGBdelta.sGrn / .Dims.WideM1
   .BottomRGBi.sRed = .BottomRGBdelta.sRed / .Dims.WideM1
   End If
   
   .LeftRGBi.sBlu = .LeftRGBdelta.sBlu / .Dims.HighM1
   .LeftRGBi.sGrn = .LeftRGBdelta.sGrn / .Dims.HighM1
   .LeftRGBi.sRed = .LeftRGBdelta.sRed / .Dims.HighM1
    iiRed = (.TopRGBi.sRed - .BottomRGBi.sRed) / .Dims.HighM1
    iiGrn = (.TopRGBi.sGrn - .BottomRGBi.sGrn) / .Dims.HighM1
    iiBlu = (.TopRGBi.sBlu - .BottomRGBi.sBlu) / .Dims.HighM1
  
   iRed = .BottomRGBi.sRed + iiRed * ClipBot
   iGreen = .BottomRGBi.sGrn + iiGrn * ClipBot
   iBlue = .BottomRGBi.sBlu + iiBlu * ClipBot
   
   sngBlue = ClipLeft * .BottomRGBi.sBlu
   sngGreen = ClipLeft * .BottomRGBi.sGrn
   sngRed = ClipLeft * .BottomRGBi.sRed
   
   iiBlu2 = .LeftRGBi.sBlu + _
    (ClipLeft * .TopRGBi.sBlu - sngBlue) / (.Dims.HighM1)
   iiGrn2 = .LeftRGBi.sGrn + _
    (ClipLeft * .TopRGBi.sGrn - sngGreen) / (.Dims.HighM1)
   iiRed2 = .LeftRGBi.sRed + _
    (ClipLeft * .TopRGBi.sRed - sngRed) / (.Dims.HighM1)
   
  End With
  
  With FCG
  vertBlue = .LowLeft.Blue + sngBlue + ClipBot * iiBlu2 + 0.00001
  vertGreen = .LowLeft.Green + sngGreen + ClipBot * iiGrn2 + 0.00001
  vertRed = .LowLeft.Red + sngRed + ClipBot * iiRed2 + 0.00001
  End With
  
  For DrawY = DrawBot To DrawTop
   sngBlue = vertBlue
   sngGreen = vertGreen
   sngRed = vertRed
   vertBlue = vertBlue + iiBlu2
   vertGreen = vertGreen + iiGrn2
   vertRed = vertRed + iiRed2
   For DrawX = DrawLeft To DrawRight
    Surf.Dib(DrawX, DrawY).Blue = sngBlue
    Surf.Dib(DrawX, DrawY).Green = sngGreen
    Surf.Dib(DrawX, DrawY).Red = sngRed
    sngBlue = sngBlue + iBlue
    sngGreen = sngGreen + iGreen
    sngRed = sngRed + iRed
   Next DrawX
   iBlue = iBlue + iiBlu
   iGreen = iGreen + iiGrn
   iRed = iRed + iiRed
  Next DrawY
  End If

End Sub
Public Sub CopyEraseToMain(Surf As AnimSurf2D)
 For DrawX = 0 To Surf.UB1D
  Surf.LDib1D(DrawX) = Surf.EraseDib(DrawX)
 Next DrawX
End Sub
Public Sub CopyMainToErase(Surf As AnimSurf2D)
 For DrawY = 0 To Surf.UB1D
  Surf.EraseDib(DrawY) = Surf.LDib1D(DrawY)
 Next DrawY
End Sub
Public Sub SetCornerColor(TripleInt As RGBTripleInt, BGR1&)
 TripleInt.Blue = (BGR1 And vbBlue) / 65536
 TripleInt.Green = (BGR1 And &HFF00&) / 256
 TripleInt.Red = BGR1 And &HFF&
End Sub
