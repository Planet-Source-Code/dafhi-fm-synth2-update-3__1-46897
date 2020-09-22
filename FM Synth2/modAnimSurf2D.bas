Attribute VB_Name = "modAnimSurf2D"
Option Explicit

Private Type BGRAQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As BGRAQuad
End Type

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lpBits As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type PointApi
 x As Long
 y As Long
End Type

Private Type DimsAPI
 Wide As Long
 High As Long
 WideM1 As Long
 HighM1 As Long
End Type

Private Type PrecisionRGBA
 sRed As Single
 sGrn As Single
 sBlu As Single
 aLph As Single
End Type

Public Type AnimSurf2D
 Dims As DimsAPI            ' Width, Height, Width - 1, Height - 1
 halfW As Single            ' Width / 2
 halfH As Single            ' Height / 2
 TotalPixels As Long        ' Width * Height
 UB1D As Long               ' TotalPixels - 1
 TopLeft As Long
 SafeAry2D As SAFEARRAY2D
 SafeAry2D_L As SAFEARRAY2D
 SafeAry1D As SAFEARRAY1D
 SafeAry1D_L As SAFEARRAY1D
 mem_hDIb As Long
 mem_hBmpPrev As Long
 mem_hDC As Long
 BMinfo As BITMAPINFO
 Dib() As BGRAQuad           ' access 2D image as Bytes
 Dib1D() As BGRAQuad         ' 1D
 LDib() As Long             ' 2D Longs
 LDib1D() As Long           ' 1D
 EraseDib() As Long         ' Retain a copy of original
 sRGB() As PrecisionRGBA    ' high-depth color processing
 Verts() As Integer         ' advanced custom blit architecture
 HStack() As Integer        ' advanced custom blit architecture
 BackColor As Long
End Type

Public Sub MakeSurf(Surf As AnimSurf2D, ByVal LWidth&, ByVal LHeight&)
Dim MemBits&

 SetDims Surf, LWidth, LHeight

 If Surf.TotalPixels > 0 Then
 
    ClearSurface Surf
  
    With Surf.BMinfo.bmiHeader
        .biSize = Len(Surf.BMinfo.bmiHeader)
        .biWidth = Surf.Dims.Wide
        .biHeight = Surf.Dims.High
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = 4 * Surf.TotalPixels
    End With
    
    Surf.mem_hDC = CreateCompatibleDC(0)
    If (Surf.mem_hDC <> 0) Then
    Surf.mem_hDIb = CreateDIBSection(Surf.mem_hDC, Surf.BMinfo, _
            DIB_RGB_COLORS, _
            MemBits, _
            0, 0)
        If Surf.mem_hDIb <> 0 Then
            Surf.mem_hBmpPrev = SelectObject(Surf.mem_hDC, Surf.mem_hDIb)
        Else
            DeleteObject Surf.mem_hDC
            Surf.mem_hDC = 0
        End If
    End If
    
    SetSafeArrays Surf, MemBits
 
    Surf.UB1D = Surf.TotalPixels - 1
 
    ReDim Surf.EraseDib(Surf.UB1D)
  
 End If
 
End Sub
Private Sub SetDims(Surf As AnimSurf2D, Wide&, High&)

 Surf.Dims.Wide = Wide
 Surf.Dims.High = High

 Surf.Dims.WideM1 = Surf.Dims.Wide - 1
 Surf.Dims.HighM1 = Surf.Dims.High - 1
 
 Surf.halfW = Surf.Dims.Wide / 2&
 Surf.halfH = Surf.Dims.High / 2&
 
 Surf.TotalPixels = Surf.Dims.Wide * Surf.Dims.High
 Surf.UB1D = Surf.TotalPixels - 1
 
 Surf.TopLeft = Surf.TotalPixels - Surf.Dims.Wide
 
End Sub
Private Sub SetSafeArrays(Surf As AnimSurf2D, MemBits As Long, Optional BytesPixel As Byte = 4)

    With Surf.SafeAry2D
    .cbElements = BytesPixel
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(1).lLbound = 0
    .Bounds(0).cElements = Surf.Dims.High
    .Bounds(1).cElements = Surf.Dims.Wide
    .pvData = MemBits
    End With
    CopyMemory ByVal VarPtrArray(Surf.Dib), VarPtr(Surf.SafeAry2D), 4
 
    With Surf.SafeAry2D_L
    .cbElements = BytesPixel
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(1).lLbound = 0
    .Bounds(0).cElements = Surf.Dims.High
    .Bounds(1).cElements = Surf.Dims.Wide
    .pvData = MemBits
    End With
    CopyMemory ByVal VarPtrArray(Surf.LDib), VarPtr(Surf.SafeAry2D_L), 4
 
    With Surf.SafeAry1D
    .cbElements = BytesPixel
    .cDims = 1
    .lLbound = 0
    .cElements = Surf.TotalPixels
    .pvData = MemBits
    End With
    CopyMemory ByVal VarPtrArray(Surf.Dib1D), VarPtr(Surf.SafeAry1D), 4
  
    With Surf.SafeAry1D_L
    .cbElements = BytesPixel
    .cDims = 1
    .lLbound = 0
    .cElements = Surf.TotalPixels
    .pvData = MemBits
    End With
    CopyMemory ByVal VarPtrArray(Surf.LDib1D), VarPtr(Surf.SafeAry1D_L), 4
    
End Sub

Public Sub CreateSurfaceFromFile(Surf As AnimSurf2D, strFileName$, Optional DoMask As Boolean = True, Optional ByVal MaskColor = -1, Optional ByVal sRed! = 1!, Optional ByVal sGreen! = 1!, Optional ByVal sBlue! = 1!, Optional ByVal Alpha As Byte = 255)
Dim tBM As BITMAP, sPic As StdPicture
Dim CDC&, Loops&

    Set sPic = LoadPicture(strFileName)
    
    CDC = CreateCompatibleDC(0)           ' Temporary device
    DeleteObject SelectObject(CDC, sPic)  ' Converted bitmap
    
    GetObjectAPI sPic, Len(tBM), tBM
    
    MakeSurf Surf, tBM.bmWidth, tBM.bmHeight
    
    If Surf.TotalPixels > 0 Then
 
      BitBlt Surf.mem_hDC, 0, 0, tBM.bmWidth, tBM.bmHeight, _
             CDC, 0, 0, vbSrcCopy
             
      CreateMaskStructure Surf, DoMask, MaskColor, sRed, sGreen, sBlue, Alpha
             
    End If 'Surf.TotalPixels > 0
 
    DeleteDC CDC
 
End Sub
Private Sub CreateMaskStructure(Surf As AnimSurf2D, Optional DoMask As Boolean = True, Optional ByVal MaskColor = -1, Optional ByVal sRed! = 1!, Optional ByVal sGreen! = 1!, Optional ByVal sBlue! = 1!, Optional ByVal Alpha As Byte = 255)
Dim DrawX&, DrawY&
Dim BGRed&
Dim BGGrn&
Dim BGBlu&
Dim tsRed!
Dim tsGrn!
Dim tsBlu!
Dim tsMax!
Dim sAlpha!
Dim MR_X&
Dim MR_Blt&
Dim cBlt&
Dim HStack_Ptr&

    ReDim Surf.sRGB(Surf.Dims.WideM1, Surf.Dims.HighM1)
    ReDim Surf.Verts(Surf.Dims.HighM1)
    ReDim Surf.HStack(Surf.Dims.Wide, Surf.Dims.HighM1)
      
    For DrawY = 0 To Surf.Dims.HighM1
    For DrawX = 0 To Surf.Dims.WideM1
     Surf.Dib(DrawX, DrawY).Alpha = 0
    Next
    Next
      
    If Alpha > 0 Then
     tsMax = (sRed + sGreen + sBlue) * 255 / Alpha
    Else
     tsMax = 2550
    End If
    tsRed = sRed / tsMax
    tsGrn = sGreen / tsMax
    tsBlu = sBlue / tsMax
      
    If DoMask Then
      
     If MaskColor = -1 Then
      
     For DrawY = 0 To Surf.Dims.HighM1
      HStack_Ptr = -1
      MR_Blt = 0
      cBlt = -1
      For DrawX = 0 To Surf.Dims.WideM1
       BGRed = Surf.Dib(DrawX, DrawY).Red
       BGGrn = Surf.Dib(DrawX, DrawY).Green
       BGBlu = Surf.Dib(DrawX, DrawY).Blue
       Surf.sRGB(DrawX, DrawY).sRed = BGRed
       Surf.sRGB(DrawX, DrawY).sGrn = BGGrn
       Surf.sRGB(DrawX, DrawY).sBlu = BGBlu
       sAlpha = BGRed * tsRed + BGGrn * tsGrn + BGBlu * tsBlu
       sAlpha = sAlpha / 255
       Surf.sRGB(DrawX, DrawY).aLph = sAlpha
       If sAlpha < 0.002 Then
        If MR_Blt = 1 Then
         HStack_Ptr = HStack_Ptr + 1
         Surf.HStack(HStack_Ptr, DrawY) = cBlt
         cBlt = -1
         MR_Blt = 0
        End If
       Else 'sAlpha >= .002
        If MR_Blt = 0 Then
         HStack_Ptr = HStack_Ptr + 1
         Surf.HStack(HStack_Ptr, DrawY) = cBlt + 1
         cBlt = -1
         MR_Blt = 1
        End If
       End If
       cBlt = cBlt + 1
      Next DrawX
      Surf.Verts(DrawY) = HStack_Ptr
      If MR_Blt = 1 Then
       HStack_Ptr = HStack_Ptr + 1
       Surf.HStack(HStack_Ptr, DrawY) = cBlt
      End If
      Next DrawY
      
     Else 'MaskColor <> -1
      
      BGBlu = (MaskColor And vbRed) * 65536
      BGGrn = MaskColor And vbGreen
      BGRed = (MaskColor And vbBlue) / 65536
      MaskColor = BGRed Or BGGrn Or BGBlu
        
      For DrawY = 0 To Surf.Dims.HighM1
      HStack_Ptr = -1
      MR_Blt = 0
      cBlt = -1
      For DrawX = 0 To Surf.Dims.WideM1
       BGRed = Surf.Dib(DrawX, DrawY).Red
       BGGrn = Surf.Dib(DrawX, DrawY).Green
       BGBlu = Surf.Dib(DrawX, DrawY).Blue
       'precision color array
       Surf.sRGB(DrawX, DrawY).sRed = BGRed
       Surf.sRGB(DrawX, DrawY).sGrn = BGGrn
       Surf.sRGB(DrawX, DrawY).sBlu = BGBlu
       sAlpha = BGRed * tsRed + BGGrn * tsGrn + BGBlu * tsBlu
       sAlpha = sAlpha / 255
       Surf.sRGB(DrawX, DrawY).aLph = sAlpha
       If Surf.LDib(DrawX, DrawY) = MaskColor Then
        If MR_Blt = 1 Then
         HStack_Ptr = HStack_Ptr + 1
         Surf.HStack(HStack_Ptr, DrawY) = cBlt
         cBlt = -1
         MR_Blt = 0
        End If
       Else
        If MR_Blt = 0 Then
         HStack_Ptr = HStack_Ptr + 1
         Surf.HStack(HStack_Ptr, DrawY) = cBlt + 1
         cBlt = -1
         MR_Blt = 1
        End If
       End If
       cBlt = cBlt + 1
      Next DrawX
      Surf.Verts(DrawY) = HStack_Ptr
      If MR_Blt = 1 Then
       HStack_Ptr = HStack_Ptr + 1
       Surf.HStack(HStack_Ptr, DrawY) = cBlt
      End If
      Next DrawY
      
     End If 'MaskColor = -1
       
    Else 'DoMask = False
      
     tsMax = sRed
     If sGreen > tsMax Then tsMax = sGreen
     If sBlue > tsMax Then tsMax = sBlue
     tsRed = sRed / tsMax
     tsGrn = sGreen / tsMax
     tsBlu = sBlue / tsMax
     sAlpha = Alpha / 255
     For DrawY = 0 To Surf.Dims.HighM1
     For DrawX = 0 To Surf.Dims.WideM1
      BGRed = Surf.Dib(DrawX, DrawY).Red
      BGGrn = Surf.Dib(DrawX, DrawY).Green
      BGBlu = Surf.Dib(DrawX, DrawY).Blue
      BGRed = BGRed * tsRed
      BGGrn = BGGrn * tsGrn
      BGBlu = BGBlu * tsBlu
      Surf.sRGB(DrawX, DrawY).sRed = BGRed
      Surf.sRGB(DrawX, DrawY).sGrn = BGGrn
      Surf.sRGB(DrawX, DrawY).sBlu = BGBlu
      Surf.Dib(DrawX, DrawY).Red = BGRed
      Surf.Dib(DrawX, DrawY).Green = BGGrn
      Surf.Dib(DrawX, DrawY).Blue = BGBlu
      Surf.sRGB(DrawX, DrawY).aLph = sAlpha
     Next DrawX
     Surf.Verts(DrawY) = 0
     Surf.HStack(1, DrawY) = Surf.Dims.WideM1
     Next DrawY
       
    End If
      
End Sub
Public Sub CreateSurfaceFromBitmapStructure(Surf As AnimSurf2D, BM As BITMAP)

 ClearSurface Surf
 SetDims Surf, BM.bmWidth, BM.bmHeight
 SetSafeArrays Surf, BM.bmBits, (BM.bmBitsPixel / 8)
 CreateMaskStructure Surf

End Sub
'BitBlt wrapper
Public Sub BlitToDC(ByVal lHDC As Long, Surf As AnimSurf2D, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
        )

    If (lDestWidth < 0) Then lDestWidth = Surf.BMinfo.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = Surf.BMinfo.bmiHeader.biHeight
    
    BitBlt lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, Surf.mem_hDC, lSrcLeft, lSrcTop, eRop
    
End Sub
Public Sub Tile(SurfDest As AnimSurf2D, SurfSrc As AnimSurf2D, ByVal DestX&, ByVal DestY&)
Dim XLeft&
Dim XRight&
Dim YTop&
Dim YBot&
Dim DrawX&
Dim DrawY&

 XLeft = Int(DestX) Mod SurfSrc.Dims.Wide
 If XLeft > 0 Then XLeft = XLeft - SurfSrc.Dims.Wide
 
 YBot = Int(DestY) Mod SurfSrc.Dims.High
 If YBot > 0 Then YBot = YBot - SurfSrc.Dims.High
 
 XRight = SurfDest.Dims.WideM1
 YTop = SurfDest.Dims.HighM1
 
 For DrawY = YBot To YTop Step SurfSrc.Dims.High
  For DrawX = XLeft To XRight Step SurfSrc.Dims.Wide
   BitBlt SurfDest.mem_hDC, DrawX, DrawY, _
               SurfSrc.Dims.Wide, SurfSrc.Dims.High, _
                  SurfSrc.mem_hDC, 0, 0, vbSrcCopy
  Next
 Next

End Sub
Public Sub SolidColorFill(Surf As AnimSurf2D, Red As Byte, Green As Byte, Blue As Byte)
Dim BGR&
Dim DrawX&
Dim DrawY&

 BGR = RGB(Blue, Green, Red)
 With Surf
 If .TotalPixels > 0 Then
   For DrawY = 0 To .Dims.HighM1
    For DrawX = 0 To .Dims.WideM1
    
    ''LDib is the array that bitblt copies from and blits to
    .LDib(DrawX, DrawY) = BGR
    
    Next
   Next
 End If
 End With
 
End Sub
Public Sub ClearSurface(Surf As AnimSurf2D)

 'erase pointers
 
 CopyMemory ByVal VarPtrArray(Surf.Dib), 0&, 4
 CopyMemory ByVal VarPtrArray(Surf.LDib), 0&, 4
 CopyMemory ByVal VarPtrArray(Surf.Dib1D), 0&, 4
 CopyMemory ByVal VarPtrArray(Surf.LDib1D), 0&, 4
 
 With Surf
     If (.mem_hDC <> 0) Then
         If (.mem_hDIb <> 0) Then
             SelectObject .mem_hDC, .mem_hBmpPrev
             DeleteObject .mem_hDIb
         End If
         DeleteObject .mem_hDC
     End If
     DeleteDC .mem_hDC
     .mem_hDC = 0: .mem_hDIb = 0: .mem_hBmpPrev = 0 ': .mem_Bits = 0
 End With
 
 Erase Surf.EraseDib
 Erase Surf.sRGB
 Erase Surf.Verts
 Erase Surf.HStack

End Sub

Public Sub AlphaBlit(SurfDest As AnimSurf2D, SurfSrc As AnimSurf2D, px!, py!, Optional DoTile As Boolean)
Dim DrawLeft& 'Destination
Dim DrawRight&
Dim DrawTop&
Dim DrawBot&
Dim DrawX&
Dim DrawY&
Dim BGColor&
Dim BGGrn&
Dim BGBlu&

Dim AddDrawWidth&
Dim SrcClipLeft&
Dim SrcClipBot&
Dim SrX&
Dim SrY&
Dim sngAlpha!
Dim FGColor&
Dim FGGrn&
Dim FGBlu&

Dim StartX& 'Tile
Dim StartY&
Dim EndX&
Dim EndY&
Dim LngX&
Dim LngY&

Dim H_Segs& 'advanced mask architecture
Dim SegsL&
Dim SLX&
Dim SLE&
Dim TSL&
Dim RB&

 If SurfDest.TotalPixels > 0 Then
 
 If DoTile Then
  StartX = Int(px - SurfSrc.halfW) Mod SurfSrc.Dims.Wide - SurfSrc.Dims.Wide
  StartY = Int(py - SurfSrc.halfH) Mod SurfSrc.Dims.High - SurfSrc.Dims.High
  EndX = SurfDest.Dims.Wide
  EndY = SurfDest.Dims.High
 Else
  StartX = Int(px - SurfSrc.halfW)
  StartY = Int(py - SurfSrc.halfH)
  EndX = StartX
  EndY = StartY
 End If
 
 For LngY = StartY To EndY Step SurfSrc.Dims.High
 For LngX = StartX To EndX Step SurfSrc.Dims.Wide
 
 DrawLeft = LngX
 DrawRight = DrawLeft + SurfSrc.Dims.WideM1
 DrawBot = LngY
 DrawTop = DrawBot + SurfSrc.Dims.HighM1
 
 If DrawLeft < 0 Then
  SrcClipLeft = -DrawLeft
  DrawLeft = 0
 Else
  SrcClipLeft = 0
 End If
 If DrawRight > SurfDest.Dims.WideM1 Then
  DrawRight = SurfDest.Dims.WideM1
 End If
 
 If DrawBot < 0 Then
  SrcClipBot = -DrawBot
  DrawBot = DrawLeft
 Else
  SrcClipBot = 0
  DrawBot = SurfDest.Dims.Wide * DrawBot + DrawLeft
 End If
 If DrawTop > SurfDest.Dims.HighM1 Then
  DrawTop = SurfDest.Dims.HighM1
 End If
 
 DrawTop = SurfDest.Dims.Wide * DrawTop + DrawLeft
 
 AddDrawWidth = DrawRight - DrawLeft
 
 SrY = SrcClipBot
 
 RB = SrcClipLeft + AddDrawWidth
 
 For DrawY = DrawBot To DrawTop Step SurfDest.Dims.Wide
  H_Segs = SurfSrc.Verts(SrY)
  TSL = 0
  DrawX = DrawY
  For SegsL = 0 To H_Segs Step 2
   SLE = TSL
   SrX = SurfSrc.HStack(SegsL, SrY)
   SLX = SLE + SrX 'draw start
   If SLX > SrcClipLeft Then
    If SLE <= SrcClipLeft Then
     DrawX = DrawX + SLX - SrcClipLeft
    Else
     DrawX = DrawX + SrX 'slx - sle
    End If
    SLE = SLX + SurfSrc.HStack(SegsL + 1, SrY)
   Else
    SLE = SLX + SurfSrc.HStack(SegsL + 1, SrY)
    SLX = SrcClipLeft
   End If
   TSL = SLE + 1
   If SLE > RB Then SLE = RB
   For SrX = SLX To SLE
    FGColor = SurfSrc.LDib(SrX, SrY)
    BGColor = SurfDest.LDib1D(DrawX)
    BGGrn = BGColor And vbGreen
    FGGrn = FGColor And vbGreen
    BGBlu = BGColor And &HFF&
    FGBlu = FGColor And &HFF&
    sngAlpha = SurfSrc.sRGB(SrX, SrY).aLph
    SurfDest.LDib1D(DrawX) = _
        BGColor + sngAlpha * (FGColor - BGColor) And vbBlue Or _
        BGGrn + sngAlpha * (FGGrn - BGGrn) And vbGreen Or _
        BGBlu + sngAlpha * (FGBlu - BGBlu) And &HFF&
    DrawX = DrawX + 1
   Next SrX
  Next SegsL
  SrY = SrY + 1
 Next DrawY
 
 Next LngX
 Next LngY
  
 End If 'Surf.TotalPixels > 0
  
End Sub
Public Sub RGBBlit(SurfDest As AnimSurf2D, SurfSrc As AnimSurf2D, px!, py!, Optional DoTile As Boolean, Optional ByVal AlphaRed As Byte = 255, Optional ByVal AlphaGreen As Byte = 255, Optional ByVal AlphaBlue As Byte = 255, Optional ByVal Alpha As Byte = 255)
Dim DrawLeft& 'Destination
Dim DrawRight&
Dim DrawTop&
Dim DrawBot&
Dim DrawX&
Dim DrawY&
Dim BGRed&
Dim BGGrn&
Dim BGBlu&

Dim AddDrawWidth&
Dim SrcClipLeft&
Dim SrcClipBot&
Dim SrX&
Dim SrY&
Dim BGColor&
Dim FGColor&
Dim sngAlpha!
Dim sngAlpha2!
Dim sFGRed!
Dim sFGGrn!
Dim sFGBlu!

Dim StartX& 'Tile
Dim StartY&
Dim EndX&
Dim EndY&
Dim LngX&
Dim LngY&

Dim H_Segs& 'advanced mask architecture
Dim SegsL&
Dim SLX&
Dim SLE&
Dim TSL&
Dim RB&

 If SurfDest.TotalPixels > 0 Then
 
 If DoTile Then
  StartX = Int(px - SurfSrc.halfW) Mod SurfSrc.Dims.Wide - SurfSrc.Dims.Wide
  StartY = Int(py - SurfSrc.halfH) Mod SurfSrc.Dims.High - SurfSrc.Dims.High
  EndX = SurfDest.Dims.Wide
  EndY = SurfDest.Dims.High
 Else
  StartX = Int(px - SurfSrc.halfW)
  StartY = Int(py - SurfSrc.halfH)
  EndX = StartX
  EndY = StartY
 End If
 
 For LngY = StartY To EndY Step SurfSrc.Dims.High
 For LngX = StartX To EndX Step SurfSrc.Dims.Wide
 
 DrawLeft = LngX
 DrawRight = DrawLeft + SurfSrc.Dims.WideM1
 DrawBot = LngY
 DrawTop = DrawBot + SurfSrc.Dims.HighM1
 
 If DrawLeft < 0 Then
  SrcClipLeft = -DrawLeft
  DrawLeft = 0
 Else
  SrcClipLeft = 0
 End If
 If DrawRight > SurfDest.Dims.WideM1 Then
  DrawRight = SurfDest.Dims.WideM1
 End If
 
 If DrawBot < 0 Then
  SrcClipBot = -DrawBot
  DrawBot = DrawLeft
 Else
  SrcClipBot = 0
  DrawBot = SurfDest.Dims.Wide * DrawBot + DrawLeft
 End If
 If DrawTop > SurfDest.Dims.HighM1 Then
  DrawTop = SurfDest.Dims.HighM1
 End If
 
 DrawTop = SurfDest.Dims.Wide * DrawTop + DrawLeft
 
 AddDrawWidth = DrawRight - DrawLeft
 
 SrY = SrcClipBot
 
 RB = SrcClipLeft + AddDrawWidth
 
 If Alpha = 255 Then
 
  If AlphaRed = 255 And AlphaGreen = 255 And AlphaBlue = 255 Then
  
  For DrawY = DrawBot To DrawTop Step SurfDest.Dims.Wide
   H_Segs = SurfSrc.Verts(SrY)
   TSL = 0
   DrawX = DrawY
   For SegsL = 0 To H_Segs Step 2
    SLE = TSL
    SrX = SurfSrc.HStack(SegsL, SrY)
    SLX = SLE + SrX 'draw start
    If SLX > SrcClipLeft Then
     If SLE <= SrcClipLeft Then
      DrawX = DrawX + SLX - SrcClipLeft
     Else
      DrawX = DrawX + SrX
     End If
     SLE = SLX + SurfSrc.HStack(SegsL + 1, SrY)
    Else
     SLE = SLX + SurfSrc.HStack(SegsL + 1, SrY)
     SLX = SrcClipLeft
    End If
    TSL = SLE + 1
    If SLE > RB Then SLE = RB
    For SrX = SLX To SLE
     SurfDest.LDib1D(DrawX) = SurfSrc.LDib(SrX, SrY)
     DrawX = DrawX + 1
    Next SrX
   Next SegsL
   SrY = SrY + 1
  Next DrawY
  
  Else 'Red, Green, or Blue < 255, Alpha = 255
  
  sFGBlu = AlphaBlue / 255
  sFGGrn = AlphaGreen / 255
  sFGRed = AlphaRed / 255
  
  For DrawY = DrawBot To DrawTop Step SurfDest.Dims.Wide
   H_Segs = SurfSrc.Verts(SrY)
   TSL = 0
   DrawX = DrawY
   For SegsL = 0 To H_Segs Step 2
    SLE = TSL
    SrX = SurfSrc.HStack(SegsL, SrY)
    SLX = SLE + SrX 'draw start
    If SLX > SrcClipLeft Then
     If SLE <= SrcClipLeft Then
      DrawX = DrawX + SLX - SrcClipLeft
     Else
      DrawX = DrawX + SrX
     End If
     SLE = SLX + SurfSrc.HStack(SegsL + 1, SrY)
    Else
     SLE = SLX + SurfSrc.HStack(SegsL + 1, SrY)
     SLX = SrcClipLeft
    End If
    TSL = SLE + 1
    If SLE > RB Then SLE = RB
    For SrX = SLX To SLE
     FGColor = SurfSrc.LDib(SrX, SrY)
     SurfDest.LDib1D(DrawX) = _
      (FGColor And &HFF) * sFGBlu Or _
      (FGColor And &HFF00&) * sFGGrn And &HFF00 Or _
      (FGColor And &HFF0000) * sFGRed And &HFF0000
     DrawX = DrawX + 1
    Next SrX
   Next SegsL
   SrY = SrY + 1
  Next DrawY
  
  End If 'Red, Green and Blue alphas = 255
  
 Else 'Alpha < 255
 
  sngAlpha = Alpha / 255
  sngAlpha2 = 1 - sngAlpha
  
  If AlphaRed = 255 And AlphaGreen = 255 And AlphaBlue = 255 Then
   sFGRed = sngAlpha
   sFGGrn = sngAlpha
   sFGBlu = sngAlpha
  Else 'Red, Green, or Blue alpha < 255
   sFGRed = sngAlpha * AlphaRed / 255
   sFGGrn = sngAlpha * AlphaGreen / 255
   sFGBlu = sngAlpha * AlphaBlue / 255
  End If

  For DrawY = DrawBot To DrawTop Step SurfDest.Dims.Wide
   H_Segs = SurfSrc.Verts(SrY)
   TSL = 0
   DrawX = DrawY
   For SegsL = 0 To H_Segs Step 2
    SLE = TSL
    SrX = SurfSrc.HStack(SegsL, SrY)
    SLX = SLE + SrX 'draw start
    If SLX > SrcClipLeft Then
     If SLE <= SrcClipLeft Then
      DrawX = DrawX + SLX - SrcClipLeft
     Else
      DrawX = DrawX + SrX
     End If
     SLE = SLX + SurfSrc.HStack(SegsL + 1, SrY)
    Else
     SLE = SLX + SurfSrc.HStack(SegsL + 1, SrY)
     SLX = SrcClipLeft
    End If
    TSL = SLE + 1
    If SLE > RB Then SLE = RB
    For SrX = SLX To SLE
     FGColor = SurfSrc.LDib(SrX, SrY)
     BGColor = SurfDest.LDib1D(DrawX)
     SurfDest.LDib1D(DrawX) = _
      (FGColor And &HFF) * sFGBlu + (BGColor And &HFF) * sngAlpha2 Or _
      (FGColor And &HFF00&) * sFGGrn + (BGColor And &HFF00&) * sngAlpha2 And &HFF00 Or _
      (FGColor And &HFF0000) * sFGRed + (BGColor And &HFF0000) * sngAlpha2 And &HFF0000
     DrawX = DrawX + 1
    Next SrX
   Next SegsL
   SrY = SrY + 1
  Next DrawY
  
 End If 'Alpha = 255
 
 Next LngX
 Next LngY
  
 End If 'Surf.TotalPixels > 0
  
End Sub


