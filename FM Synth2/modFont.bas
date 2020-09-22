Attribute VB_Name = "modFont"
'Get Path Demo
'by Scythe
'www.scythe-tools.de

'This is not the best code
'but it seems to be the only example
'existing for GetPath and Visual basic
'(I searched but i dont find any others)

'Thanks to Charles P.V.
'for the idea to create a vector font thru Path

Option Explicit

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPath Lib "gdi32" (ByVal hdc As Long, lpPoint As PointApi, lpTypes As Byte, ByVal nSize As Long) As Long
Private Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As Long, lppt As PointApi, ByVal cCount As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'Private Declare Function FlattenPath Lib "gdi32" (ByVal hdc As Long) As Long
'If you dont want to use Bezier then use this api to
'convert all curves in the path into lines
'this shound be easyer for 3D text

Private Type PointApi
 x As Long
 y As Long
End Type
Public Sub DrawText2(Surf As AnimSurf2D, ByVal x%, ByVal y%, StrText As String, FontColor As Long, Optional ByVal FontSize As Integer = 10)
 TextOut Surf.mem_hDC, x, y, StrText, FontSize
End Sub
Public Sub DrawText(Surf As AnimSurf2D, ByVal x%, ByVal y%, KnobA1 As KnobAlpha1)
 Dim Pth(1000) As PointApi   'This will hold the Path
 Dim Tpe(1000) As Byte       'What is it (Bezier or Line)
 Dim Pbz(1000) As PointApi   'Holds the data for the Bezier
 Dim BzCtr As Long
 Dim Ctr As Long
 Dim I As Long
 Dim NewFig As Long
 Dim Pic1 As PictureBox
 Dim SizeRatio&
 
 Set Pic1 = frmMain.Picture1
 
 Airbrush.Red = KnobA1.TextColor And &HFF&
 Airbrush.Green = (KnobA1.TextColor And &HFF00&) / 256&
 Airbrush.Blue = (KnobA1.TextColor And vbBlue) / 65536
 
 Airbrush.Pressure = 255
 Airbrush.diameter = 1.6
 Airbrush.definition = 2.3

 'Lets create a path
 BeginPath Pic1.hdc
 
 SizeRatio = 50

 'Do something to fill our Path
 'You wont see this
 Pic1.ScaleMode = vbPixels
 Pic1.CurrentX = x * SizeRatio
 Pic1.CurrentY = y * SizeRatio
 Pic1.FontSize = KnobA1.FontSize * SizeRatio
 Pic1.Font = "Comic Sans MS"
 Pic1.Print KnobA1.StrText

 'End the Path
 EndPath Pic1.hdc
 
 'Now get the Path
 Ctr = GetPath(Pic1.hdc, Pth(0), Tpe(0), 1000)
 For I = 0 To Ctr
  Pth(I).y = Surf.Dims.High * SizeRatio - Pth(I).y '- SizeRatio * (FontSize)
 Next I

 Ctr = Ctr - 1

 'Now draw Path
 For I = 0 To Ctr
  
  Select Case Tpe(I)

  Case 6 'Start a new figure
   'Set the Starting point
   'Pset and currentx wont work (dont know why)
   'Pic1.Line (Pth(I).X, Pth(I).Y)-(Pth(I).X, Pth(I).Y)
   BlitAirbrush Surf, Pth(I).x / SizeRatio, Pth(I).y / SizeRatio
   'Hold the startpoit = endpoint
   NewFig = I

  Case 5 'end of a Bezier
   'Now we must increase the Bezier counter
   'Pbz(BzCtr).X = Pth(I).X
   'Pbz(BzCtr).Y = Pth(I).Y
   BzCtr = BzCtr + 1
   ''Pbz(BzCtr).X = Pth(i).X
   ''Pbz(BzCtr).Y = Pth(i).Y
   'Draw the Bezier
   'Pic1.PSet (Pth(I).X, Pth(I).Y), vbBlue
   BlitAirbrush Surf, Pth(I).x / SizeRatio, Pth(I).y / SizeRatio
   ''PolyBezierTo Pic.hdc, Pbz(0), BzCtr
   'Reset Counter
   BzCtr = 0
   'Close the figure
   'Pic1.Line (Pth(I).X, Pth(I).Y)-(Pth(NewFig).X, Pth(NewFig).Y)
  
  Case 3 'end as line
   'if we have an bezier open then draw it
   If BzCtr > 0 Then
    'Set the last bezier Point
    'Pbz(BzCtr).X = Pth(I).X
    'Pbz(BzCtr).Y = Pth(I).Y
    'Pic1.PSet (Pth(I).X, Pth(I).Y), vbBlue
    BlitAirbrush Surf, Pth(I).x / SizeRatio, Pth(I).y / SizeRatio
    ''PolyBezierTo Pic.hdc, Pbz(0), BzCtr
    'Set the current X,Y
    'Pic1.PSet (Pth(I - 1).X, Pth(I - 1).Y)
    BlitAirbrush Surf, Pth(I - 1).x / SizeRatio, Pth(I - 1).y / SizeRatio
    BzCtr = 0
   End If
   'Draw last line
   'Pic1.Line -(Pth(I).X, Pth(I).Y)
   BlitAirbrush Surf, Pth(I).x / SizeRatio, Pth(I).y / SizeRatio
   'Close the figure
   'Pic1.Line (Pth(I).X, Pth(I).Y)-(Pth(NewFig).X, Pth(NewFig).Y)

  Case 4 'Bezier
   'Pbz(BzCtr).X = Pth(I).X
   'Pbz(BzCtr).Y = Pth(I).Y
   'Pic1.PSet (Pth(I).X, Pth(I).Y), vbBlue
   BlitAirbrush Surf, Pth(I).x / SizeRatio, Pth(I).y / SizeRatio
   BzCtr = BzCtr + 1

  Case Else 'Line
   'if we have an bezier open then draw it
   If BzCtr > 0 Then
    'Set the last bezier Point
    'Pbz(BzCtr).X = Pth(I).X
    'Pbz(BzCtr).Y = Pth(I).Y
    'Pic1.PSet (Pth(I).X, Pth(I).Y), vbBlue
    BlitAirbrush Surf, Pth(I).x / SizeRatio, Pth(I).y / SizeRatio
    ''PolyBezierTo Pic.hdc, Pbz(0), BzCtr
    'Set the current X,Y
    'Pic1.PSet (Pth(I - 1).X, Pth(I - 1).Y)
   BlitAirbrush Surf, Pth(I - 1).x / SizeRatio, Pth(I - 1).y / SizeRatio
    BzCtr = 0
   End If
   'Draw line
   'Pic1.Line -(Pth(I).X, Pth(I).Y)

  End Select

 Next I
 
 Set Pic1 = Nothing

End Sub


