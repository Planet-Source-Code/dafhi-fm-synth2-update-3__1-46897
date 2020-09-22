Attribute VB_Name = "modWaveforms"
Option Explicit

Public Type Pos_And_Incr
 sPos      As Double
 iPos      As Double
End Type

Public Type FM_Note
 BaseFreq     As Pos_And_Incr
 BaseFM       As Pos_And_Incr
 Freq2        As Pos_And_Incr
 FM2          As Pos_And_Incr
 EnvelopeMode As Long
 cSamplesToEnvelopeTarget As Long
 sAmpEnv    As Double
 iAmpEnv    As Double
 BufStart   As Long
 BufStop    As Long
 AddToStart As Long
 ElemRef    As Long
 KeyRef     As Long
 'sVol       As Single
 HalfToneOffset As Long
 'KeyIndex  As Long
 'ImmediateWrite As Boolean
 'KeyDown        As Boolean
End Type

Public Const pi As Double = 3.14159265358979
Public Const TwoPi As Double = 2& * pi

Public Const NOTE_1OF12 As Double = 2 ^ (1 / 12)
Public Const NOTE_2OF12 As Double = 2 ^ (2 / 12)
Public Const NOTE_3OF12 As Double = 2 ^ (3 / 12)
Public Const NOTE_4OF12 As Double = 2 ^ (4 / 12)
Public Const NOTE_5OF12 As Double = 2 ^ (5 / 12)
Public Const NOTE_6OF12 As Double = 2 ^ (6 / 12)
Public Const NOTE_7OF12 As Double = 2 ^ (7 / 12)
Public Const NOTE_8OF12 As Double = 2 ^ (8 / 12)
Public Const NOTE_9OF12 As Double = 2 ^ (9 / 12)
Public Const NOTE_10OF12 As Double = 2 ^ (10 / 12)
Public Const NOTE_11OF12 As Double = 2 ^ (11 / 12)

Public Type SawToothStruct_Positive_Increasing
 currentV As Double
 maxV As Single
 minV As Single
 iStep As Double
End Type

Public Function SawToothPositive(SawWavePI As SawToothStruct_Positive_Increasing)
 SawWavePI.currentV = SawWavePI.currentV + SawWavePI.iStep
 If SawWavePI.currentV > SawWavePI.maxV! Then
  SawWavePI.currentV = SawWavePI.minV + SawWavePI.currentV - SawWavePI.maxV
 End If
 SawToothPositive = SawWavePI.currentV
End Function
Public Function Saw(Phh#) As Double   ' Sawtooth
 If Phh < 0.5! Then
   Saw = 0.35! * Sin(Phh)
 Else
   Saw = (0.35! * Sin(1! - Phh)) + 0.5!
 End If
End Function

