VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Oscillator"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'DirectSound and Visual Basic - FM Synth 2 update 3
'real-time 16-bit FM synthesizer by dafhi

' - instructions -

'"Play" computer keyboard like a piano.
'Mouse and Shift-Mouse adjust the controls.
'Plus and Minus (both sets) adjust polyphony.

'Adjustable SampleRate (below) is not guaranteed
'to work with all sound cards.

Private Const SampleRate = 44100 'raise BufUBound if SampleRate
                                 'crunches signal.

Private Const BufUBound As Long = 1799
Private Const BufferSamples As Long = BufUBound + 1

'output signal array gets copied to sound buffer
Dim lpBuffer() As Integer

'rendering signal
Dim sRender() As Single

Dim DX7 As New DirectX7
Dim DS As DirectSound
Dim DSB As DirectSoundBuffer
Dim DSBD As DSBUFFERDESC
Dim PCM As WAVEFORMATEX


'=== behind the scenes

'Position in the play buffer
Dim Pos As DSCURSORS
Dim PlayPos As Long
Dim PreviousPlayPos As Long

Dim reference_iFreq As Double
Dim sFreq As Double

Dim MainCalcStart&

Dim SampleRunM1&

Dim AverageSampleRun&

Dim I&
Dim J&

Dim TmpBufStop&

Dim TmpStr As String

Const BYT2 As Byte = 2
Const BYT64 As Byte = 64
Const BYT128 As Byte = 128

Const GrayScaleRGB As Long = 1 + 256 + 65536

Const EnvelopeMode_ATTACK As Long = 1
Const EnvelopeMode_SUSTAIN As Long = 2
Const EnvelopeMode_RELEASE As Long = 3

Const CONST_MAX_VERTICAL As Integer = 32767
Const CONST_MIN_VERTICAL As Integer = -32768

Const CONST_HALF_VERTICAL_RESOLUTION = CONST_MAX_VERTICAL / 2

Const MaxSecondsForEnvelope As Long = 5

Const MaxPolyphony As Long = 256

Dim Knob_FM_Freq_Ratio As KnobAlpha1
Dim Knob_FM_Ampl       As KnobAlpha1
Dim KnobPitchLayer2   As KnobAlpha1
Dim KnobVolume      As KnobAlpha1
Dim Knob_GL_Pitch   As KnobAlpha1
Dim Knob_GL_Attack  As KnobAlpha1
Dim Knob_GL_Release As KnobAlpha1

Private InfoBoard As AnimSurf2D

Private iAttackRate As Single
Private iReleaseRate As Single

Private Notes As DataLinkType

Private Polyphony As Long

'Private Type Patch
' sAttack As Single
' sRelease As Single
' sFMVal As Single
' sFMAmp As Single
 'sVol As Single
' HalfToneOffset As Long
'End Type

'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Sub Form_Load()
Dim Wide%
Dim High%
Dim dot_size_percent!
Dim knob_size_percent!
Dim MouseStepsPerIncrement As Byte
Dim TopLeftColor&
Dim TopRightColor&
Dim BottomLeftColor&
Dim BottomRightColor&
Dim KnobColor&
Dim DotColor&
 
InitializeStack Notes, MaxPolyphony

Polyphony = 10

reference_iFreq = 220 * NOTE_3OF12 * TwoPi

ForeColor = vbWhite
ScaleMode = vbPixels

'useful for debugging
MakeSurf InfoBoard, 70, 54

FCG_ColorTopLeft FourColorGradient, 255, 55, 255
FCG_ColorTopRight FourColorGradient, 120, 120, 120
WrapFCG InfoBoard, FourColorGradient

' === Control Properties

Randomize

TopLeftColor = RGB(120, 88, 60)
TopRightColor = RGB(155, 155, 155)
BottomLeftColor = 70 * GrayScaleRGB
BottomRightColor = RGB(55, 55, 55)
KnobColor = vbBlack
DotColor = RGB(255, 92, 0)
 
Knob_FM_Freq_Ratio.TextColor = vbBlue
Knob_GL_Attack.TextColor = RGB(18, 68, 45)
Knob_GL_Release.TextColor = RGB(65, 32, 0)
Knob_FM_Ampl.TextColor = RGB(35, 0, 155)
KnobPitchLayer2.TextColor = RGB(165, 152, 140)
KnobVolume.TextColor = RGB(168, 150, 140)
Knob_GL_Pitch.TextColor = RGB(255, 0, 138)

Knob_FM_Freq_Ratio.vVal = 0
Knob_FM_Ampl.vVal = 1.5
Knob_GL_Attack.vVal = 0 * SampleRate ' 1 * SampleRate will give 1 second.
Knob_GL_Release.vVal = 0.66 * SampleRate
KnobPitchLayer2.vVal = 1
KnobVolume.vVal = 0.1

Wide = 85
High = 42

knob_size_percent = 0.8
dot_size_percent = 0.26

'FM Freq
InitializeKnob Knob_FM_Freq_Ratio, 3, 3, Wide, High, _
  0, 5, 300, , _
  knob_size_percent, dot_size_percent, "FM Frequency", 9, TopLeftColor, TopRightColor, BottomLeftColor, BottomRightColor, KnobColor, DotColor

'FM Amp
InitializeKnob Knob_FM_Ampl, _
 92, 3, Wide + 6, High, _
  0, 1.5, 100, 2, _
  knob_size_percent, dot_size_percent, "FM Amplify", 9, TopLeftColor, TopRightColor, BottomLeftColor, BottomRightColor, KnobColor, DotColor
                                      'True = disable vVal from being set
                                      'to vMin

'Note Attack
InitializeKnob Knob_GL_Attack, 3, 47, Wide, High + 1, _
  0, SampleRate * MaxSecondsForEnvelope, 450, 1, _
  knob_size_percent, dot_size_percent, "Attack", 10, TopLeftColor, TopRightColor, BottomLeftColor, BottomRightColor, KnobColor, DotColor
  
iAttackRate = 1 / (Knob_GL_Attack.vVal + 1)

'Note Release
InitializeKnob Knob_GL_Release, 3, 92, Wide, High + 1, _
  0, SampleRate * MaxSecondsForEnvelope, 450, 1, _
  knob_size_percent, dot_size_percent, "Release", 10, TopLeftColor, TopRightColor, BottomLeftColor, BottomRightColor, KnobColor, DotColor

iReleaseRate = -1 / (Knob_GL_Release.vVal + 1)

Wide = Wide + 6

'Pitch Layer 2
InitializeKnob KnobPitchLayer2, 187, 3, Wide, High, _
  1, NOTE_7OF12, 295, , _
  knob_size_percent, dot_size_percent, "Pitch Layer 2", 10, TopLeftColor, TopRightColor, BottomLeftColor, BottomRightColor, KnobColor, DotColor
  
High = 88
  
'Volume
InitializeKnob KnobVolume, 92, 47, Wide, High, _
  0, 1, 120, , _
  knob_size_percent, dot_size_percent, "Volume", 12, TopLeftColor, TopRightColor, BottomLeftColor, BottomRightColor, KnobColor, DotColor, True

'Global Pitch
InitializeKnob Knob_GL_Pitch, 187, 47, Wide, High, _
  -24, 24, 48, 9, _
  knob_size_percent, dot_size_percent, "Global Pitch", 10, TopLeftColor, TopRightColor, BottomLeftColor, BottomRightColor, KnobColor, DotColor, True
  

' === DirectSound

Set DS = DX7.DirectSoundCreate(vbNullString)
DS.SetCooperativeLevel hWnd, DSSCL_NORMAL

PCM.nFormatTag = WAVE_FORMAT_PCM
PCM.nChannels = 1
PCM.lSamplesPerSec = SampleRate
PCM.nBitsPerSample = 16
PCM.nBlockAlign = PCM.nChannels * PCM.nBitsPerSample / 8
PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign

DSBD.LFlags = DSBCAPS_STATIC

DSBD.lBufferBytes = 2 * (BufUBound + 1)
Set DSB = DS.CreateSoundBuffer(DSBD, PCM)
ReDim lpBuffer(BufUBound)
ReDim sRender(BufUBound)

InitKeys True

Show

''This defaults QWERT to sound like CDEFG
sFreq = reference_iFreq / SampleRate
  
DSB.Play DSBPLAY_LOOPING
 
Do While DoEvents
 SoundBuffer
 I = ScaleHeight - 14
 CurrentX = 0
 CurrentY = I
 'BlitToDC Me.hdc, InfoBoard, 0, I
 'Print "Polyphony: " & (Notes.InUsePtr - Notes.LB + 1)
 'Print SampleRunM1, TmpBufStart - PlayPos
Loop

DSB.Stop

Set DSB = Nothing
Set DS = Nothing

End Sub

'With Max Polyphony, it is common to replace the most distant note
'in time with the one you want to add.
Private Sub Form_KeyDown(IntKey As Integer, Shift As Integer)

 Select Case IntKey
 
 Case vbKeyEscape
  Unload Me
  
 Case 107 'big plus key
  If Polyphony < MaxPolyphony Then
   Polyphony = Polyphony + 1
   ClearStack Notes
   InitializeStack Notes, Polyphony
  End If
  Caption = "Polyphony: " & Polyphony
  
 Case 109 'minus key above big plus
  If Polyphony > 1 Then
   Polyphony = Polyphony - 1
   ClearStack Notes
   InitializeStack Notes, Polyphony
  End If
  Caption = "Polyphony: " & Polyphony
  
 Case vbKeySpace
  ClearStack Notes
  
 Case Else
  'prevent key repeat
  If Not KeyPressed(IntKey) Then
   AddNote IntKey
   KeyPressed(IntKey) = True
  End If
  
 End Select
 
End Sub
Private Sub AddNote(KeyReference As Integer)
 
 If GetNoteIndexFromKey(KeyReference) > 0 Then
  If modHybridLink.AddLink(Notes) Then 'open spot in Notes stack
   InitializeNoteInLink Notes, KeyReference, Notes.Elem_Tail
  Else 'Notes stack is full
   RemoveLink Notes, Notes.Elem_Head
   AddLink Notes
   InitializeNoteInLink Notes, KeyReference, Notes.Elem_Tail
  End If
 End If
 
End Sub
Private Sub Form_KeyUp(IntKey As Integer, Shift As Integer)
 KeyPressed(IntKey) = False
 I = Notes.GetStackElemFromKey(IntKey)
 If I <> 0 Then
  TargetNoteRelease Notes.Stack(I).Whatever
 End If
End Sub



Private Sub SoundBuffer()

DSB.GetCurrentPosition Pos
PlayPos = Pos.lPlay / PCM.nBlockAlign

If PlayPos > BufUBound Then PlayPos = PlayPos - BufferSamples

If PreviousPlayPos < PlayPos Then
 MapSignal MainCalcStart, PlayPos - 1
 MainCalcStart = PlayPos
ElseIf PreviousPlayPos > PlayPos Then
 MapSignal MainCalcStart, BufUBound
 MainCalcStart = 0
End If

PreviousPlayPos = PlayPos

End Sub
Private Sub MapSignal(ByVal BufStart&, ByVal BufStop&)
Dim SampleRun&

 'zero the render signal
 For I = BufStart& To BufStop
  sRender(I) = 0!
 Next I
 
 SampleRun = BufStop - BufStart + 1
 
 AverageSampleRun = (AverageSampleRun + SampleRun) / 2&
 
 Render_HiFi Notes, BufStart, SampleRun
 
End Sub
Private Sub Render_HiFi(NotesDLT As DataLinkType, BufStart&, SampleRun&)
Dim I2&
Dim J2&
Dim K2&

 If BufStart > BufUBound Then BufStart = BufStart - BufferSamples
 
 If SampleRun > BufferSamples Then SampleRun = BufferSamples
 SampleRunM1 = SampleRun - 1
 
 J2 = NotesDLT.Elem_Head
 K2 = NotesDLT.InUsePtr
 For I2 = NotesDLT.LB To K2
  WorkEnvelopPoints NotesDLT.Stack(J2).Whatever, BufStart, SampleRun
  J2 = NotesDLT.Stack(J2).ZLink.Handoff
 Next I2
 
 ConvertSignal BufStart, SampleRun
 
 WriteBuf DSB, lpBuffer, BufStart, SampleRun, DSBLOCK_DEFAULT
 
End Sub
Private Sub WorkEnvelopPoints(Note As FM_Note, BufStart&, SampleRun&)
Dim SamplesCompleted&
Dim I2&

 Note.BufStart = BufStart

 Do While SamplesCompleted < SampleRun
  
  If Note.BufStart > BufUBound Then
   I = Note.BufStart - BufUBound
   Note.BufStart = LBound(sRender) + I
  End If
  
  Note.AddToStart = SampleRunM1 - SamplesCompleted
  
  Note.BufStop = Note.BufStart + Note.AddToStart
  If Note.BufStop > BufUBound Then Note.AddToStart = BufUBound - Note.BufStart
  
  Select Case Note.EnvelopeMode
  Case EnvelopeMode_ATTACK
   If Note.cSamplesToEnvelopeTarget < Note.AddToStart Then Note.AddToStart = Note.cSamplesToEnvelopeTarget
   AddTo_HiFi_Signal Note
   I = Note.AddToStart + 1
   If Note.cSamplesToEnvelopeTarget < 1 Then
    Note.EnvelopeMode = EnvelopeMode_SUSTAIN
   End If
   Note.cSamplesToEnvelopeTarget = Note.cSamplesToEnvelopeTarget - I
   SamplesCompleted = SamplesCompleted + I
   
  Case EnvelopeMode_SUSTAIN
   AddTo_HiFi_Signal Note
   SamplesCompleted = SamplesCompleted + Note.AddToStart + 1
  
  Case EnvelopeMode_RELEASE
   If Note.cSamplesToEnvelopeTarget < Note.AddToStart Then Note.AddToStart = Note.cSamplesToEnvelopeTarget
   AddTo_HiFi_Signal Note
   SamplesCompleted = SamplesCompleted + Note.AddToStart + 1
   If Note.cSamplesToEnvelopeTarget < 1 Then
    RemoveLink Notes, (Note.ElemRef)
    SamplesCompleted = SampleRun
   End If
   I = Note.AddToStart + 1
   Note.cSamplesToEnvelopeTarget = Note.cSamplesToEnvelopeTarget - I
  
  End Select
  
 Loop
 
End Sub

Private Sub AddTo_HiFi_Signal(Note As FM_Note)

 Note.BufStop = Note.BufStart + Note.AddToStart
 
 If Note.EnvelopeMode = EnvelopeMode_SUSTAIN Then
 
  For I = Note.BufStart To Note.BufStop
  
   sRender(I) = sRender(I) + _
    Sin(Note.BaseFreq.sPos + Knob_FM_Ampl.vVal * Sin(Note.BaseFM.sPos)) + _
    Sin(Note.Freq2.sPos + Knob_FM_Ampl.vVal * Sin(Note.FM2.sPos))
    
   Note.BaseFreq.sPos = Note.BaseFreq.sPos + Note.BaseFreq.iPos
   Note.BaseFM.sPos = Note.BaseFM.sPos + Note.BaseFM.iPos
   Note.Freq2.sPos = Note.Freq2.sPos + Note.Freq2.iPos
   Note.FM2.sPos = Note.FM2.sPos + Note.FM2.iPos
   
  Next
 
 Else
 
  For I = Note.BufStart To Note.BufStop
  
   sRender(I) = sRender(I) + _
    Note.sAmpEnv * ((Sin(Note.BaseFreq.sPos + Knob_FM_Ampl.vVal * Sin(Note.BaseFM.sPos)) + _
    Sin(Note.Freq2.sPos + Knob_FM_Ampl.vVal * Sin(Note.FM2.sPos))))
    
   Note.BaseFreq.sPos = Note.BaseFreq.sPos + Note.BaseFreq.iPos
   Note.BaseFM.sPos = Note.BaseFM.sPos + Note.BaseFM.iPos
   Note.Freq2.sPos = Note.Freq2.sPos + Note.Freq2.iPos
   Note.FM2.sPos = Note.FM2.sPos + Note.FM2.iPos
   Note.sAmpEnv = Note.sAmpEnv + Note.iAmpEnv
   
  Next
  
 End If
 
 Note.BufStart = Note.BufStop + 1
 
End Sub
Private Sub ConvertSignal(Optional ByVal BufStart& = 0&, Optional ByVal Len1& = BufferSamples)
Dim levL!
Dim bufLvl!

 TmpBufStop = BufStart + Len1 - 1
 
 If TmpBufStop > BufUBound Then
  For I = BufStart& To BufUBound
   bufLvl = (sRender(I) * KnobVolume.vVal) * CONST_HALF_VERTICAL_RESOLUTION
   If bufLvl > CONST_MAX_VERTICAL Then
    bufLvl = CONST_MAX_VERTICAL
   ElseIf bufLvl < CONST_MIN_VERTICAL Then
    bufLvl = CONST_MIN_VERTICAL
   End If
   lpBuffer(I) = bufLvl
  Next
  BufStart = LBound(lpBuffer)
  TmpBufStop = TmpBufStop - BufferSamples
 End If

 For I = BufStart& To TmpBufStop
  bufLvl = (sRender(I) * KnobVolume.vVal) * CONST_HALF_VERTICAL_RESOLUTION
  If bufLvl > CONST_MAX_VERTICAL Then
   bufLvl = CONST_MAX_VERTICAL
  ElseIf bufLvl < CONST_MIN_VERTICAL Then
   bufLvl = CONST_MIN_VERTICAL
  End If
  lpBuffer(I) = bufLvl
 Next

End Sub
Private Sub WriteBuf(DSB1 As DirectSoundBuffer, BufSrc16() As Integer, ByVal BufStart&, ByVal Length&, Optional LFlags As CONST_DSBLOCKFLAGS = DSBLOCK_DEFAULT)

 TmpBufStop = BufStart + Length - 1
 
 If TmpBufStop > BufUBound Then
  Length = BufUBound - BufStart + 1
  DSB1.WriteBuffer BufStart * PCM.nBlockAlign, Length * PCM.nBlockAlign, BufSrc16(0&), LFlags
  BufStart = LBound(lpBuffer)
  Length = TmpBufStop - BufUBound
 End If
 
 DSB1.WriteBuffer BufStart * PCM.nBlockAlign, Length * PCM.nBlockAlign, BufSrc16(0&), LFlags
 
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Clicked_Which_Knob_Ary Me, Button, Shift, x, y
 If vSelect > 0 Then Call Form_MouseMove(Button, Shift, x, y)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 
 If Button = 0 Then Clicked_Which_Knob_Ary Me, Button, Shift, x, y
 
 modControls.ShiftKeyDown = Shift
 
 Select Case vSelect
 Case Knob_FM_Freq_Ratio.Index
  CalcKnobVal Me, Knob_FM_Freq_Ratio, y
  Caption = "FM Freq / Note Freq   " & Knob_FM_Freq_Ratio.vVal
  SetFreqForAllNotesPlaying Notes
 Case Knob_FM_Ampl.Index
  CalcKnobVal Me, Knob_FM_Ampl, y
  Caption = "FM Amount " & Knob_FM_Ampl.vVal
 Case KnobPitchLayer2.Index
  CalcKnobVal Me, KnobPitchLayer2, y
  Caption = "second layer pitch " & KnobPitchLayer2.vVal
  SetFreqForAllNotesPlaying Notes
 Case KnobVolume.Index
  CalcKnobVal Me, KnobVolume, y
  Caption = "Volume " & KnobVolume.vVal * 100
 Case Knob_GL_Pitch.Index
  CalcKnobVal Me, Knob_GL_Pitch, y
  sFreq = reference_iFreq * 2 ^ (Knob_GL_Pitch.vVal / 12) / SampleRate
  SetFreqForAllNotesPlaying Notes
  If Knob_GL_Pitch.vVal <> 0 Then
   If Abs(Knob_GL_Pitch.vVal) = 1 Then
    TmpStr = " semitone"
   Else
    TmpStr = " semitones"
   End If
   If Knob_GL_Pitch.vVal > 0 Then
    Caption = "Global Pitch  +" & Knob_GL_Pitch.vVal & TmpStr
   Else
    Caption = "Global Pitch  " & Knob_GL_Pitch.vVal & TmpStr
   End If
  Else
   Caption = "Global Pitch  Normal"
  End If
 Case Knob_GL_Attack.Index
  CalcKnobVal Me, Knob_GL_Attack, y
  iAttackRate = 1 / (Knob_GL_Attack.vVal + 1)
  Caption = "Attack: " & Round((Knob_GL_Attack.vVal / SampleRate), 2) & " seconds"
 Case Knob_GL_Release.Index
  CalcKnobVal Me, Knob_GL_Release, y
  iReleaseRate = -1 / (Knob_GL_Release.vVal + 1)
  Caption = "Release: " & Round((Knob_GL_Release.vVal / SampleRate), 2) & " seconds"
 End Select
 
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 ResetKnob
End Sub

Private Sub Form_Paint()
 Draw_Knobs Me
End Sub

Private Sub Form_Resize()
 DoEvents
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 ClearControlSurfaces
 ClearSurface InfoBoard
End Sub



Private Sub InitializeNoteInLink(NotesDLT As DataLinkType, KeyReference As Integer, ElemRef As Integer) ', Optional bResetPos As Long)
 InitNoteVariables NotesDLT, NotesDLT.Stack(NotesDLT.Elem_Tail).Whatever, KeyReference
End Sub
Private Sub InitNoteVariables(NotesDLT As DataLinkType, Note As FM_Note, KeyReference As Integer) ', Optional ByVal bResetPos As Integer)

 Note.ElemRef = NotesDLT.Elem_Tail
 NotesDLT.GetStackElemFromKey(KeyReference) = Note.ElemRef
 
 Note.KeyRef = KeyReference
 
 SetNoteFrequencyVariables Note, sFreq * 2 ^ ((GetNoteIndexFromKey(KeyReference) - 60 + Note.HalfToneOffset) / 12)
 ResetPos Note
 
 Note.cSamplesToEnvelopeTarget = Knob_GL_Attack.vVal - 1
 Note.iAmpEnv = iAttackRate
 
 Note.EnvelopeMode = EnvelopeMode_ATTACK
 
End Sub
Private Sub SetNoteFrequencyVariables(Note As FM_Note, iFreq As Double)
 Note.BaseFreq.iPos = iFreq
 Note.Freq2.iPos = iFreq * KnobPitchLayer2.vVal
 Note.BaseFM.iPos = iFreq * Knob_FM_Freq_Ratio.vVal
 Note.FM2.iPos = Note.Freq2.iPos * Knob_FM_Freq_Ratio.vVal
End Sub
Private Sub SetFreqForAllNotesPlaying(NotesDLT As DataLinkType)
 I = NotesDLT.Elem_Head
 For J = NotesDLT.LB To NotesDLT.InUsePtr
  With NotesDLT.Stack(I)
   SetNoteFrequencyVariables .Whatever, sFreq * 2 ^ ((GetNoteIndexFromKey(.Whatever.KeyRef) - 60) / 12)
   I = .ZLink.Handoff
  End With
 Next J
End Sub
Private Sub ResetPos(Note As FM_Note)
 Note.BaseFM.sPos = 0
 Note.BaseFreq.sPos = 0
 Note.Freq2.sPos = 0
 Note.FM2.sPos = 0
 Note.sAmpEnv = iAttackRate
End Sub
Private Sub TargetNoteRelease(Note As FM_Note)
 Note.cSamplesToEnvelopeTarget = SampleRunM1
 Note.EnvelopeMode = EnvelopeMode_RELEASE
 Note.iAmpEnv = iReleaseRate
 Note.sAmpEnv = Note.sAmpEnv + Note.iAmpEnv
 Note.cSamplesToEnvelopeTarget = Int(Knob_GL_Release.vVal * Note.sAmpEnv) - 1
End Sub

