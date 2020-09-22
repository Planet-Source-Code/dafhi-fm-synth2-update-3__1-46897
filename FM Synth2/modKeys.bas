Attribute VB_Name = "modKeys"
Option Explicit

Public GetNoteIndexFromKey(255) As Long
Public KeyPressed(255) As Long

Public Const NotesLB As Long = 0
Public Const NotesUB As Long = NotesLB + 108

Public Sub InitKeys(DoFullRange As Boolean)
 
 AsNoV vbKeyZ, 4, 0
 AsNoV vbKeyS, 4, 1
 AsNoV vbKeyX, 4, 2
 AsNoV vbKeyD, 4, 3
 AsNoV vbKeyC, 4, 4
 AsNoV vbKeyV, 4, 5
 AsNoV vbKeyG, 4, 6
 AsNoV vbKeyB, 4, 7
 AsNoV vbKeyH, 4, 8
 AsNoV vbKeyN, 4, 9
 AsNoV vbKeyJ, 4, 10
 AsNoV vbKeyM, 4, 11
 
 AsNoV 188, 5, 0    ' <
 AsNoV vbKeyL, 5, 1 ' L
 AsNoV 190, 5, 2    ' >
 AsNoV 186, 5, 3    'r semicolon
 AsNoV 191, 5, 4    ' ?
 
 If DoFullRange Then

 AsNoV vbKeyQ, 5, 5
 AsNoV vbKey2, 5, 6
 AsNoV vbKeyW, 5, 7
 AsNoV vbKey3, 5, 8
 AsNoV vbKeyE, 5, 9
 AsNoV vbKey4, 5, 10
 AsNoV vbKeyR, 5, 11
 
 AsNoV vbKeyT, 6, 0
 AsNoV vbKey6, 6, 1
 AsNoV vbKeyY, 6, 2
 AsNoV vbKey7, 6, 3
 AsNoV vbKeyU, 6, 4
 
 AsNoV vbKeyI, 6, 5
 AsNoV vbKey9, 6, 6
 AsNoV vbKeyO, 6, 7
 AsNoV vbKey0, 6, 8
 AsNoV vbKeyP, 6, 9
 AsNoV 189, 6, 10 ' -
 AsNoV 219, 6, 11 ' [
 AsNoV 221, 7, 0 ' [
 
 Else

 AsNoV vbKeyQ, 5, 0
 AsNoV vbKey2, 5, 1
 AsNoV vbKeyW, 5, 2
 AsNoV vbKey3, 5, 3
 AsNoV vbKeyE, 5, 4
 AsNoV vbKeyR, 5, 5
 AsNoV vbKey5, 5, 6
 AsNoV vbKeyT, 5, 7
 AsNoV vbKey6, 5, 8
 AsNoV vbKeyY, 5, 9
 AsNoV vbKey7, 5, 10
 AsNoV vbKeyU, 5, 11
 
 AsNoV vbKeyI, 6, 0
 AsNoV vbKey9, 6, 1
 AsNoV vbKeyO, 6, 2
 AsNoV vbKey0, 6, 3
 AsNoV vbKeyP, 6, 4
 AsNoV 219, 6, 5 ' [
 AsNoV 187, 6, 6 ' [
 AsNoV 221, 6, 7 ' + =
 
 End If
 
End Sub

Private Sub AsNoV(KeyIndex As Byte, Octave As Byte, KeyOffset As Byte)
 GetNoteIndexFromKey(KeyIndex) = Octave * 12 + KeyOffset + NotesLB
End Sub
