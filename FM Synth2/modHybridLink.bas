Attribute VB_Name = "modHybridLink"
Option Explicit

Private Type LinkAPI
 Recv    As Integer
 Handoff As Integer
 Avail   As Integer
 Reserve As Integer
End Type

Private Type DataLink
 Whatever As FM_Note
 ZLink    As LinkAPI
End Type

Public Type DataLinkType
 Stack()       As DataLink 'link of bytes
 InUsePtr      As Long         '
 Elem_Head     As Integer 'corresp. start element
 Elem_Tail     As Integer 'corresp. last element
 Elem          As Integer
 LB            As Integer
 UB            As Integer
 Reserve       As Integer
 GetStackElemFromKey(255) As Integer
 ElementCount  As Long
End Type

Private Elem&, I&, J&
Private TmpHand%, TmpRecv%

Public Const LB1 As Long = -32768
Public Const UB1 As Long = 32767

Public Sub InitializeStack(TheData As DataLinkType, ByVal ElemCount_65536_Max&)
 If ElemCount_65536_Max > 65536 Then ElemCount_65536_Max = 65536
 TheData.Elem_Head = LB1
 TheData.Elem_Tail = LB1
 TheData.InUsePtr = LB1 - 1
 TheData.ElementCount = ElemCount_65536_Max
 TheData.UB = TheData.InUsePtr + TheData.ElementCount
 ReDim TheData.Stack(LB1 To TheData.UB)
 Elem = TheData.UB
 For I = LB1 To Elem
  TheData.Stack(I).ZLink.Avail = I
 Next I
 TheData.LB = LB1
End Sub
Public Function AddLink(TheData As DataLinkType) As Boolean
Dim Bln1 As Boolean
 If TheData.InUsePtr < TheData.UB Then
  TheData.InUsePtr = TheData.InUsePtr + 1
  Elem = TheData.Stack(TheData.InUsePtr).ZLink.Avail
  If TheData.InUsePtr = LB1 Then
   TheData.Elem_Head = Elem
   TheData.Elem_Tail = Elem
  Else
   TheData.Stack(TheData.Elem_Tail).ZLink.Handoff = Elem
   TheData.Stack(Elem).ZLink.Recv = TheData.Elem_Tail
   TheData.Elem_Tail = Elem
  End If
  Bln1 = True
 End If
 AddLink = Bln1
End Function
Public Function RemoveLink(TheData As DataLinkType, Element As Integer) As Boolean
Dim Bln1 As Boolean
 
 If TheData.InUsePtr >= TheData.LB Then
 
  TheData.Stack(TheData.InUsePtr).ZLink.Avail = Element
 
  If Element <> TheData.Elem_Head Then
   TmpRecv = TheData.Stack(Element).ZLink.Recv
   If Element <> TheData.Elem_Tail Then
    TmpHand = TheData.Stack(Element).ZLink.Handoff
    TheData.Stack(TmpHand).ZLink.Recv = TmpRecv
    TheData.Stack(TmpRecv).ZLink.Handoff = TmpHand
   Else
    TheData.Elem_Tail = TmpRecv
   End If
  Else
   If Element <> TheData.Elem_Tail Then
    TmpHand = TheData.Stack(Element).ZLink.Handoff
    TheData.Elem_Head = TmpHand
   End If
  End If
 
  Bln1 = True
 
  TheData.InUsePtr = TheData.InUsePtr - 1
 
  RemoveLink = Bln1
 
 End If
 
End Function
Public Sub ClearStack(TheData As DataLinkType)
 Elem = TheData.Elem_Head
 J = TheData.InUsePtr
 For I = TheData.LB To J
  RemoveLink TheData, TheData.Elem_Head
  Elem = TheData.Stack(Elem).ZLink.Handoff
 Next I
End Sub
