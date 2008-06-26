Attribute VB_Name = "BitwiseOperations"
Option Explicit
'Author: Nathan Moschkin (Featured Developer)
'http://www.freevbcode.com/ShowCode.asp?ID=2045
Private OnBits(0 To 31) As Long

Public Function LShiftLong(ByVal Value As Long, _
    ByVal Shift As Integer) As Long

    MakeOnBits

    If (Value And (2 ^ (31 - Shift))) Then GoTo OverFlow

    LShiftLong = ((Value And OnBits(31 - Shift)) * (2 ^ Shift))

    Exit Function

OverFlow:
    Debug.Print "LShiftLong OverFlow"
    LShiftLong = ((Value And OnBits(31 - (Shift + 1))) * _
       (2 ^ (Shift))) Or &H80000000

End Function

Public Function RShiftLong(ByVal Value As Long, _
   ByVal Shift As Integer) As Long
    Dim hi As Long
    MakeOnBits
    If (Value And &H80000000) Then hi = &H40000000

    RShiftLong = (Value And &H7FFFFFFE) \ (2 ^ Shift)
    RShiftLong = (RShiftLong Or (hi \ (2 ^ (Shift - 1))))
End Function
 
Public Function LShift2(ByVal Value As Integer, _
    ByVal Shift As Integer) As Integer

    MakeOnBits

    If (Value And (2 ^ (7 - Shift))) Then GoTo OverFlow

    LShift2 = ((Value And OnBits(7 - Shift)) * (2 ^ Shift))

    Exit Function

OverFlow:

    LShift2 = ((Value And OnBits(7 - (Shift + 1))) * _
       (2 ^ (Shift))) Or &H8

End Function

Public Function RShift2(ByVal Value As Integer, _
   ByVal Shift As Integer) As Integer
    Dim hi As Integer
    MakeOnBits
    If (Value And &H8) Then hi = &H4

    RShift2 = (Value And &H7FFFFFFE) \ (2 ^ Shift)
    RShift2 = (RShift2 Or (hi \ (2 ^ (Shift - 1))))
End Function


Private Sub MakeOnBits()
    Dim J As Integer, _
        v As Long
  
    For J = 0 To 30
  
        v = v + (2 ^ J)
        OnBits(J) = v
  
    Next J
  
    OnBits(J) = v + &H80000000

End Sub

 ' Thank you Lewis Moten.  Why doen't VB support this?
 Private Function LShift3(ByVal pnValue As Double, ByVal pnShift As Double) As Double
     ' Equivilant to C's Bitwise << operator
'     LShift = pnValue * (2 ^ pnShift)
 End Function

 Private Function RShift3(ByVal pnValue As Double, ByVal pnShift As Double) As Double
     ' Equivilant to C's Bitwise >> operator
 '    RShift = pnValue \ (2 ^ pnShift)
 End Function

'Public Function LShift(w As Long, c As Integer) As Integer
'    Dim lngShifted As Long
'    lngShifted = w * (2 ^ c)
'    If lngShifted < -32768 Then lngShifted = lngShifted + 65536
'    LShift = Val("&h" & Hex$((&HFFFF And lngShifted)))
'End Function

' Public Function RShift(ByVal lThis As Long, ByVal lBits As Long) As Long
'   If (lBits <= 0) Then
'      RShift = lThis
'   ElseIf (lBits > 63) Then
'      RShift = 0
'      Exit Function ' ... error ...
'   ElseIf (lBits > 31) Then
'      RShift = 0
'   Else
'      If (lThis And m_lPower2(31)) = m_lPower2(31) Then
'         RShift = (lThis And &H7FFFFFFF) \ m_lPower2(lBits) Or m_lPower2(31 - lBits)
'      Else
'         RShift = lThis \ m_lPower2(lBits)
'      End If
'   End If
'   RShift = GetLoWord(RShift)
'End Function



'*----------------------------------------------------------*
'* Name       : vbShiftLeft                                 *
'*----------------------------------------------------------*
'* Purpose    : Shift 32-bit integer value left 'n' bits.   *
'*----------------------------------------------------------*
'* Parameters : Value  Required. Value to shift.            *
'*            : Count  Required. Number of bit positions to *
'*            :        shift value.                         *
'*----------------------------------------------------------*
'* Description: This function is equivalent to the 'C'      *
'*            : language construct '<<'.                    *
'*----------------------------------------------------------*
'http://vbcity.com/page.asp?f=howto&p=bit_shift

Public Function LShift(ByVal Value As Long, _
                            Count As Integer) As Long
Dim i As Integer
Dim tLshift As Double
  LShift = Value
    tLshift = LShift

  For i = 1 To Count
    tLshift = tLshift * 2
  Next
  LShift = CLng(tLshift)
End Function

'*----------------------------------------------------------*
'* Name       : vbShiftRight                                *
'*----------------------------------------------------------*
'* Purpose    : Shift 32-bit integer value right 'n' bits.  *
'*----------------------------------------------------------*
'* Parameters : Value  Required. Value to shift.            *
'*            : Count  Required. Number of bit positions to *
'*            :        shift value.                         *
'*----------------------------------------------------------*
'* Description: This function is equivalent to the 'C'      *
'*            : language construct '>>'.                    *
'*----------------------------------------------------------*
'http://vbcity.com/page.asp?f=howto&p=bit_shift

Public Function RShift(ByVal Value As Long, _
                             Count As Integer) As Long
Dim i As Integer

  RShift = Value

  For i = 1 To Count
    RShift = RShift \ 2
  Next

End Function

