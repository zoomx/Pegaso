Attribute VB_Name = "CRC32"
Option Explicit

Public Crc32Table(256) As Long
Public Crc32POLY As Long
Public FACTOR As Long
Public Init32Flag As Integer
Public Initialized As Byte
Public CrcTable(256) As Long
' calculates the 32 bit Crc using the byte method

Private Function Crc32Byte(ByVal Crc As Long, ByVal Octet As Byte) As Long

Dim J As Integer

 ' XOR data byte with high byte of accumulator
 Crc = (Crc Xor Octet) And &HFF   'Controllare e riattivare
 For J = 0 To 7
    If (Crc And 1) Then
       'SHIFT RIGHT Crc, 1
       Crc = RShiftLong(Crc, 1)
       Crc = Crc Xor Crc32POLY
    Else
       'SHIFT RIGHT Crc,1
       Crc = RShiftLong(Crc, 1)
    End If
 Next J
 Crc32Byte = Crc

End Function

' initialize CRC table

Private Sub Init32Crc()

  Dim i As Integer

  Init32Flag = 1                        'only need to do this once
  Crc32POLY = &HEDB88320
  FACTOR = &HFFFFFF

  For i = 0 To 255
    Crc32Table(i) = Crc32Byte(0, i)
  Next i

  'Print "CRC-32 Table Built"

End Sub

' Fetch CRC table entry

Public Function FetchCrc32(ByVal Index As Byte) As Long
If Init32Flag = 0 Then
    Call Init32Crc
End If
FetchCrc32 = Crc32Table(Index)
End Function

' compute updated CRC

Public Function UpdateCrc32(ByVal Octet As Byte, _
                     ByVal Crc As Long) As Long
  Dim RightShifted As Long
  Dim Index As Integer

  If Init32Flag = 0 Then
    Call Init32Crc
  End If

  'compute CRC
  RightShifted = Crc
  'SHIFT RIGHT RightShifted,8   'Riattivare
  RightShifted = RShiftLong(RightShifted, 8)
  'RightShifted = RShift(RightShifted, 8)
  Index = &HFF And (Crc Xor Octet)
  UpdateCrc32 = Crc32Table(Index) Xor (RightShifted And FACTOR)
  'Crc32Table[(unsigned char)(Crc ^ (long)(Octet))] ^ ((Crc >> 8) & 0x00FFFFFF);
End Function

Public Function UpdateCRC(ByVal Octet As Byte, _
                   ByVal Crc As Long) As Long
  Dim LeftShifted  As Long
  Dim RightShifted As Long

  If Initialized = 0 Then
    
    InitCrc
  End If

  'compute CRC
  LeftShifted = Crc
  'SHIFT LEFT LeftShifted,8
  LeftShifted = LShiftLong(LeftShifted, 8)
  'LeftShifted = LShift2(LeftShifted, 8)
  RightShifted = Crc
  Debug.Print "CRC="; Crc
  'SHIFT RIGHT RightShifted,8
  'RightShifted = RShiftLong(RightShifted, 8)
  RightShifted = RShift(RightShifted, 8)
  Debug.Print "RightShifted="; RightShifted
  UpdateCRC = LeftShifted Xor (CrcTable(RightShifted Xor Octet))
End Function

Private Sub InitCrc()

  Dim i As Integer

  Initialized = 1              'only need to do this once
  Debug.Print "InitCRC"
  
  For i = 0 To 255
    CrcTable(i) = CalcTable(i, &H1021, 0)
    'Debug.Print "CalcTable "; i; " "; CrcTable(i)
  Next i

  Debug.Print "CRC-16 Table Built"

End Sub

Private Function CalcTable(ByVal Octet As Long, ByVal GenPoly As Long, ByVal Accum As Long) As Long
  Dim i As Integer
  Dim J As Integer

  'SHIFT LEFT Octet,8
  Octet = LShiftLong(Octet, 8)
  'Octet = LShift2(Octet, 8)

  For J = 1 To 8
    i = 9 - J
    If ((Octet Xor Accum) And &H8000) Then
      'SHIFT LEFT Accum,1
      Accum = LShiftLong(Accum, 8)
      'Accum = LShift2(Accum, 8)
      Accum = Accum Xor GenPoly
    Else
      'SHIFT LEFT Accum,1
      Accum = LShiftLong(Accum, 8)
      'Accum = LShift2(Accum, 8)

    End If
    'SHIFT LEFT Octet,1
    Octet = LShiftLong(Octet, 1)
    'Octet = LShift2(Octet, 1)
  Next J
  CalcTable = Accum

End Function

Public Function UpdateCrc16(ByVal Octet As Byte, _
                     ByVal Crc As Long)
  Dim LeftShifted  As Long
  Dim RightShifted As Long

  If Initialized = 0 Then
    Call InitCrc
  End If

  'compute CRC-16
  LeftShifted = Crc
  'SHIFT LEFT LeftShifted,8
  LeftShifted = LShiftLong(LeftShifted, 8)
  RightShifted = Crc
  'SHIFT RIGHT RightShifted,8
  RightShifted = RShiftLong(RightShifted, 8)
  UpdateCrc16 = CrcTable(RightShifted And 255) Xor LeftShifted Xor Octet

End Function

