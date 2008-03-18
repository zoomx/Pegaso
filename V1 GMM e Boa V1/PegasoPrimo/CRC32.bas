Attribute VB_Name = "CRC32"
Option Explicit

Public Crc32Table(256) As Long
Public Crc32POLY As Long
Public FACTOR As Long
Public Init32Flag As Integer

' calculates the 32 bit Crc using the byte method

Private Function Crc32Byte(ByVal CRC As Long, ByVal Octet As Byte) As Long

Dim J As Integer

 ' XOR data byte with high byte of accumulator
 CRC = (CRC Xor Octet) And &HFF
 For J = 0 To 7
    If (CRC And 1) Then
       SHIFT RIGHT Crc, 1
       CRC = CRC Xor Crc32POLY
    Else
       SHIFT RIGHT Crc,1
    End If
 Next J
 Crc32Byte = CRC

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

  Print "CRC-32 Table Built"

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
                     ByVal CRC As Long) As Long
  Dim RightShifted As Long
  Dim Index As Integer

  If Init32Flag = 0 Then
    Call Init32Crc
  End If

  'compute CRC
  RightShifted = CRC
  SHIFT RIGHT RightShifted,8
  Index = &HFF And (CRC Xor Octet)
  UpdateCrc32 = Crc32Table(Index) Xor (RightShifted And FACTOR)
  'Crc32Table[(unsigned char)(Crc ^ (long)(Octet))] ^ ((Crc >> 8) & 0x00FFFFFF);
End Function

Public Function UpdateCRC(ByVal Octet As Byte, _
                   ByVal CRC As Long)
  Dim LeftShifted  As WORD
  Dim RightShifted As WORD

  If Initialized = 0 Then
    InitCRC
  End If

  'compute CRC
  LeftShifted = CRC
  SHIFT LEFT LeftShifted,8
  RightShifted = CRC
  SHIFT RIGHT RightShifted,8
  UpdateCRC = LeftShifted Xor (CRCtable(RightShifted Xor Octet))
End Function

