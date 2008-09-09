VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mcLCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Coordinate
   X As Integer
   Y As Integer
End Type

Dim BasePoint As Coordinate

Dim SegWidth As Integer
Dim SegHeight As Integer

Dim P As PictureBox
Property Let BackColor(Color As Long)
    'Imposta il colore di sfondo.
    P.BackColor = Color
End Property

Private Sub DrawChar(Char As String)
    Select Case Char
        Case "0"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
            DrawSegment (6)
        Case "1"
            DrawSegment (2)
            DrawSegment (3)
        Case "2"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (5)
            DrawSegment (41)
            DrawSegment (42)
        Case "3"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
        Case "4"
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (6)
        Case "5"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
        Case "6"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
        Case "7"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
        Case "8"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
        Case "9"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
        Case "A"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
        Case "B"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (72)
            DrawSegment (81)
            DrawSegment (82)
        Case "C"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (41)
            DrawSegment (42)
        Case "D"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (81)
            DrawSegment (82)
        Case "E"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (71)
        Case "F"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (71)
        Case "G"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (72)
        Case "H"
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
        Case "I"
            DrawSegment (81)
            DrawSegment (82)
        Case "J"
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
        Case "K"
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (92)
            DrawSegment (94)
        Case "L"
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
            DrawSegment (6)
        Case "M"
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (91)
            DrawSegment (92)
        Case "N"
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (91)
            DrawSegment (94)
        Case "O"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
            DrawSegment (6)
        Case "P"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
        Case "Q"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (94)
        Case "R"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (2)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (94)
        Case "S"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
        Case "T"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (81)
            DrawSegment (82)
        Case "U"
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
            DrawSegment (6)
        Case "V"
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (92)
            DrawSegment (93)
        Case "W"
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (93)
            DrawSegment (94)
        Case "X"
            DrawSegment (91)
            DrawSegment (92)
            DrawSegment (93)
            DrawSegment (94)
        Case "Y"
            DrawSegment (82)
            DrawSegment (91)
            DrawSegment (92)
        Case "Z"
            DrawSegment (11)
            DrawSegment (12)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (92)
            DrawSegment (93)
        Case "b"
            DrawSegment (3)
            DrawSegment (5)
            DrawSegment (6)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (71)
            DrawSegment (72)
        Case "c"
            DrawSegment (5)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (71)
            DrawSegment (72)
        Case "d"
            DrawSegment (2)
            DrawSegment (3)
            DrawSegment (5)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (71)
            DrawSegment (72)
        Case "h"
            DrawSegment (5)
            DrawSegment (3)
            DrawSegment (6)
            DrawSegment (71)
            DrawSegment (72)
        Case "i"
            DrawSegment (82)
        Case "l"
            DrawSegment (81)
            DrawSegment (82)
        Case "n"
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (3)
            DrawSegment (5)
        Case "o"
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (3)
            DrawSegment (5)
        Case "r"
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (5)
        Case "t"
            DrawSegment (71)
            DrawSegment (6)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (5)
        Case "u"
            DrawSegment (5)
            DrawSegment (41)
            DrawSegment (42)
            DrawSegment (3)
        Case "/"
            DrawSegment (92)
            DrawSegment (93)
        Case "\"
            DrawSegment (91)
            DrawSegment (94)
        Case "-"
            DrawSegment (71)
            DrawSegment (72)
        Case "+"
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (81)
            DrawSegment (82)
        Case "*"
            DrawSegment (71)
            DrawSegment (72)
            DrawSegment (81)
            DrawSegment (82)
            DrawSegment (91)
            DrawSegment (92)
            DrawSegment (93)
            DrawSegment (94)
   End Select
End Sub
Private Sub DrawSegment(SegNum As Integer)
    '
    '      1
    '     ___
    '    |   |
    ' 6  |   |  2
    '    |-7-|
    ' 5  |   |  3
    '    |___|
    '
    '      4
    '
    With BasePoint
    Select Case SegNum
        Case 11
            P.Line (.X + 1, .Y + 0)-(.X + ((SegWidth) \ 2) - 1, .Y + 0)
            P.Line (.X + 2, .Y + 1)-(.X + ((SegWidth) \ 2) - 0, .Y + 1)
            P.Line (.X + 3, .Y + 2)-(.X + ((SegWidth) \ 2) - 1, .Y + 2)
       Case 12
            P.Line (.X + ((SegWidth) \ 2) + 1, .Y + 0)-(.X + SegWidth - 1, .Y + 0)
            P.Line (.X + ((SegWidth) \ 2) + 0, .Y + 1)-(.X + SegWidth - 2, .Y + 1)
            P.Line (.X + ((SegWidth) \ 2) + 1, .Y + 2)-(.X + SegWidth - 3, .Y + 2)
       Case 2
            P.Line (.X + SegWidth - 1, .Y + 1)-(.X + SegWidth - 1, .Y + (SegHeight \ 2) - 1)
            P.Line (.X + SegWidth - 2, .Y + 2)-(.X + SegWidth - 2, .Y + (SegHeight \ 2))
            P.Line (.X + SegWidth - 3, .Y + 3)-(.X + SegWidth - 3, .Y + (SegHeight \ 2) - 1)
       Case 3
            P.Line (.X + SegWidth - 1, .Y + (SegHeight \ 2) + 2)-(.X + SegWidth - 1, .Y + SegHeight)
            P.Line (.X + SegWidth - 2, .Y + (SegHeight \ 2) + 1)-(.X + SegWidth - 2, .Y + SegHeight - 1)
            P.Line (.X + SegWidth - 3, .Y + (SegHeight \ 2) + 2)-(.X + SegWidth - 3, .Y + SegHeight - 2)
       Case 41
            P.Line (.X + 3, .Y + SegHeight - 2)-(.X + ((SegWidth) \ 2) - 1, .Y + SegHeight - 2)
            P.Line (.X + 2, .Y + SegHeight - 1)-(.X + ((SegWidth) \ 2) - 0, .Y + SegHeight - 1)
            P.Line (.X + 1, .Y + SegHeight - 0)-(.X + ((SegWidth) \ 2) - 1, .Y + SegHeight - 0)
       Case 42
            P.Line (.X + ((SegWidth) \ 2) + 1, .Y + SegHeight - 2)-(.X + SegWidth - 3, .Y + SegHeight - 2)
            P.Line (.X + ((SegWidth) \ 2) + 0, .Y + SegHeight - 1)-(.X + SegWidth - 2, .Y + SegHeight - 1)
            P.Line (.X + ((SegWidth) \ 2) + 1, .Y + SegHeight - 0)-(.X + SegWidth - 1, .Y + SegHeight - 0)
       Case 5
            P.Line (.X, .Y + (SegHeight \ 2) + 2)-(.X, .Y + SegHeight)
            P.Line (.X + 1, .Y + (SegHeight \ 2) + 1)-(.X + 1, .Y + SegHeight - 1)
            P.Line (.X + 2, .Y + (SegHeight \ 2) + 2)-(.X + 2, .Y + SegHeight - 2)
       Case 6
            P.Line (.X, .Y + 1)-(.X, .Y + (SegHeight \ 2) - 1)
            P.Line (.X + 1, .Y + 2)-(.X + 1, .Y + (SegHeight \ 2))
            P.Line (.X + 2, .Y + 3)-(.X + 2, .Y + (SegHeight \ 2) - 1)
       Case 71
            P.Line (.X + 3, .Y + (SegHeight \ 2) - 1)-(.X + ((SegWidth) \ 2) - 1, .Y + (SegHeight \ 2) - 1)
            P.Line (.X + 2, .Y + (SegHeight \ 2) - 0)-(.X + ((SegWidth) \ 2) - 0, .Y + (SegHeight \ 2) - 0)
            P.Line (.X + 3, .Y + (SegHeight \ 2) + 1)-(.X + ((SegWidth) \ 2) - 1, .Y + (SegHeight \ 2) + 1)
       Case 72
            P.Line (.X + ((SegWidth) \ 2) + 1, .Y + (SegHeight \ 2) - 1)-(.X + SegWidth - 3, .Y + (SegHeight \ 2) - 1)
            P.Line (.X + ((SegWidth) \ 2) - 0, .Y + (SegHeight \ 2) - 0)-(.X + SegWidth - 2, .Y + (SegHeight \ 2) - 0)
            P.Line (.X + ((SegWidth) \ 2) + 1, .Y + (SegHeight \ 2) + 1)-(.X + SegWidth - 3, .Y + (SegHeight \ 2) + 1)
       Case 81
            P.Line ((.X + (SegWidth \ 2)) - 1, .Y + 3)-(.X + (SegWidth \ 2) - 1, .Y + (SegHeight \ 2) - 1)
            P.Line ((.X + (SegWidth \ 2)) + 0, .Y + 2)-(.X + (SegWidth \ 2) + 0, .Y + (SegHeight \ 2) - 0)
            P.Line ((.X + (SegWidth \ 2)) + 1, .Y + 3)-(.X + (SegWidth \ 2) + 1, .Y + (SegHeight \ 2) - 1)
       Case 82
            P.Line ((.X + (SegWidth \ 2)) - 1, .Y + (SegHeight \ 2) + 3)-(.X + (SegWidth \ 2) - 1, .Y + SegHeight - 1)
            P.Line ((.X + (SegWidth \ 2)) + 0, .Y + (SegHeight \ 2) + 2)-(.X + (SegWidth \ 2) + 0, .Y + SegHeight - 0)
            P.Line ((.X + (SegWidth \ 2)) + 1, .Y + (SegHeight \ 2) + 3)-(.X + (SegWidth \ 2) + 1, .Y + SegHeight - 1)
       Case 91
            P.Line (.X + 3, .Y + 2)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) - 6)
            P.Line (.X + 3, .Y + 3)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) - 5)
            P.Line (.X + 3, .Y + 4)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) - 4)
            P.Line (.X + 3, .Y + 5)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) - 3)
            P.Line (.X + 3, .Y + 6)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) - 2)
       Case 93
            P.Line (.X + 3, .Y + SegHeight - 2)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) + 6)
            P.Line (.X + 3, .Y + SegHeight - 3)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) + 5)
            P.Line (.X + 3, .Y + SegHeight - 4)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) + 4)
            P.Line (.X + 3, .Y + SegHeight - 5)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) + 3)
            P.Line (.X + 3, .Y + SegHeight - 6)-(.X + (SegWidth \ 2) - 3, .Y + (SegHeight \ 2) + 2)
       Case 92
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) - 6)-(.X + SegWidth - 3, .Y + 2)
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) - 5)-(.X + SegWidth - 3, .Y + 3)
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) - 4)-(.X + SegWidth - 3, .Y + 4)
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) - 3)-(.X + SegWidth - 3, .Y + 5)
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) - 2)-(.X + SegWidth - 3, .Y + 6)
       Case 94
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) + 6)-(.X + SegWidth - 3, .Y + SegHeight - 2)
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) + 5)-(.X + SegWidth - 3, .Y + SegHeight - 3)
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) + 4)-(.X + SegWidth - 3, .Y + SegHeight - 4)
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) + 3)-(.X + SegWidth - 3, .Y + SegHeight - 5)
            P.Line (.X + (SegWidth \ 2) + 3, .Y + (SegHeight \ 2) + 2)-(.X + SegWidth - 3, .Y + SegHeight - 6)
    End Select
    End With
End Sub
Public Property Let Caption(ByVal Value As String)
    'Imposta la caption.
    Dim OrigX As Integer
    OrigX = BasePoint.X
    P.Cls
    With BasePoint
        While Value <> ""
            Select Case Left$(Value, 1)
                Case "."
                    P.Line (.X + (SegWidth \ 2) - (SegHeight / 14), .Y + (SegHeight))-(.X + (SegWidth \ 2), .Y + (SegHeight) - (SegHeight / 14)), , BF
                    .X = .X + SegWidth - 2
                Case ":"
                    P.Line (.X + (SegWidth \ 2) - (SegHeight / 14), .Y + (SegHeight))-(.X + (SegWidth \ 2), .Y + (SegHeight) - (SegHeight / 14)), , BF
                    .X = .X + SegWidth - 2
                Case ","
                    P.Line (.X + (SegWidth \ 2) - (SegHeight / 14), .Y + (SegHeight))-(.X + (SegWidth \ 2), .Y + (SegHeight) - (SegHeight / 14)), , BF
                    .X = .X + SegWidth - 2
                Case Else
                    DrawChar (Left$(Value, 1))
                    .X = .X + SegWidth + 5
            End Select
            Value = Right$(Value, Len(Value) - 1)
        Wend
        .X = OrigX
    End With
End Property
Property Let ForeColor(Color As Long)
    'Imposta il colore della caption.
    P.ForeColor = Color
End Property

Public Sub NewLCD(PBox As PictureBox)
    'Prepara la PictureBox per la visualizzazione.
    Set P = PBox
    P.ScaleMode = vbPixels
    P.AutoRedraw = True
    BasePoint.X = 2
    BasePoint.Y = 2
    SegHeight = P.ScaleHeight - 6
    SegWidth = (SegHeight \ 2) + 3
End Sub


