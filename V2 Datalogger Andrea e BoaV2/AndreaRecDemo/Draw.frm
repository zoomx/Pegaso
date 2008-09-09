VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "CLICK FORM !!!"
   ClientHeight    =   4020
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Spiral"
      Height          =   420
      Left            =   5970
      TabIndex        =   2
      Top             =   1425
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "circle"
      Height          =   495
      Left            =   5940
      TabIndex        =   1
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "ray"
      Height          =   495
      Left            =   5940
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Cls
MoveTo Me.ScaleWidth / 2, Me.ScaleHeight / 2

For i = 1 To 360 Step 5
    DrawLine -100, i
Next i

End Sub

Private Sub Command2_Click()
Me.Cls
MoveTo Me.ScaleWidth / 2, Me.ScaleHeight / 2

For i = 1 To 360
    DrawLine 1, i
Next i

End Sub

Private Sub Command3_Click()
Me.Cls
MoveTo Me.ScaleWidth / 2, Me.ScaleHeight / 2

For i = 1 To 360 * 3
    DrawLine i / 200, 2 * i
Next i

End Sub

Private Sub DrawLine(ByVal length As Single, ByVal Angle As Single)

Pi = 4 * Atn(1)
cx = Me.CurrentX
cy = Me.CurrentY

'Angle is in Degrees
Angle = Angle Mod 360
Angle = Angle * Pi / 180
xp = 0
yp = Abs(length)
Rx = xp * Cos(Angle) - yp * Sin(Angle)
Ry = xp * Sin(Angle) + yp * Cos(Angle)
rxg = cx + Rx
ryg = cy - Ry

Line (cx, cy)-(rxg, ryg)

' if negative length go back to start position
If length < 0 Then
    Me.CurrentX = cx
    Me.CurrentY = cy
End If


End Sub

Private Sub Form_DblClick()
Me.Cls
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveTo X, Y

For i = Me.ScaleWidth / 2 To 0 Step -2
        r = Rnd * 15
        Me.ForeColor = QBColor(r)
        DrawLine -CSng(i), CSng(i)
Next i

End Sub

Private Sub MoveTo(X As Single, Y As Single)
Me.CurrentX = X
Me.CurrentY = Y
End Sub

