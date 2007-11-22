VERSION 5.00
Begin VB.Form fFindModem 
   Caption         =   "Find Modem"
   ClientHeight    =   3195
   ClientLeft      =   3690
   ClientTop       =   3195
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3450
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cChooseModem 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Text            =   "Choose a modem"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox tModem 
      Height          =   1695
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox tCOM 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lMessage 
      Alignment       =   2  'Center
      Caption         =   "Push start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Serials"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "fFindModem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bStart_Click()
Dim i As Integer
Dim j As Integer
Dim coms(10) As Integer
Dim risposta As String

lMessage.Caption = "WAIT!"

'Find all coms from 1 to 10
On Error GoTo errore
j = 1
For i = 1 To 10
    ComOk = True
    fMain.MSComm1.CommPort = i
    fMain.MSComm1.PortOpen = True
        If ComOk = True Then
            tCOM = tCOM + "COM" + Str(i) + vbCrLf
            coms(j) = i
            j = j + 1
            fMain.MSComm1.PortOpen = False
        End If
Next i

If j > 1 Then
    For i = 1 To j - 1
        fMain.MSComm1.CommPort = coms(i)
        fMain.MSComm1.PortOpen = True
        risposta = MandaComando("ATI", 1)
        If risposta <> "" Then
            tModem = tModem + risposta + vbCrLf
            'aggiungere il modem alla lista
        Else
            tModem = tModem + "no modem found" + vbCrLf
        End If
    Next i
End If

lMessage.Caption = "Choose"

errore:
    ComOk = False
    Resume Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    fMain.Show
End Sub
