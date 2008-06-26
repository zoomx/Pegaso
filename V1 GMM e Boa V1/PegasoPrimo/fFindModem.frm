VERSION 5.00
Begin VB.Form fFindModem 
   Caption         =   "Find Modem"
   ClientHeight    =   3195
   ClientLeft      =   3690
   ClientTop       =   3195
   ClientWidth     =   3450
   Icon            =   "fFindModem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cChooseModem 
      Enabled         =   0   'False
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
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
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

Private Sub bQuit_Click()
    Unload Me
    fMain.Show
End Sub

Private Sub bStart_Click()
Dim i As Integer
Dim J As Integer
Dim coms(10) As Integer
Dim risposta As String

lMessage.Caption = "WAIT!"
CloseCom
fMain.MSComm1.Handshaking = comRTS
fMain.MSComm1.Settings = "19200,n,8,1"
'Find all coms from 1 to 10
On Error GoTo errore
J = 1
For i = 1 To 20
    ComOk = True
    fMain.MSComm1.CommPort = i
    fMain.MSComm1.PortOpen = True
        If ComOk = True Then
            tCOM = tCOM + "COM" + Str(i) + vbCrLf
            coms(J) = i
            J = J + 1
            fMain.MSComm1.PortOpen = False
        End If
Next i

If J > 1 Then
    For i = 1 To J - 1
        fMain.MSComm1.CommPort = coms(i)
        fMain.MSComm1.PortOpen = True
        risposta = MandaComando("I", 1)
        'Debug.Print risposta
        If risposta <> "" Then
            tModem = tModem + risposta + vbCrLf
            'aggiungere il modem alla lista
            cChooseModem.AddItem ("COM" & Str(coms(i)))
            'oppure
            'cChooseModem.AddItem (risposta)
        Else
            tModem = tModem + "no modem found" + vbCrLf
        End If
        fMain.MSComm1.PortOpen = False
        'cChooseModem.AddItem ("COM" & Str(i))
    Next i
End If
cChooseModem.Enabled = True
lMessage.Caption = "Choose"

errore:
    ComOk = False
    Resume Next

End Sub

Private Sub cChooseModem_Click()
    Dim Stringa As String
    Dim Index As Integer
    Stringa = cChooseModem.Text
    'Debug.Print stringa
    Stringa = Mid(Stringa, 4)
    'Debug.Print stringa
    Index = Val(Stringa)
    'Debug.Print Index
    ComPort = Index
    Debug.Print "ComPort="; ComPort
    fMain.bModem.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    fMain.Show
End Sub

Private Sub tCOM_DblClick()
        Dim Stringa As String
    Stringa = tCOM.SelText
    Debug.Print "COM scelta "; Stringa
    ComPort = Val(Stringa)
    Debug.Print "ComPort="; ComPort
    fMain.bModem.Enabled = True
End Sub
