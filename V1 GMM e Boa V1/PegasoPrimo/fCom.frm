VERSION 5.00
Begin VB.Form fCom 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porta seriale"
   ClientHeight    =   3420
   ClientLeft      =   4155
   ClientTop       =   3495
   ClientWidth     =   4305
   Icon            =   "fCom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton oCom1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COM 4"
      Height          =   495
      Index           =   4
      Left            =   1440
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton oCom1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COM 3"
      Height          =   495
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton oCom1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COM 2"
      Height          =   495
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton oCom1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COM 1"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton oCom1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COM 1"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton bFine 
      Caption         =   "&Ok"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Istituto Nazionale di Geofisica e Vulcanologia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "fCom.frx":0CCA
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pegaso"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1140
      TabIndex        =   3
      Top             =   960
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selezionare la porta di comunicazione seriale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   855
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
End
Attribute VB_Name = "fCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Dim i As Integer
    'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    'Eventualmente mettere qui un test sulle com esistenti
    oCom1(1).Value = True
    bFine.Enabled = True
    'Label3.Caption = Versione
    i = Val(sReadINI("Cavo", "UltimaCom", FileIni))
    If i = 0 Then i = 1
    oCom1(i).Value = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        CloseCom
        ComPort = 0
        Me.Hide
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub bFine_Click()
If fMain.MSComm1.PortOpen = False Then
    fMain.MSComm1.CommPort = ComPort
    fMain.MSComm1.Settings = "19200,n,8,1"
    fMain.MSComm1.InBufferSize = 2048
   'Altri settaggi com
    fMain.MSComm1.Handshaking = comNone
    fMain.MSComm1.RTSEnable = False

    'fMain.bProgramma.Enabled = True
    'fMain.bScarica.Enabled = True
    'fMain.bTestSensori.Enabled = True
End If
    
    Unload Me
    fMain.Show
End Sub

Private Sub oCom1_Click(Index As Integer)
    ComPort = Index
    bFine.Enabled = True
End Sub
