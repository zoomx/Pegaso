VERSION 5.00
Begin VB.Form FindModem 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
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
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox tCOM 
      Height          =   1695
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
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
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FindModem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, pPorts As Any, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Private Type PORT_INFO_2
   pPortName As Long
   pMonitorName As Long
   pDescription As Long
   fPortType As Long
   Reserved As Long
End Type
 
'PORT_INFO_2 fPortType-Konstanten
Private Const PORT_TYPE_WRITE = &H1 'Schreiben auf dem Port ist möglich
Private Const PORT_TYPE_READ = &H2 'Lesen des Ports ist möglich
Private Const PORT_TYPE_REDIRECTED = &H4 'Der Port ist im Offlinedruck Schemaist akteviert
Private Const PORT_TYPE_NET_ATTACHED = &H8 'Der Drucker ist ein Netzwerkdrucker
 
'eine der Standard Fehlerkonstanten
Private Const ERROR_CANCELLED = 1223
 
Dim Ports() As PORT_INFO_2

Private Sub bQuit_Click()
    Unload Me
    fMain.Show

End Sub

Private Sub bStart_Click()
   Dim Retval As Long, BufferSize As Long
   Dim CountPorts As Long
   Dim I As Integer
   
   Dim coms(10) As Integer
   Dim commNumber As Integer
   Dim commPort As String
   Dim j As Integer

   'Anzahl Ports ermitteln
   Retval = EnumPorts(vbNullChar, 2, ByVal 0&, 0&, BufferSize, CountPorts)
   
   'Puffer erstellen und Portinfos ermitteln
   ReDim Ports(BufferSize / Len(Ports(0)) + 1)
   Retval = EnumPorts(vbNullChar, 2, Ports(0), Len(Ports(0)) * (UBound(Ports) + 1), BufferSize, CountPorts)

    'Namen jedes Ports ermitteln
    j = 1

    For I = 0 To CountPorts - 1
        commPort = PtrToString(Ports(I).pPortName)
            If Left(commPort, 3) = "COM" Then
                'Combo1.AddItem PtrToString(Ports(I).pPortName)
                tCOM = tCOM + commPort + vbCrLf
                commNumber = Mid(commPort, 4, 1)
                Debug.Print commNumber
                coms(j) = commNumber
                j = j + 1
            End If
    Next I
 
End Sub


'String anhand eines Pointers ermitteln
Private Function PtrToString(ByVal StringPtr As Long) As String
   Dim TmpStr As String
   
   TmpStr = Space(lstrlen(StringPtr))
   Call lstrcpy(TmpStr, StringPtr)
   
   PtrToString = TmpStr
End Function

