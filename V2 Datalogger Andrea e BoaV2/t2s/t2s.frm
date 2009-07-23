VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form t2s 
   Caption         =   "t2s"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "GET DATA"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox TEXT_IP 
      Height          =   270
      Left            =   6120
      TabIndex        =   2
      Text            =   "blackpnt.dyndns.org"
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton CMD_EXIT 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox TEXT_MAIN 
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton CMD_CONNECT 
      Caption         =   "CONNECT"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton CMD_CLOSE 
      Caption         =   "DISCONNECT"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox TEXT_PORT 
      Height          =   270
      Left            =   6120
      TabIndex        =   0
      Text            =   "1470"
      Top             =   1200
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock SOCK 
      Left            =   2880
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "REMOTE IP"
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "TCP PORT"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "t2s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CloseFile As String
Dim Filopen As Boolean
Dim FileNumber As Integer
Dim Buffer As String
Dim StopAll As Boolean


Private Sub CMD_CONNECT_Click()
    SOCK.RemoteHost = TEXT_IP
    SOCK.RemotePort = TEXT_PORT
    SOCK.Connect
    CMD_CONNECT.Enabled = False
    CMD_CLOSE.Enabled = True
    TEXT_MAIN.SelText = "Connecting to " + TEXT_IP + "...." + vbCrLf
End Sub


Private Sub CMD_CLOSE_Click()
    SOCK.SendData CloseFile
    StopAll = False
    CMD_CONNECT.Enabled = True
    CMD_CLOSE.Enabled = False
    TEXT_MAIN.SelText = "Closed." + vbCrLf
    If SOCK.State <> sckClosed Then SOCK.Close
End Sub


Private Sub CMD_EXIT_Click()
    End
End Sub


Private Sub Command1_Click()
    Dim FileName As String
    
    Dim GetAlive As String
    Dim Openfile As String
    Dim Continue As String
    'Dim CloseFile As String
    Dim Buffer As String
    Dim DataString As String
    Dim Prompt As Byte
    Dim FineBuffer As Boolean
    Dim counter As Integer

    FileName = App.Path + "\" + GeneraNome + "logs.txt"
    FileNumber = 1
    
    GetAlive = "V" & vbCr
    Openfile = "O1R/logs.txt" & vbCr
    Continue = "R1" & vbCr
    CloseFile = "C1" & vbCr

    Open FileName For Output As FileNumber
    Filopen = True
    
    SOCK.SendData GetAlive
    Debug.Print "getting alive"
    Debug.Print Buffer
    Sleep 1000
    
    SOCK.SendData GetAlive
    Debug.Print "getting alive"
    Debug.Print Buffer
    Sleep 1000
    
    SOCK.SendData Openfile
    Debug.Print "OpenFile"
    Sleep 1000
    'wait
    counter = 0
    StopAll = True
    Do
        DoEvents
        'Debug.Print "getting frame"
        'SOCK.SendData Continue
'        Buffer = InputComTimeOutTerm(10, 62)
'        Debug.Print Buffer
        If Buffer <> "" Then
            Buffer = Left(Buffer, Len(Buffer) - 1)
        End If
        If Buffer = "E07" Then
            FineBuffer = True
        Else
            DoEvents
'            DataString = DataString + Buffer
            'counter = counter + 1
            'Debug.Print "Counter->" & counter
        End If
    'sleep 500
    DoEvents
    Loop Until FineBuffer = True Or StopAll = False
    
    
    'SOCK.GetData msg, vbString, bytesTotal
    'TEXT_MAIN.SelText = msg
    Filopen = False
    Close #1
    
End Sub

Private Sub Form_Load()
    Filopen = False
End Sub

Private Sub SOCK_Connect()
    TEXT_MAIN.SetFocus
    TEXT_MAIN.SelText = "Connected." + vbCrLf
End Sub


Private Sub SOCK_Close()
    StopAll = False
    CMD_CONNECT.Enabled = True
    CMD_CLOSE.Enabled = False
    SOCK.Close
    TEXT_MAIN.SelText = "Closed." + vbCrLf
End Sub


Private Sub SOCK_DataArrival(ByVal bytesTotal As Long)
    Dim msg As String

    SOCK.GetData msg, vbString, bytesTotal
    'Open App.Path + "\" + GeneraNome + "debug.txt" For Output As #1
    'Print #1, msg
    Debug.Print msg
    TEXT_MAIN.SelText = msg
    Buffer = msg
    If Filopen = True Then
        Print #FileNumber, msg
        
    End If
End Sub

Private Sub SOCK_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    TEXT_MAIN.SelText = "Socket error: " + Description + vbCrLf
    CMD_CONNECT.Enabled = True
    CMD_CLOSE.Enabled = False
    SOCK.Close
End Sub


Private Sub TEXT_MAIN_KeyPress(KeyAscii As Integer)
    If SOCK.State <> sckClosed Then SOCK.SendData KeyAscii
    KeyAscii = 0  ' disable local echo
End Sub

Public Function GeneraNome() As String
    GeneraNome = Format(Year(Now), "0000")
    GeneraNome = GeneraNome + Format(Month(Now), "00")
    GeneraNome = GeneraNome + Format(Day(Now), "00")
    GeneraNome = GeneraNome + "_"
    GeneraNome = GeneraNome + Format(Hour(Now), "00")
    GeneraNome = GeneraNome + Format(Minute(Now), "00")
    GeneraNome = GeneraNome + Format(Second(Now), "00")
End Function

