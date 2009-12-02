VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form t2s 
   Caption         =   "t2s"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bGetStatus 
      Caption         =   "Z"
      Height          =   375
      Left            =   7320
      TabIndex        =   22
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton bQueryS1 
      Caption         =   "S1"
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton bGetAlive 
      Caption         =   "GET ALIVE"
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton bSendCR 
      Caption         =   "CR"
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton bClearText 
      Caption         =   "CLR"
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton bQueryInformation 
      Caption         =   "QUERY INFORMATION"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6120
      TabIndex        =   17
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton bCreateFile 
      Caption         =   "CREATE FILE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton bFreeFile 
      Caption         =   "FREE FILE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton bGetInfo 
      Caption         =   "GET INFO"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLOSE FILE"
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GET LAST FRAME"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton bOpenFile 
      Caption         =   "OPEN FILE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton bGetFrame 
      Caption         =   "GET FRAME"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton bDelete 
      Caption         =   "DELETE FILE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   5880
      Width           =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5160
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET DATA"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox TEXT_IP 
      Height          =   270
      Left            =   6120
      TabIndex        =   2
      Text            =   "blackpnt.dyndns.org"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton CMD_EXIT 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox TEXT_MAIN 
      Height          =   6975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   0
      Width           =   5655
   End
   Begin VB.CommandButton CMD_CONNECT 
      Caption         =   "CONNECT"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton CMD_CLOSE 
      Caption         =   "DISCONNECT"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox TEXT_PORT 
      Height          =   270
      Left            =   6120
      TabIndex        =   0
      Text            =   "1470"
      Top             =   720
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
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "TCP PORT"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "t2s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CloseFile As String
Dim Filopen As Boolean
Dim FileNumber As Integer
Dim FileNumber2 As Integer
Dim Buffer As String
Dim StopAll As Boolean

    Dim GetAlive As String
    Dim Openfile As String
    Dim Continue As String
    Dim InfoFile As String
    Dim CreateFile As String
    Dim FreeFile2 As String
    Dim QueryInformation As String


Private Sub bClearText_Click()
    TEXT_MAIN.Text = ""
End Sub

Private Sub bCreateFile_Click()
    SOCK.SendData CreateFile
End Sub

Private Sub bDelete_Click()
    TEXT_MAIN.Text = "Deleting file"
    SOCK.SendData "E /LOGS.TXT" & vbCr
    Sleep 1000
    TEXT_MAIN.Text = TEXT_MAIN.Text & vbCrLf & Buffer
End Sub

Private Sub bFreeFile_Click()
    SOCK.SendData FreeFile2
End Sub

Private Sub bGetAlive_Click()
    SOCK.SendData GetAlive
End Sub

Private Sub bGetFrame_Click()

    SOCK.SendData Continue
End Sub

Private Sub bGetInfo_Click()
    SOCK.SendData InfoFile

End Sub

Private Sub bGetStatus_Click()
    SOCK.SendData "Z" + vbCr
End Sub

Private Sub bOpenFile_Click()
    SOCK.SendData Openfile
End Sub

Private Sub bQueryInformation_Click()
    SOCK.SendData QueryInformation
End Sub

Private Sub bQueryS1_Click()
    SOCK.SendData "S 1" + vbCr
    
End Sub

Private Sub bSendCR_Click()
    SOCK.SendData vbCr
End Sub

Private Sub CMD_CONNECT_Click()
    SOCK.RemoteHost = TEXT_IP
    SOCK.RemotePort = TEXT_PORT
    SOCK.Connect
    CMD_CONNECT.Enabled = False
    CMD_CLOSE.Enabled = True
    bDelete.Enabled = True
    Command1.Enabled = True
    bGetFrame.Enabled = True
    bOpenFile.Enabled = True
    bFreeFile.Enabled = True
    bCreateFile.Enabled = True
    bGetInfo.Enabled = True
    bQueryInformation.Enabled = True
    Filopen = False
    
    GetAlive = "V" & vbCr
    Openfile = "O 1 R /LOGS.TXT" & vbCr
    Continue = "R 1" & vbCr
    CloseFile = "C 1" & vbCr
    InfoFile = "I 1" & vbCr
    FreeFile2 = "F" + vbCr
    CreateFile = "O 1 A /LOGS.TXT" & vbCr
    QueryInformation = "Q" + vbCr

    TEXT_MAIN.SelText = "Connecting to " + TEXT_IP + "...." + vbCrLf
End Sub


Private Sub CMD_CLOSE_Click()
    SOCK.SendData CloseFile
    Close FileNumber
    Close FileNumber2
    StopAll = False
    CMD_CONNECT.Enabled = True
    CMD_CLOSE.Enabled = False
    bDelete.Enabled = False
    Command1.Enabled = False
    bOpenFile.Enabled = False
    bGetFrame.Enabled = False
    TEXT_MAIN.SelText = "Closed." + vbCrLf
    If SOCK.State <> sckClosed Then SOCK.Close
End Sub


Private Sub CMD_EXIT_Click()
    End
End Sub


Private Sub Command1_Click()
    Dim FileName As String
    Dim FileName2 As String
    
    
    'Dim CloseFile As String
    Dim Buffer As String
    Dim DataString As String
    Dim Prompt As Byte
    Dim FineBuffer As Boolean
    Dim counter As Integer

    'MSComm1.CommPort = 1
    'MSComm1.Handshaking = comNone
    'MSComm1.PortOpen = True

    FileName = App.Path + "\" + GeneraNome + "logs.txt"
    FileNumber = FreeFile
    
    
    
'    GetAlive = "V" & vbCr
'    Openfile = "O 1 R /LOGS.TXT" & vbCr
'    Continue = "R 1" & vbCr
'    CloseFile = "C 1" & vbCr
'    InfoFile = "I 1" & vbCr
'    CreateFile = ""

    Open FileName For Output As FileNumber
    Filopen = True
    
    FileName2 = App.Path + "\" + GeneraNome + " mod.txt"
    FileNumber2 = FreeFile

    Open FileName2 For Output As FileNumber2
    
    
    SOCK.SendData GetAlive
    'MSComm1.Output = GetAlive
    
    Debug.Print "getting alive"
    
    Sleep 1000
'    Do
'        do events
'    Loop Until Buffer <> ""
    Debug.Print Buffer

    
    SOCK.SendData GetAlive
    'MSComm1.Output = GetAlive

    Debug.Print "getting alive"
    Debug.Print Buffer
    Sleep 1000
    
    SOCK.SendData CloseFile
    Sleep 1000
    
    SOCK.SendData Openfile
    'MSComm1.Output = Openfile
    Debug.Print "OpenFile"
    Sleep 1000
    'wait
    counter = 0
    StopAll = True
    
    SOCK.SendData InfoFile
    'MSComm1.Output = Openfile
    Debug.Print "InfoFile"
    Sleep 1000
    Debug.Print Buffer
    Print #FileNumber, Buffer
    
    Do
        DoEvents
        Debug.Print "getting frame"

        SOCK.SendData Continue
        'SOCK.SendData "R 3 512 3998043"
        'MSComm1.Output = Continue
        'Buffer = InputComTimeOutTerm(10, 62)
        Sleep 1000
        'Buffer = MSComm1.Input
        Do
            DoEvents
        Loop Until Buffer <> ""
        Debug.Print Buffer
        If Buffer <> "" Then
            Print #FileNumber, Buffer;
            'Buffer = Left(Buffer, Len(Buffer) - 1)
        End If
        If Len(Buffer) > 6 Then
            Print #FileNumber2, Buffer;
        End If
        If Buffer = "E07>" Then
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
    If SOCK.State = sckClosed Then Exit Sub
    SOCK.SendData CloseFile
    Filopen = False
    Close #FileNumber
    Close #FileNumber2
    
End Sub

Private Sub Command2_Click()
    SOCK.SendData "R 1 512 512"
End Sub

Private Sub Command3_Click()
    SOCK.SendData CloseFile
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

