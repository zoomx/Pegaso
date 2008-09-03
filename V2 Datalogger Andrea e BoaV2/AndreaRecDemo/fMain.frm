VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form fMain 
   Caption         =   "Pegaso On Line"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   1560
   End
   Begin VB.CommandButton bDemoData 
      Caption         =   "Da&ta"
      Height          =   615
      Left            =   3960
      TabIndex        =   27
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton bFindModem 
      Caption         =   "&Find modem"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton bDownload 
      Caption         =   "&Download file"
      Height          =   495
      Left            =   2160
      TabIndex        =   24
      Top             =   5640
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4200
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton bConnectModem 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   960
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Text            =   "127.0.0.1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   5040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   480
      TabIndex        =   17
      Top             =   4440
      Width           =   3135
      Begin VB.PictureBox pDepth 
         Height          =   615
         Left            =   1560
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Depth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current meter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   5520
      TabIndex        =   10
      Top             =   240
      Width           =   4695
      Begin VB.PictureBox pFlow 
         Height          =   615
         Left            =   1560
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.PictureBox pPolar 
         Height          =   3910
         Left            =   240
         Picture         =   "fMain.frx":0CCA
         ScaleHeight     =   3855
         ScaleWidth      =   4065
         TabIndex        =   11
         Top             =   1080
         Width           =   4120
      End
   End
   Begin VB.CommandButton bTestLine 
      Caption         =   "Test Line"
      Height          =   495
      Left            =   9720
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton bConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Water"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox pTurbidity 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin VB.PictureBox pOxygen 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.PictureBox ppH 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.PictureBox pCond 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.PictureBox pTemperature 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Turbidity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Oxygen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "pH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Conducibility"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblMessages 
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   5760
      Width           =   6255
   End
   Begin VB.Label lblStatus 
      Caption         =   "Not Connected"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   6240
      Width           =   5055
   End
   Begin VB.Label Label7 
      Caption         =   "Remote Host"
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LCDtemp As New mcLCD
Private LCDcond As New mcLCD
Private LCDpH As New mcLCD
Private LCDflow As New mcLCD
Private LCDoxy As New mcLCD
Private LCDturb As New mcLCD
Private LCDdepth As New mcLCD

Private Sub bConnect_Click()
    If bConnect.Caption = "&Connect" Then
        'Inizializza la connessione.
        tcpClient.Close
        'Imposta il server specificato.
        tcpClient.RemoteHost = txtServer.Text
        'Utilizza la porta 100.
        tcpClient.RemotePort = 100
        lblStatus.Caption = "Connecting port " & tcpClient.RemotePort & "..."
        'InsertText lblStatus.Caption, Me
        'Tenta la connessione.
        tcpClient.Connect
    Else
        'Chiude la connessione.
        'InsertText "Connessione chiusa.", Me
        tcpClient.Close
        bConnect.Caption = "&Connect"
        lblStatus = "Not connected"
    End If
End Sub

Private Sub bConnectModem_Click()
    'apre porta
    'Controlla che ci sia un modem
    'chiama la rubrica
    'chiama modem remoto
    Me.Hide
    fModem.Show

End Sub

Private Sub bDemoData_Click()
    Static Switch As Boolean
'    Dim FileName As String
'    Dim FileNumber As Integer

    If Switch = False Then
        FileNumber = FreeFile
        FileName = App.Path + "\data.csv"
        Open FileName For Input As #FileNumber
        Timer1.Enabled = True
        Switch = True
    Else
        Timer1.Enabled = False
        Close FileNumber
        Switch = False
    End If
        
End Sub

Private Sub bDownload_Click()
    Dim GetAlive As String
    Dim OpenFile As String
    Dim Continue As String
    Dim CloseFile As String
    Dim Buffer As String
    Dim DataString As String
    Dim Prompt As Byte
    Dim FineBuffer As Boolean
    Dim counter As Integer
    
    
    GetAlive = "V" & vbCr
    OpenFile = "O1R/logs.txt" & vbCr
    Continue = "R1" & vbCr
    CloseFile = "C1" & vbCr
    
    
    Debug.Print "Getting alive"
    fMain.MSComm1.OutBufferCount = 0
    fMain.MSComm1.Output = GetAlive
    Buffer = InputComTimeOutTerm(10, 62)
    Debug.Print Buffer
    lblMessages.Caption = Buffer
    
    Debug.Print "Getting alive"
    fMain.MSComm1.OutBufferCount = 0
    fMain.MSComm1.Output = GetAlive
    Buffer = InputComTimeOutTerm(10, 62)
    Debug.Print Buffer
    lblMessages.Caption = Buffer

    
    Debug.Print "Opening file"
    fMain.MSComm1.Output = OpenFile
    Prompt = InputComTimeOutBin3(10)
    Buffer = Str(Prompt)
    Debug.Print Buffer
    lblMessages.Caption = Buffer
    
    
    FineBuffer = False
    counter = 0
    DataString = ""
    Do
        Debug.Print "getting frame"
        fMain.MSComm1.Output = Continue
        Buffer = InputComTimeOutTerm(10, 62)
        Debug.Print Buffer
        Buffer = Left(Buffer, Len(Buffer) - 1)
        If Buffer = "E07" Then
            FineBuffer = True
        Else
            DataString = DataString + Buffer
            counter = counter + 1
            Debug.Print "Counter->" & counter
        End If
        
    Loop Until FineBuffer = True
    
    
    Debug.Print "Closing file"
    fMain.MSComm1.Output = CloseFile
    Prompt = InputComTimeOutBin3(10)
    Buffer = Str(Prompt)
    Debug.Print Buffer
    lblMessages.Caption = Buffer
    
    Debug.Print "Closing communication"
    CloseCom
    lblStatus.Caption = "Not connected"
    
    Debug.Print "Saving file"
    FileNumber = FreeFile
    FileName = App.Path + "\Logs2.txt"
    Open FileName For Output As #FileNumber
    Print #FileNumber, DataString
    Close FileNumber
    
End Sub

Private Sub bFindModem_Click()
    Me.Hide
    fFindModem.Show

End Sub

Private Sub tcpClient_Connect()
    'Connessione riuscita.
    'InsertText "Connesso al server " & tcpClient.RemoteHost & " " & tcpClient.RemoteHostIP & ".", Me
    lblStatus.Caption = "Connected"
    bConnect.Caption = "&Disconnect"
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    'Sono in arrivo dati.
    Dim strData As String
    tcpClient.GetData strData, vbByte
    'Debug.Print bytesTotal
    ParseData strData
    'Debug.Print "Parsed"
'    If Left(strData, 7) = "\msgbox" Then
'    MsgBox Mid(strData, 8)
'    Else
'    InsertText strData, Me
'    End If
End Sub

Private Sub bEnd_Click()
    End
End Sub

Private Sub bTestLine_Click()
    DrawLine 250, 0
    DrawLine 500, 90
    DrawLine 1000, 180
    DrawLine 1500, 270
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    With LCDtemp
        .NewLCD pTemperature
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "16.1"
    End With
    With LCDcond
        .NewLCD pCond
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "4.71"
    End With
    With LCDpH
        .NewLCD ppH
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "8.0"
    End With
    With LCDflow
        .NewLCD pFlow
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "20.5"
    End With
    With LCDoxy
        .NewLCD pOxygen
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "3.6"
    End With
    With LCDturb
        .NewLCD pTurbidity
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "15.3"
    End With
    With LCDdepth
        .NewLCD pDepth
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "22.5"
    End With
    FileIni = sGetAppPath & "AndreaRec.ini"
    i = Val(sReadINI("Modem", "UltimaCom", FileIni))
    If i = 0 Then
        'There is not a COM port defined
        'fMain.bModem.Enabled = False
    End If
    ComPort = i
End Sub

Sub DrawLine(ByVal length As Single, ByVal Angle As Single)
    Dim Pi As Double
    Dim cx As Single
    Dim cy As Single
    Dim xp As Single
    Dim yp As Single
    Dim Rx As Single
    Dim Ry As Single
    Dim rxg As Single
    Dim ryg As Single
    
    
    Pi = 4 * Atn(1)
    cx = pPolar.Width / 2
    cy = pPolar.Height / 2
    
    'Angle is in Degrees
    Angle = Angle + 180
    Angle = Angle Mod 360
    
    Angle = Angle * Pi / 180
    xp = 0
    yp = Abs(length)
    Rx = xp * Cos(Angle) - yp * Sin(Angle)
    Ry = xp * Sin(Angle) + yp * Cos(Angle)
    rxg = cx + Rx
    ryg = cy - Ry
    
    ryg = pPolar.Height - ryg
    
    pPolar.ForeColor = vbRed
    pPolar.DrawWidth = 5
    pPolar.Line (cx, cy)-(rxg, ryg)
    
    'Line (cx, cy)-(rxg, ryg)
    
'    ' if negative length go back to start position
'    If length < 0 Then
'        Me.CurrentX = cx
'        Me.CurrentY = cy
'    End If
'

End Sub

Private Sub Timer1_Timer()
    Static stringa As String
    If Not EOF(FileNumber) Then
        Line Input #FileNumber, stringa
        Dati = Split(stringa, ";")
        If UBound(Dati) = 2 Then
            LCDtemp.Caption = Dati(0)
            LCDcond.Caption = Dati(1)
            LCDdepth.Caption = Dati(2)
        End If
    
    End If
    
End Sub
