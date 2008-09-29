VERSION 5.00
Begin VB.Form YmodemRx 
   Caption         =   "Receiving File"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "YmodemRx.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bAbort 
      Caption         =   "&Abort"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lRetry 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label label5 
      Caption         =   "Retry Errors"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lTotalPackets 
      Caption         =   "of"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label labPacket 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Packet"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lSize 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Size"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "YmodemRx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This don't work in batch mode!!!!

'Public Const STARTY As String = "C"
'Public Const SOH As Byte = 1
'Public Const STX As Byte = 2
'Public Const EOT As Byte = 4
'Public Const ACK As Byte = 6
'Public Const NAK As Byte = 15
'Public Const CAN As Byte = 18
'Public Const MAXTRY As Byte = 3
'Public Const xyBufferSize As Integer = 1024

Public FirstByte As Byte
Public PacketNumber As Byte
Public PacketNumber2 As Byte
Public CRClo As Byte
Public CRChi As Byte
Public CRC As Long

Public Sub YModemDownload()
    Dim Packet As String
    Dim lPacket As Integer
    Dim PacketNumbers As Long 'Number of packets to be received
    Dim PacketNum As Byte    'Number of actual packet received
    Dim PacketNumProx As Byte 'Number of packet to be received
    Dim Retry As Byte
    Dim PacketData As String
    Dim filename As String
    Dim FileSize As Long
    Dim lPacketUsed As Integer
    Dim nFile As Long
    Dim LastPacketSize As Integer
    Dim CRCPacket As Long
    Dim CRCPacketCalc As Long
    'Dim LastPacketNumber As Integer
    Dim Stringa As String
    'Start YMODEM transaction
    Retry = 0
    PacketNumProx = 0
    
Riprova1:
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = STARTY

    lRetry.Caption = Str(Retry)
    Text1.Text = "Starting communication with C" + vbCrLf + "waiting for packet 0" + vbCrLf
Riprova2:
    Packet = ReceivePacket128
    
    nFile = FreeFile
'    Open "FirstPacket3.dat" For Output As #nFile
'    Stringa = Char2ascii(Packet)
'    'Debug.Print Stringa
'    Print #nFile, Stringa;
'    Close nFile

    
    lPacket = Len(Packet)
    'Control for timeouts and short packet
    Select Case lPacket
        Case Is = 0
            Retry = Retry + 1
            lRetry.Caption = Str(Retry)
            Debug.Print "Retry "; lRetry
            lRetry.Caption = Str(Retry)
            If Retry >= 5 Then
                AbortForTimeout
                Esci
                Exit Sub
            Else
                GoTo Riprova1
            End If
         Case 1 To 132
         Debug.Print "Packet too short-->"; lPacket
            Text1.Text = Text1.Text + "Packet too short-->" & lPacket & vbCrLf
            
            Retry = Retry + 1
            lRetry.Caption = Str(Retry)
            Debug.Print "Retry "; lRetry
            If Retry >= 5 Then
                AbortForTimeout
                Esci
                Exit Sub
            Else
                fMain.MSComm1.Output = Chr(NAK)
                GoTo Riprova2
            End If

         Case 133
            'all OK
    End Select
    Text1.Text = Text1.Text + "Packet received! It is long "
    Text1.Text = Text1.Text + Str(lPacket) + " bytes." + vbCrLf

    'Text1.Text = Text1.Text + Packet + vbCrLf
    FirstByte = GetByte(Packet, 1) 'Asc(Left(Packet, 1))
    'Check First Byte
    If FirstByte = SOH Then
        Debug.Print "First byte OK!"
        Text1.Text = Text1.Text + "First byte OK!" + vbCrLf
    Else
        Debug.Print "First byte is " + FirstByte
        Text1.Text = Text1.Text + "Error on first byte!" + vbCrLf
        fMain.MSComm1.Output = Chr(NAK)
        Retry = 0
        GoTo Riprova2
    End If
    
    PacketNumber = GetByte(Packet, 2)
    PacketNumber2 = GetByte(Packet, 3)
    'Check Packet number
    If PacketNumber = 0 Then
        Debug.Print "Packet Number OK!"
    Else
        Debug.Print "Packet Number is wrong" + PacketNumber
        Text1.Text = Text1.Text + "Packet Number is wrong" + PacketNumber + vbCrLf
        fMain.MSComm1.Output = Chr(NAK)
        Retry = 0
        GoTo Riprova2

    End If
    
    If PacketNumber2 = 255 Then
        Debug.Print "Packet Number2 OK!"
    Else
        Debug.Print "Packet Number2 is wrong " + PacketNumber2
        Text1.Text = Text1.Text + "Packet Number2 is wrong " + PacketNumber2 + vbCrLf
        fMain.MSComm1.Output = Chr(NAK)
        Retry = 0
        GoTo Riprova2
    End If

    PacketData = Mid(Packet, 4, 128)
    'Debug.Print PacketData
    CRChi = GetByte(Packet, 132)
    CRClo = GetByte(Packet, 133)
    Debug.Print "CRChi is "; CRChi
    Debug.Print "CRClo is "; CRClo
    CRCPacket = CLng(CRChi) * 256 + CRClo
    Debug.Print "CRC Packet is "; CRCPacket
    CRC16Setup
    CRCPacketCalc = CRC16(PacketData)
    Debug.Print "CRC calculated with CRC16 is "; CRCPacketCalc
'    CRCPacketCalc = CRC16ter(PacketData)
'    Debug.Print "CRC calculated with CRC16ter is "; CRCPacketCalc
'    CRCPacketCalc = CRC16A(PacketData)
'    Debug.Print "CRC calculated with CRC16a is "; CRCPacketCalc
'    CRCPacketCalc = CalcCRC(PacketData)
'    Debug.Print "CRC calculated with CalcCrc is "; CRCPacketCalc

   
    'check CRC
    If CRCPacket <> CRCPacketCalc Then
        Debug.Print "CRC failed!!"
    
    End If
    
    labPacket.Caption = Str(PacketNumber)
    
'    'Save first packet
'    nFile = FreeFile
'    Open "FirstPacket2.dat" For Output As #nFile
'    Print #nFile, PacketData;
'    Close nFile
'
'    nFile = FreeFile
'    Open "FirstPacket3.dat" For Output As #nFile
'    Stringa = Char2ascii(PacketData)
'    'Debug.Print Stringa
'    Print #nFile, Stringa;
'    Close nFile
    
    Stringa = ""
    filename = ""
    
    'get FileName and size
    GetFileName PacketData, filename, Stringa
    lFileName.Caption = filename
    lSize.Caption = Stringa
    FileSize = Val(Stringa)
    

    
    Text1.Text = Text1.Text + "First packet is ok." + vbCrLf + "Getting file." + vbCrLf
    'Receive next packets
    Retry = 0
    lRetry.Caption = Str(Retry)

    fMain.MSComm1.Output = Chr(ACK)
    fMain.MSComm1.Output = STARTY
    PacketNumProx = 1
    Debug.Print "Filename is "; filename
    filename = sGetAppPath + filename
    nFile = FreeFile
    Open filename For Output As #nFile
ciclo1:
    
    
    
    'Controllo e azioni sulla lunghezza del packet come prima
    
    fMain.MSComm1.InputLen = 1
    
    'FirstByte = GetByte(Packet, 1) 'Asc(Left(Packet, 1))
    Stringa = InputComTimeOutBin(3, 1)
    Debug.Print "Get first byte"
    FirstByte = GetByte(Stringa, 1)
    'Check First Byte
    
    If FirstByte = SOH Then
        lPacketUsed = 128
        Debug.Print "SOH"
    End If
    If FirstByte = STX Then
        lPacketUsed = 1024
        Debug.Print "STX"
    End If
    
    'Debug.Print "Lenght of expected packets="; lPacketUsed
    If FirstByte = EOT Then
        Debug.Print "EOT received"
        fMain.MSComm1.Output = Chr(ACK)
        fMain.MSComm1.Output = STARTY
        fMain.MSComm1.Output = Chr(ACK)
        'close file
        Close nFile
        Text1.Text = Text1.Text + "File received!" + vbCrLf
        Unload Me
        Exit Sub
    End If
    
Riprova3:

    If lPacketUsed = 128 Then
        Packet = ReceivePacket128
    Else
        Packet = ReceivePacket1024
    End If
    
    lPacket = Len(Packet)
    Debug.Print "Received "; lPacket

    If lPacketUsed = 128 Then
        If lPacket < 132 Then
            Debug.Print " packet too short"; lPacket
            fMain.MSComm1.Output = Chr(NAK)
            GoTo Riprova3

        End If
    Else
        If lPacket < 1028 Then
            Debug.Print " packet too short"; lPacket
            fMain.MSComm1.Output = Chr(NAK)
            GoTo Riprova3
        End If

    End If
   
    
    
    'This can change during transmission!!!
    PacketNumbers = Int(FileSize / lPacketUsed) + 1
    LastPacketSize = FileSize - (PacketNumbers - 1) * lPacketUsed 'Se cambia lPacketUsed il conto va a puttane
    lTotalPackets.Caption = "of " + Str(PacketNumbers)

    Debug.Print "Packet is long "; Len(Packet)
    PacketNumber = GetByte(Packet, 1)
    Debug.Print "Packet number "; PacketNumber
    PacketNumber2 = GetByte(Packet, 2)
    'Check Packet number
    'Check Packet number
    If PacketNumber = PacketNumProx Then
        Debug.Print "Packet Number OK!"
    Else
        Debug.Print "Error Packet Number is "; PacketNumber; "instead of "; PacketNumProx
    End If
    
    If PacketNumber2 = 255 - PacketNumProx Then
        Debug.Print "Packet Number2 OK!"
    Else
        Debug.Print "Packet Number2 is "; PacketNumber2
    End If
    
    
    
    PacketData = Mid(Packet, 3, lPacketUsed)
    If PacketNumber = PacketNumbers Then
        'E' l'ultimo pacchetto, facciamo il trim
        PacketData = Left(PacketData, LastPacketSize)
    End If
    Print #nFile, PacketData;
    'Debug.Print PacketData
    If Len(Packet) > 133 Then
        'Packet is 1024 long
        CRChi = GetByte(Packet, 1027)
        CRClo = GetByte(Packet, 1028)

    Else
        'Packet is 128 long
        CRChi = GetByte(Packet, 131)
        CRClo = GetByte(Packet, 132)
    End If

    
    'check CRC
    CRCPacket = CLng(CRChi) * 256 + CRClo
    Debug.Print "CRC Packet is "; CRCPacket
    CRC16Setup
    CRCPacketCalc = CRC16(PacketData)
    Debug.Print "CRC calculated with CRC16 is "; CRCPacketCalc
    CRCPacketCalc = CRC16ter(PacketData)
    Debug.Print "CRC calculated with CRC16ter is "; CRCPacketCalc
    CRCPacketCalc = CRC16A(PacketData)
    Debug.Print "CRC calculated with CRC16a is "; CRCPacketCalc
    CRCPacketCalc = CalcCRC(PacketData)
    Debug.Print "CRC calculated with CalcCrc is "; CRCPacketCalc
    
    
    
    labPacket.Caption = Str(PacketNumber)


    
    Text1.Text = Text1.Text + "Packet " + Str(PacketNumber) + " received! It is long "
    Text1.Text = Text1.Text + Str(lPacket) + " bytes." + vbCrLf
    Text1.Text = Text1.Text + "Getting next" + vbCrLf
    PacketNumProx = PacketNumProx + 1
    fMain.MSComm1.Output = Chr(ACK)
    GoTo ciclo1

    fMain.MSComm1.Output = Chr(CAN)
    fMain.MSComm1.Output = Chr(CAN)
End Sub

Private Function ReceivePacket128() As String
    Dim Packet As String
    ReceivePacket128 = InputComTimeOutBin(5, 133)
    'Debug.Print "inputCom--->"; Char2ascii(ReceivePacket128)
End Function

Private Function ReceivePacket1024() As String
    Dim Packet As String
    ReceivePacket1024 = InputComTimeOutBin(5, 1028)
    'Debug.Print "inputCom--->"; Char2ascii(ReceivePacket128)
End Function

Private Sub Esci()
    Unload Me
    fMain.MousePointer = vbDefault
    'fMain.MSComm1.InputMode = comInputModeText
    'fMain.AbilitaTasti
    fMain.Show

End Sub

Private Sub bAbort_Click()
    Esci
End Sub

Private Function GetByte(Stringa As String, Position As Long) As Byte
    'Debug.Print "Stringa "; Len(Stringa)
    If Stringa = "" Then
        GetByte = 0
        Exit Function
    End If
    If Position > Len(Stringa) Then
        GetByte = 0
        Exit Function
    End If
    
    GetByte = Asc(Mid(Stringa, Position, 1))
End Function

Private Sub AbortForTimeout()
    Text1.Text = Text1.Text + "Max retry exceeded, aborting"
    MsgBox "Max retry exceeded, aborting", vbOKOnly, "ERROR!"
End Sub

Private Sub GetFileName(Packet As String, filename As String, Size As String)
    Dim Actual As Byte
    Dim iPacket As Long
    iPacket = 1
    
'    Dim i As Long
'    i = FreeFile
'    Open "FirstPacket.dat" For Output As #i
'    Print #i, Packet;
'    Close i
    
    'Get File Name
    Do
        Actual = GetByte(Packet, iPacket)
        filename = filename + Chr(Actual)
        'Debug.Print FileName; " "; Actual
        iPacket = iPacket + 1
    Loop Until Actual = 0
    'Debug.Print FileName
    'Debug.Print Len(FileName)
    filename = Left(filename, Len(filename) - 1)
    'Debug.Print "GetFileName FileName "; filename
    'Debug.Print Len(FileName)
    
    'Get Lenght
    Do
        Actual = GetByte(Packet, iPacket)
        Size = Size + Chr(Actual)
        iPacket = iPacket + 1
    Loop Until Actual = 32 Or Actual = 0
    'Debug.Print Size
    'Debug.Print Len(Size)
    Size = Left(Size, Len(Size) - 1)
    'Debug.Print "GetFileName size ";Size
    'Debug.Print Len(Size)

End Sub
