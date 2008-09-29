VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form fMain 
   Caption         =   "Pegaso I"
   ClientHeight    =   6375
   ClientLeft      =   3405
   ClientTop       =   2340
   ClientWidth     =   8820
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bFindModem2 
      Caption         =   "FindModem2"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton bCrcTest 
      Caption         =   "CRC test"
      Height          =   495
      Left            =   3600
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton bClear 
      Caption         =   "Clear Txt"
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton bActivate 
      Caption         =   "Activate"
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton bCloseComm 
      Caption         =   "Close Communication"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton bGetTodayFiles 
      Caption         =   "Get Yesterday files"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   615
      Left            =   2520
      TabIndex        =   17
      Top             =   240
      Width           =   3615
      Begin VB.Label lStatus 
         Caption         =   "Not Connected"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton bTranslatePTM 
      Caption         =   "Translate PTM Data"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   5520
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1680
      Picture         =   "Main.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   510
      TabIndex        =   15
      Top             =   0
      Width           =   515
   End
   Begin VB.CommandButton bRymodem 
      Caption         =   "Receive &Ymodem"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   6480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton bTranslate 
      Caption         =   "&Translate PMM Data"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton bTerminal 
      Caption         =   "&Open Terminal"
      Height          =   735
      Left            =   5400
      Picture         =   "Main.frx":11CE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   855
      Left            =   7920
      Picture         =   "Main.frx":14D8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton bFindModem 
      Caption         =   "&Find Modem"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton bModem 
      Caption         =   "&Modem Call"
      Height          =   855
      Left            =   240
      Picture         =   "Main.frx":1922
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RTSEnable       =   -1  'True
   End
   Begin VB.CommandButton bSetDate 
      Caption         =   "&Set Date&&Time"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton bGetCurrentData 
      Caption         =   "&Get current Date&&Time"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   1800
      TabIndex        =   9
      Top             =   960
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   3735
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label labelText 
         Caption         =   "Operations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton bConnect 
      Caption         =   "&Connect GMM"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "PEGASO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bActivate_Click()
    Static eEnabled As Boolean
    
    If eEnabled = False Then
    
        bGetTodayFiles.Enabled = True
        bConnect.Enabled = True
        bGetTodayFiles.Enabled = True
        bRymodem.Enabled = True
        eEnabled = True
    Else
        bGetTodayFiles.Enabled = False
        bConnect.Enabled = False
        bGetTodayFiles.Enabled = False
        bRymodem.Enabled = False
        eEnabled = False
    End If

End Sub

Private Sub bClear_Click()
    Text1.Text = ""
End Sub

Private Sub bCloseComm_Click()
    CloseCom
    bConnect.Enabled = False
    bGetTodayFiles.Enabled = False
    lStatus.Caption = "Not Connected"
End Sub

Private Sub bCrcTest_Click()
    
    Dim sMessage As String
    Dim iCRC As Long
    'CRC16Setup
    'TestCRC16
    
    sMessage = "123456789"
    iCRC = CRC16(sMessage)
    Debug.Print "crc16(" & sMessage & ")=" & Hex(iCRC) & " BB3D"
    CRC16Setup
    iCRC = CRC16(sMessage)
    Debug.Print "crc16(" & sMessage & ")=" & Hex(iCRC) & " BB3D"
    iCRC = CRC16ter(sMessage)
    Debug.Print "crc16ter(" & sMessage & ")=" & Hex(iCRC) & " BB3D"
    iCRC = CRC16A(sMessage)
    Debug.Print "crc16a(" & sMessage & ")=" & Hex(iCRC) & " BB3D"
    iCRC = CalcCRC(sMessage)
    Debug.Print "CalcCRC(" & sMessage & ")=" & Hex(iCRC) & " BB3D"
    iCRC = GetCrc32(sMessage)
    Debug.Print "GetCrc32(" & sMessage & ")=" & Hex(iCRC) & " BB3D"

End Sub

Private Sub bEnd_Click()
    UnloadAllForms Me.Name
    Unload Me
    End
End Sub

Private Sub bFindModem_Click()
    Me.Hide
    fFindModem.Show
End Sub

Private Sub bFindModem2_Click()
    Me.Hide
    FindModem.Show

End Sub

Private Sub bGetCurrentData_Click()
    Dim Stringa As String
    MSComm1.Output = "T" + vbCrLf
    Stringa = InputComTimeOut(10)
    If Stringa <> "TimeOut" Then
        Text1.Text = "Actual date in GMM" & vbCrLf
        Text1.Text = Text1.Text & Stringa
        
    Else
        MsgBox "Error in getting Time from GMM! No answer!", vbCritical
        
    End If
    
End Sub

Private Sub bGetTodayFiles_Click()
    Dim Stringa As String
    Dim Yesterday As String
    MSComm1.InBufferCount = 0
    Text1.Text = Text1.Text + "Sending DT" + vbCrLf
    MSComm1.Output = "DT" + vbCr
    Stringa = InputComTimeOut(5)
    Stringa = InputComTimeOut(5)
    If Stringa = "TimeOut" Then

    End If
    Text1.Text = Text1.Text + "01 " + Stringa + vbCrLf
    
    Yesterday = Format(Now - 1, "yymmdd")
    MSComm1.Output = Yesterday + vbCr
    Text1.Text = Text1.Text + Yesterday + vbCrLf
    Stringa = InputComTimeOut(5)
    Text1.Text = Text1.Text + "02 " + Stringa + vbCrLf
    Stringa = InputComTimeOut(5)
    Text1.Text = Text1.Text + "03 " + Stringa + vbCrLf
    MSComm1.InBufferCount = 0
    
    Me.MousePointer = vbHourglass
    YmodemRx.Show
    YmodemRx.YModemDownload

    Text1.Text = Text1.Text + "File received!" + vbCrLf



    Stringa = InputComTimeOut(5) 'Get blank line
    Stringa = InputComTimeOut(5) 'Get Transmission Successful. message
    Text1.Text = Stringa + vbCrLf
    Stringa = InputComTimeOut(5) 'Get blank line
    Stringa = InputComTimeOut(5) 'Get GMM> line
    Text1.Text = Text1.Text + "01 " + Stringa + vbCrLf
    

    MSComm1.InBufferCount = 0

    Text1.Text = Text1.Text + "Sending DM" + vbCrLf
    MSComm1.Output = "DM" + vbCr
    Stringa = InputComTimeOut(5)
    Stringa = InputComTimeOut(5)
    If Stringa = "TimeOut" Then

    End If
    Text1.Text = Text1.Text + "01 " + Stringa + vbCrLf
    
    Yesterday = Format(Now - 1, "yymmdd")
    MSComm1.Output = Yesterday + vbCr
    Text1.Text = Text1.Text + Yesterday + vbCrLf
    Stringa = InputComTimeOut(5)
    Text1.Text = Text1.Text + "02 " + Stringa + vbCrLf
    Stringa = InputComTimeOut(5)
    Text1.Text = Text1.Text + "03 " + Stringa + vbCrLf
    MSComm1.InBufferCount = 0

    Me.MousePointer = vbHourglass
    YmodemRx.Show
    YmodemRx.YModemDownload

    Stringa = InputComTimeOut(5) 'Get blank line
    Stringa = InputComTimeOut(5) 'Get Transmission Successful. message
    Text1.Text = Stringa + vbCrLf
    Stringa = InputComTimeOut(5) 'Get blank line
    Stringa = InputComTimeOut(5) 'Get GMM> line
    Text1.Text = Text1.Text + "01 " + Stringa + vbCrLf
    
    Me.MousePointer = vbNormal

End Sub

Private Sub bModem_Click()
    'apre porta
    'Controlla che ci sia un modem
    'chiama la rubrica
    'chiama modem remoto
    Me.Hide
    fModem.Show

End Sub
Private Sub bConnect_Click()
    Dim Stringa As String
    Dim Retry As Byte
    Retry = 1
WakeUp:
    'Wake up
    MSComm1.Output = "W"
    Text1.Text = Text1.Text + "Connecting GMM " + Stringa + vbCrLf
    Stringa = InputComTimeOut(5) 'First line is a blank
    Stringa = InputComTimeOut(5) 'Here is the answer
    If Stringa = "TimeOut" Then
        'GMM in not answering retry?
        Retry = Retry + 1
        If Retry >= 6 Then
            Text1.Text = Text1.Text + "No answer from GMM aborting" + vbCrLf
        Else
            Text1.Text = Text1.Text + "No answer retrying " + Str(Retry) + vbCrLf
            GoTo WakeUp
        End If
    End If
    If Left(Stringa, 3) <> "GMM" Then
        'incorrect answer!!!
        Text1.Text = Text1.Text + "Incorrect answer from GMM " + Stringa + vbCrLf
        Exit Sub
    Else
        Text1.Text = Text1.Text + "GMM connected " + Stringa + vbCrLf
        lStatus.Caption = "Connected"
        bGetTodayFiles.Enabled = True
    End If
    MSComm1.InBufferCount = 0
End Sub

Private Sub bRymodem_Click()
    Dim Stringa As String
    Stringa = sGetAppPath
    MSComm1.CommPort = 6
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.Handshaking = comNone
    MSComm1.Handshaking = comRTS
    MSComm1.PortOpen = True
    Port = 6
    Me.MousePointer = vbHourglass
    YmodemRx.Show
    YmodemRx.YModemDownload
    'Stringa = YmodemReceiveFile(sGetAppPath)
    MSComm1.PortOpen = False
End Sub

Private Sub bSetDate_Click()
    Dim Stringa As String
    'Prende una data o la data di sistema
    Debug.Print "Start of set Date & Time"
    MSComm1.Output = "X" + vbCrLf
    Debug.Print "X"
    Stringa = InputComTimeOut(10)
    Debug.Print Stringa
    If Stringa <> "Input PicoDOS command to eXecute :" Then
        'ERROR!
    End If
    MSComm1.Output = "DATE gg/mm/aa hh:mm:ss" + vbCrLf
    Debug.Print "DATE gg/mm/aa hh:mm:ss"
    Stringa = InputComTimeOut(10)
    If Left$(Stringa, 12) <> "Clock reads:" Then
        Debug.Print Left$(Stringa, 12)
        'ERROR!
    Else
        Text1.Text = Stringa
    End If
End Sub

Private Sub bTerminal_Click()
    Me.Hide
    frmTerminal.Show
End Sub

Private Sub bTranslate_Click()
    Dim Linea As String     'Variabile dove registro ogni linea di dati ricevuta
    Dim MioFile As String
    Dim Dummy As String
    Dim FileIn As String
    Dim Blocco() As Byte        'Blocco dati temporaneo in bytes
    Dim Buffer As Byte       'buffer temporaneo per i dati
    Dim BloccoDati() As Byte    'Blocco dati
    Dim iBloccoDati As Long     'Indice all'interno di BloccoDati()
    Dim DFPNT As Long           'Numero di bytes da scaricare
    Dim Bytes As Long           'Numero bytes scaricati
    Dim LungCounter As Long
    Dim Barra As Double
    Dim IncBarra As Double 'Incremento barra contatore per ogni riga
    Dim TimeOuts As Long        'Contatore dei Time Out
    Dim iDumm As Long
    Dim Intero As Integer
    Dim Lungo As Long
    Dim Stringa As String
    Dim lStringa As Long
    Dim Float As Single
    Dim iBlocco As Long
    Dim i As Long
    Dim j As Long
    Dim Filnb As Long
    
    Dim DatePrefix As String        'Year Month day
    
    Dim Hours As Integer
    
    Dim NBS As Byte 'Number of Binary data Series
    Dim MethaneBlock(51) As Byte
    Dim H2Sblock(51) As Byte
    Dim CTDblock(111) As Byte
    Dim IDsMet As Byte
    Dim SmsMet As Integer
    Dim NbmMet As Byte
    Dim DateMinsMet As Long
    Dim ChansTrans As Byte
    Dim WeekDayn As Byte
    Dim MeanCH4 As Long
    Dim SigmaCH4 As Long
    
    Dim IDsH2S As Byte
    Dim SmsH2S As Integer
    Dim NbmH2S As Byte
    Dim DateMinsH2S As Long
    Dim MeanH2S As Long
    Dim SigmaH2S As Long
    
    Dim IDsCTD As Byte
    Dim SmsCTD As Integer
    Dim NbmCTD As Byte
    Dim DateMinsCTD As Long
    Dim HeaderCTD As String
    Dim FlagsCTD As String
    Dim TempCTD As Long
    Dim TempCTDr As Double
    Dim CondCTD As Long
    Dim CondCTDr As Double
    Dim PressCTD As Long
    Dim pressCTDr As Double
    
    
    
    Dim DatafieldCH4(6) As DateRdCH4
    
    Dim TuttiDati(150) As DataRecord
    Dim TuttiDatiIndex As Long

    'impostazioni iniziali di CmDialog1 per la scelta del file da leggere
    NewPath sGetAppPath
    On Error GoTo Annulla
    CmDialog1.InitDir = sGetAppPath
    CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    CmDialog1.Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    CmDialog1.Filter = "File PMM (*.pmm)|*.pmm|Tutti i file (*.*)|*.*"
    CmDialog1.ShowOpen
    FileIn = CmDialog1.filename
    
    Stringa = GetNameFromDir(FileIn)
    FileIn = CmDialog1.filename 'Non ho capito perchè ma dopo GetNameFromDir FileIn perde il percorso!!!
    DatePrefix = "20" & Left$(Stringa, Len(Stringa) - 4)
    
    'impostazioni iniziali di CmDialog1 per la scelta del file da scrivere
    NewPath sGetAppPath
    On Error GoTo Annulla
    CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    CmDialog1.Filter = "File Ascii (*.csv)|*.csv|Tutti i file (*.*)|*.*"
    CmDialog1.filename = Left$(FileIn, Len(FileIn) - 3) + "csv"
    CmDialog1.ShowSave
    FileOut = CmDialog1.filename
    
    
    On Error GoTo 0
    'Start Translation
    DoEvents
    
    Me.MousePointer = vbHourglass



    DFPNT = FileLen(FileIn)
    'DFPNT dovrebbe essere sempre positivo ma non si sa mai!
    If DFPNT < 0 Then DFPNT = 0



    If DFPNT = 0 Then
        MsgBox ("Non ci sono dati nel file!")
        'Qui eventualmente si puo' mettere una
        'routine di scarico dati di emergenza
        'Esci
        Exit Sub
    End If

    TuttiDatiIndex = 1
    BloccoDati = ""
    TimeOuts = 0
    Bytes = 0
    Intero = 0
    'dati = 0
    iBloccoDati = 0
    ReDim BloccoDati(DFPNT + 100)
    
    Filnb = FreeFile
    Open FileIn For Binary As #1    'Open the PMM file
    ReDim BloccoDati(LOF(1) - 1)    'Redim the array Bloccodati to the PMM dimension (5136 bytes)
    Get #1, , BloccoDati            'Read the entire file in one shot
    Close #1                        'Close the PMM file

    Filnb = FreeFile
    Open FileOut For Output As #1    'Open the dat file
    Print #Filnb, "File:"; FileIn
    
'    If iBloccoDati < DFPNT Then
'        Messaggio = "Errore letti" + Str(iBloccoDati) + " dati invece di" + Str(DFPNT)
'        MsgBox (Messaggio)
'        Esci
'    End If

    NBS = BloccoDati(0) 'Get the first byte NBS Number of Binary data Series (of sensors). For GMM is always 3
    'Debug.Print "NBS="; NBS;
    'Check NBS
    If NBS = 3 Then
        'Debug.Print " OK!!!"
      Else
        Debug.Print " ERROR!!!!!!"
        MsgBox "Error in NBS", vbCritical, "ERROR"
    End If
    
    On Error GoTo Printall
    
    For Hours = 0 To 23  'For every hour record
        On Error GoTo Printall
        'Stringa = bMID(BloccoDati, 1, 51)
        For i = 1 To 51                     'Transer data to the Methane block
            MethaneBlock(i) = BloccoDati(i + 214 * Hours)
            DoEvents
        Next
        For i = 52 To 102                   'Transer data to the H2S block
            H2Sblock(i - 51) = BloccoDati(i + 214 * Hours)
            DoEvents
        Next
        For i = 103 To 213                   'Transer data to the CTD block
            CTDblock(i - 102) = BloccoDati(i + 214 * Hours)
            DoEvents
        Next
    
        On Error GoTo 0

        'Debug.Print "Start of Methane Block"
        Print #Filnb, "Start of Methane Block"
        Buffer = MethaneBlock(1)
        'Debug.Print "Buffer="; Buffer
        IDsMet = (Buffer And 240) / 16 'get the left 4 bits and shift them to the right
        'Debug.Print "IDsMet="; IDsMet 'This is the sensor type. Must be 3 or 4 or 5
        'Debug.Print "sensor= "; SensorType(IDsMet)
        Print #Filnb, "Sensor= "; SensorType(IDsMet)
        SmsMet = CLng(MethaneBlock(1)) * 256 + MethaneBlock(2) 'get the two bytes
        SmsMet = SmsMet And 4095  'erase the left 4 bits. Must be 5
        'Debug.Print "SmsMet="; SmsMet 'This is the size of one measurement
        NbmMet = MethaneBlock(3)       'Number of dated measurements
        'Debug.Print "NbmMet="; NbmMet
        For i = 0 To 5
            TuttiDati(TuttiDatiIndex + i).Datagiorno = DatePrefix
            DateMinsMet = CLng(MethaneBlock(4 + i * 8)) * 256 + MethaneBlock(5 + i * 8)
            WeekDayn = Int(DateMinsMet / 1440) + 1 'Number of minutes since last Monday 00:00
            Intero = DateMinsMet - 1440 * (WeekDayn - 1) 'Number of minutes since midnight
            Lungo = Val(DatePrefix)
            TuttiDati(TuttiDatiIndex + i).DataMeas = CDate(DateSerial(Left(DatePrefix, 4), Mid(DatePrefix, 5, 2), Right(DatePrefix, 2)) + Intero / 1440)
            TuttiDati(TuttiDatiIndex + i).Oraminuti = DateMinsMet
            'Debug.Print "WeekDayn="; WeekDayn; " "
            'Debug.Print "DateMinsMet="; DateMinsMet; " ";
            ChansTrans = MethaneBlock(6 + i * 8)
            'Debug.Print "ChansTrans="; ChansTrans; " ";
            Dummy = MethaneBlock(7 + i)
            'Debug.Print "dummy="; Dummy
            MeanCH4 = CLng(MethaneBlock(8 + i * 8)) * 256 + MethaneBlock(9 + i * 8)
            SigmaCH4 = CLng(MethaneBlock(10 + i * 8)) * 256 + MethaneBlock(11 + i * 8)
            'Debug.Print "MeanCH4="; MeanCH4; " SigmaCH4="; SigmaCH4
            TuttiDati(TuttiDatiIndex + i).MeanCH4 = MeanCH4
            TuttiDati(TuttiDatiIndex + i).SigmaCH4 = SigmaCH4
            'Print #Filnb, DatePrefix; " "; DateMinsMet; " "; MeanCH4; " "; SigmaCH4
            
        Next i
        
        
        'Debug.Print "Start of H2S Block"
        Print #Filnb, "Start of H2S Block"
        Buffer = H2Sblock(1)
        'Debug.Print "Buffer="; Buffer
        IDsH2S = (Buffer And 240) / 16 'get the left 4 bits and shift them to the right
        'Debug.Print "IDsH2S="; IDsH2S 'This is the sensor type. Must be 6
        'Debug.Print "sensor= "; SensorType(IDsH2S)
        Print #Filnb, "Sensor= "; SensorType(IDsH2S)
        SmsH2S = CLng(H2Sblock(1)) * 256 + H2Sblock(2) 'get the two bytes
        SmsH2S = SmsH2S And 4095  'erase the left 4 bits. Must be 5
        'Debug.Print "SmsH2S="; SmsH2S 'This is the size of one measurement
        NbmH2S = H2Sblock(3)       'Number of dated measurements
        'Debug.Print "NbmH2S="; NbmH2S
        For i = 0 To 5
            DateMinsH2S = CLng(H2Sblock(4 + i * 8)) * 256 + H2Sblock(5 + i * 8)
            WeekDayn = DateMinsH2S / 1440 + 1 'Number of minutes since last Monday 00:00
            'Debug.Print "WeekDayn="; WeekDayn; " "
            'Debug.Print "DateMinsH2S="; DateMinsH2S; " ";
            ChansTrans = H2Sblock(6 + i * 8)
            'Debug.Print "ChansTrans="; ChansTrans; " ";
            Dummy = H2Sblock(7 + i * 8)
            'Debug.Print "dummy="; Dummy
            MeanH2S = CLng(H2Sblock(8 + i * 8)) * 256 + H2Sblock(9 + i * 8)
            SigmaH2S = CLng(H2Sblock(10 + i * 8)) * 256 + H2Sblock(11 + i * 8)
            'Debug.Print "MeanH2S="; MeanH2S; " SigmaH2S="; SigmaH2S
            TuttiDati(TuttiDatiIndex + i).MeanH2S = MeanH2S
            TuttiDati(TuttiDatiIndex + i).SigmaH2S = SigmaH2S
            'Print #Filnb, DatePrefix; " "; DateMinsMet; " "; MeanH2S; " "; SigmaH2S

        Next i
        
        'Debug.Print "Start of CTD Block"
        Print #Filnb, "Start of CTD Block"
        Buffer = CTDblock(1)
        'Debug.Print "Buffer="; Buffer
        IDsCTD = (Buffer And 240) / 16 'get the left 4 bits and shift them to the right
        'Debug.Print "IDsCTD="; IDsCTD 'This is the sensor type. Must be 7
        'Debug.Print "sensor= "; SensorType(IDsCTD)
        Print #Filnb, "Sensor= "; SensorType(IDsCTD)
        SmsCTD = CLng(CTDblock(1)) * 256 + CTDblock(2) 'get the two bytes
        SmsCTD = SmsCTD And 4095  'erase the left 4 bits. Must be 15
        'Debug.Print "SmsCTD="; SmsCTD 'This is the size of one measurement
        NbmCTD = CTDblock(3)       'Number of dated measurements
        'Debug.Print "NbmCTD="; NbmCTD
        For i = 0 To 5
            DateMinsCTD = CLng(CTDblock(4 + i * 18)) * 256 + CTDblock(5 + i * 18)
            WeekDayn = DateMinsCTD / 1440 + 1 'Number of minutes since last Monday 00:00
            Intero = DateMinsCTD - 1440 * (WeekDayn - 1) 'Number of minutes since midnight
            'Debug.Print "WeekDayn="; WeekDayn; " "
            'Debug.Print "DateMinsCTD="; DateMinsCTD; " ";
            ChansTrans = CTDblock(6 + i * 18)
            'Debug.Print "ChansTrans="; ChansTrans; " ";
            Dummy = CTDblock(7 + i * 18)
            'Debug.Print "dummy="; Dummy
    
            HeaderCTD = CTDblock(8 + i * 18)  'Must be C
            'Debug.Print "Header="; HeaderCTD
            FlagsCTD = CTDblock(9 + i * 18) 'E=error 0=empty 1=ok
            'Debug.Print "Flags="; FlagsCTD
            Print #Filnb, DatePrefix; " "; DateMinsCTD; " ";
            Select Case FlagsCTD
                Case 49
                    Print #Filnb, "Flags=1 OK!"
                Case 48
                    Print #Filnb, "Flags=0 Empty! "; Format(CDate(DateSerial(Left(DatePrefix, 4), Mid(DatePrefix, 5, 2), Right(DatePrefix, 2)) + Intero / 1440), "yyyy/mm/dd hh:mm")
                Case 69
                    Print #Filnb, "Flags=E ERROR! "; Format(CDate(DateSerial(Left(DatePrefix, 4), Mid(DatePrefix, 5, 2), Right(DatePrefix, 2)) + Intero / 1440), "yyyy/mm/dd hh:mm")
                Case Else
                    Print #Filnb, "Flags=undefined! "; Format(CDate(DateSerial(Left(DatePrefix, 4), Mid(DatePrefix, 5, 2), Right(DatePrefix, 2)) + Intero / 1440), "yyyy/mm/dd hh:mm")
            End Select
           
            Stringa = bMID(CTDblock, (10 + i * 18), 4)
            TempCTD = String2long(Stringa)
            Stringa = bMID(CTDblock, (14 + i * 18), 4)
            CondCTD = String2long(Stringa)
            Stringa = bMID(CTDblock, (18 + i * 18), 4)
            PressCTD = String2long(Stringa)
            
            'Debug.Print "T="; TempCTD; "C="; CondCTD; "P="; PressCTD
            TuttiDati(TuttiDatiIndex + i).Temp = TempCTD / 10000
            TuttiDati(TuttiDatiIndex + i).Cond = CondCTD / 100000
            TuttiDati(TuttiDatiIndex + i).Press = PressCTD / 1000
            'Print #Filnb, DatePrefix; " "; DateMinsCTD; " "; TempCTD; " "; CondCTD; " "; PressCTD
           
            
        Next i
        TuttiDatiIndex = TuttiDatiIndex + 6
     Next Hours

Printall:
    On Error GoTo 0
    'Print all data
    Print #Filnb, "All Data"
    Print #Filnb, "Date; MeanCH4; SigmaCH4; MeanH2S; SigmaH2S; Temp; Cond; Press"
    For i = 1 To 144
        Print #Filnb, Format(TuttiDati(i).DataMeas, "yyyy/mm/dd hh:mm"); " ;";
        'Print #Filnb, TuttiDati(i).Datagiorno; " ;";
        'Print #Filnb, TuttiDati(i).Oraminuti; " ;";
        Print #Filnb, TuttiDati(i).MeanCH4; " ;";
        Print #Filnb, TuttiDati(i).SigmaCH4; " ;";
        Print #Filnb, TuttiDati(i).MeanH2S; " ;";
        Print #Filnb, TuttiDati(i).SigmaH2S; " ;";
        Print #Filnb, TuttiDati(i).Temp; " ;";
        Print #Filnb, TuttiDati(i).Cond; " ;";
        Print #Filnb, TuttiDati(i).Press
    Next i
Annulla:
    Me.MousePointer = vbDefault
    Close #Filnb
    DoEvents

End Sub

Private Sub bTranslatePTM_Click()
    
    Dim Linea As String     'Variabile dove registro ogni linea di dati ricevuta
    Dim MioFile As String
    Dim Dummy As String
    Dim FileIn As String
    Dim Blocco() As Byte        'Blocco dati temporaneo in bytes
    Dim Buffer As Byte       'buffer temporaneo per i dati
    Dim BloccoDati() As Byte    'Blocco dati
    Dim iBloccoDati As Long     'Indice all'interno di BloccoDati()
    Dim DFPNT As Long           'Numero di bytes da scaricare
    Dim Bytes As Long           'Numero bytes scaricati
    Dim LungCounter As Long
    Dim Barra As Double
    Dim IncBarra As Double 'Incremento barra contatore per ogni riga
    Dim TimeOuts As Long        'Contatore dei Time Out
    Dim iDumm As Long
    Dim Intero As Integer
    Dim Lungo As Long
    Dim Stringa As String
    Dim lStringa As Long
    Dim Float As Single
    Dim iBlocco As Long
    Dim i As Long
    Dim j As Long
    Dim Filnb As Long
    
    Dim DateDay As Long
    Dim DateMins As Long
    Dim TuttiDati(24) As DataRecordTEC
    Dim TuttiDatiIndex As Long

    
    Dim DatePrefix As String        'Year Month day
    
    Dim Hours As Integer
    Dim SMS As Byte  'Must be 24 $18
    Dim NBM As Byte  'Must be 1
    Dim Data As Long    'Minutes since Monday 00:00
    Dim Header As String    'Must be "T"
    Dim Reboots As Byte
    Dim UsedMemory As Double
    Dim FreeMEmory As Double
    Dim Header2 As String   'Must be "S"
    Dim Flags As Byte
    Dim BatteryVoltage As Long   'in mV
    Dim BatteryCurrent As Long   'in mA
    Dim DACStemperature As Long 'in 0.01 °C
    Dim BatteryVesselTemperature As Long
    Dim DACSvesselPressure As Long  'in mbar
    Dim BatteryVesselPressure As Long
    
    Dim WeekDayn As Byte
    
    
    'impostazioni iniziali di CmDialog1 per la scelta del file da leggere
    NewPath sGetAppPath
    On Error GoTo Annulla
    CmDialog1.InitDir = sGetAppPath
    CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    CmDialog1.Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    CmDialog1.Filter = "File PTM (*.ptm)|*.ptm|Tutti i file (*.*)|*.*"
    CmDialog1.ShowOpen
    FileIn = CmDialog1.filename


    Stringa = GetNameFromDir(FileIn)
    FileIn = CmDialog1.filename
    DatePrefix = "20" & Left$(Stringa, Len(Stringa) - 4)
    
    'impostazioni iniziali di CmDialog1 per la scelta del file da scrivere
    NewPath sGetAppPath
    On Error GoTo Annulla
    CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    CmDialog1.Filter = "File Ascii (*.csv)|*.csv|Tutti i file (*.*)|*.*"
    CmDialog1.filename = Left$(FileIn, Len(FileIn) - 3) + "tec.csv"
    CmDialog1.ShowSave
    FileOut = CmDialog1.filename
    
    On Error GoTo 0
    'Start Translation
    DoEvents
    
    Me.MousePointer = vbHourglass



    DFPNT = FileLen(FileIn)
    'DFPNT dovrebbe essere sempre positivo ma non si sa mai!
    If DFPNT < 0 Then DFPNT = 0



    If DFPNT = 0 Then
        MsgBox ("Non ci sono dati nel file!")
        'Qui eventualmente si puo' mettere una
        'routine di scarico dati di emergenza
        'Esci
        Exit Sub
    End If

    TuttiDatiIndex = 1
    BloccoDati = ""
    TimeOuts = 0
    Bytes = 0
    Intero = 0
    'dati = 0
    iBloccoDati = 0
    ReDim BloccoDati(DFPNT + 100)
    
    Filnb = FreeFile
    Open FileIn For Binary As #1    'Open the PMM file
    ReDim BloccoDati(LOF(1) - 1)    'Redim the array Bloccodati to the PMM dimension (5136 bytes)
    Get #1, , BloccoDati            'Read the entire file in one shot
    Close #1                        'Close the PMM file

    Filnb = FreeFile
    Open FileOut For Output As #1    'Open the dat file
    Print #Filnb, "File:"; FileIn
    
    Print #Filnb, "Start of PTM"
    
    On Error GoTo Printall

    For Hours = 0 To 23  'For every hour record
    
        'On Error GoTo 0

        
        
        'SMS 0
        Buffer = BloccoDati(0 + Hours * 28)
        SMS = Buffer
        'Debug.Print "Buffer="; Buffer
        'Debug.Print "SMS="; SMS
        If SMS <> 24 Then
            Debug.Print "SMS error on hour "; Hours; " "; SMS; " different from 24"
        End If
        
        'NBM 1
        Buffer = BloccoDati(1 + Hours * 28)
        NBM = Buffer
        'Debug.Print "Buffer="; Buffer
        'Debug.Print "NBM="; NBM
        If NBM <> 1 Then
            Debug.Print "NBM error on hour "; Hours; NBM; " different from 1"
        End If
        
        'Date 2-3
        TuttiDati(Hours).Datagiorno = DatePrefix
        DateMins = CLng(BloccoDati(2 + Hours * 28)) * 256 + BloccoDati(3 + Hours * 28)
        WeekDayn = Int(DateMins / 1440) + 1 'Number of minutes since last Monday 00:00
        DateDay = Val(DatePrefix)
        Intero = DateMins - 1440 * (WeekDayn - 1) 'Number of minutes since midnight
        TuttiDati(Hours).DataMeas = CDate(DateSerial(Left(DatePrefix, 4), Mid(DatePrefix, 5, 2), Right(DatePrefix, 2)) + Intero / 1440)
        If Hour(TuttiDati(Hours).DataMeas) = 0 Then
            TuttiDati(Hours).DataMeas = TuttiDati(Hours).DataMeas + 1
        End If
        'Control prints
'        Print #Filnb, Left(DatePrefix, 4); " ";
'        Print #Filnb, Mid(DatePrefix, 5, 2); " ";
'        Print #Filnb, Right(DatePrefix, 2); " ";
'        Print #Filnb, DateMins; " ";
'        Print #Filnb, Intero / 1440; " ";
'        Print #Filnb, WeekDayn; " ";
'        Print #Filnb, CDate(DateSerial(Left(DatePrefix, 4), Mid(DatePrefix, 5, 2), Right(DatePrefix, 2)) + Intero / 1440)
        'Print #Filnb,
        'End of control prints
        TuttiDati(Hours).Oraminuti = DateMins
            
        'Header 4
        Buffer = BloccoDati(4 + Hours * 28)
        Header = Chr(Buffer)
        'Debug.Print "Buffer="; Buffer
        'Debug.Print "Header="; Header
        If Header <> "T" Then
            Debug.Print "Header error on hour "; Hours; Header; " different from T"
        End If
            
        'reboots 5
        Buffer = BloccoDati(5 + Hours * 28)
        Reboots = Buffer
        'Debug.Print "Buffer="; Buffer
        'Debug.Print "Reboots="; Reboots
        TuttiDati(Hours).Reboots = Reboots
            
        'Used Memory 6-7-8-9
        UsedMemory = CDbl(BloccoDati(6 + Hours * 28)) * 16777215 + CDbl(BloccoDati(7 + Hours * 28)) _
        * 65535 + CDbl(BloccoDati(8 + Hours * 28)) * 256 + BloccoDati(9 + Hours * 28)
        TuttiDati(Hours).UsedMem = UsedMemory
        
        'Free Memory 10-11-12-13
        FreeMEmory = CDbl(BloccoDati(10 + Hours * 28)) * 16777215 + CDbl(BloccoDati(11 + Hours * 28)) _
        * 65535 + CDbl(BloccoDati(12 + Hours * 28)) * 256 + BloccoDati(13 + Hours * 28)
        TuttiDati(Hours).FreeMem = FreeMEmory

        'Header2 14
        Buffer = BloccoDati(14 + Hours * 28)
        Header2 = Chr(Buffer)
        'Debug.Print "Buffer="; Buffer
        'Debug.Print "Header2="; Header2
        If Header2 <> "S" Then
            Debug.Print "Header2 error on hour "; Hours; Header2; " different from S"
        End If
            
        'Flags 15
        Buffer = BloccoDati(15 + Hours * 28)
        Flags = Buffer
        'Debug.Print "Buffer="; Buffer
        'Debug.Print "Flags="; Flags
        TuttiDati(Hours).Flags = Flags
        If Flags <> 0 Then
            'Some sensors out of range!
        End If
    
        'Battery Voltage 16-17
        Lungo = BloccoDati(16 + Hours * 28) * 256 + BloccoDati(17 + Hours * 28)
        BatteryVoltage = Lungo
        'Debug.Print "Lungo="; Lungo
        'Debug.Print "BatteryVoltage="; BatteryVoltage / 1000
        TuttiDati(Hours).BattVol = CSng(BatteryVoltage) / 1000
'        Print #Filnb, BatteryVoltage; " ";
'        Print #Filnb, BatteryVoltage / 1000; " ";
'        Print #Filnb, CSng(BatteryVoltage); " ";
'        Print #Filnb, CSng(BatteryVoltage) / 1000; " ";
'        Print #Filnb, TuttiDati(Hours).BattVol
        If BatteryVoltage < 10000 Then
            'Battery out!!!!!
        End If

        'Battery Current 18-19
        Lungo = BloccoDati(18 + Hours * 28) * 256 + BloccoDati(19 + Hours * 28)
        BatteryCurrent = Lungo
        'Debug.Print "Lungo="; Lungo
        'Debug.Print "BatteryCurrent="; BatteryCurrent
        TuttiDati(Hours).BattCurr = CSng(BatteryCurrent)
        'Print #Filnb, BatteryCurrent; " ";
        'Print #Filnb, CSng(BatteryCurrent); " ";
        'Print #Filnb, TuttiDati(Hours).BattCurr
        If BatteryCurrent < 10000 Then
            'Battery out!!!!!
        End If
     
        'DACS Temperature 20-21
        Lungo = BloccoDati(20 + Hours * 28) * 256 + BloccoDati(21 + Hours * 28)
        DACStemperature = Lungo
        'Debug.Print "Lungo="; Lungo
        'Debug.Print "DACStemperature="; DACStemperature / 100
        TuttiDati(Hours).DacsT = CSng(DACStemperature) / 100
        If DACStemperature > 18 Then
            'Error!!!!!
        End If
     
        'Battery Temperature 22-23
        Lungo = BloccoDati(22 + Hours * 28) * 256 + BloccoDati(23 + Hours * 28)
        BatteryVesselTemperature = Lungo
        'Debug.Print "Lungo="; Lungo
        'Debug.Print "BatteryVesselTemperature="; BatteryVesselTemperature / 100
        TuttiDati(Hours).BattT = CSng(BatteryVesselTemperature) / 100
        If BatteryVesselTemperature > 18 Then
            'Error!!!!!
        End If
     
     
        'DACS Pressure 24-25
        Lungo = BloccoDati(24 + Hours * 28) * 256 + BloccoDati(25 + Hours * 28)
        DACSvesselPressure = Lungo
        'Debug.Print "Lungo="; Lungo
        'Debug.Print "DACSvesselPressure="; DACSvesselPressure
        TuttiDati(Hours).DacsP = CSng(DACSvesselPressure)
        If DACSvesselPressure > 18 Then
            'Error!!!!!
        End If
     
        'Battery Pressure 26-27
        Lungo = BloccoDati(26 + Hours * 28) * 256 + BloccoDati(27 + Hours * 28)
        BatteryVesselPressure = Lungo
        'Debug.Print "Lungo="; Lungo
        'Debug.Print "BatteryVesselPressure="; BatteryVesselPressure
        TuttiDati(Hours).BattP = CSng(BatteryVesselPressure)
        If BatteryVesselPressure > 18 Then
            'Error!!!!!
        End If
     
     
     Next Hours

Printall:
    On Error GoTo 0
    'Print all data
    Print #Filnb, "All Data"
    Print #Filnb, "Date; Reboots; UsedMem; FreeMem; Flags; BattV; BattC; DACSt; BattT; DacsP; BattP"
    For i = 0 To 23
        Print #Filnb, Format(TuttiDati(i).DataMeas, "yyyy/mm/dd hh:mm"); " ;";
        'Print #Filnb, TuttiDati(i).Datagiorno; " ;";
        'Print #Filnb, TuttiDati(i).Oraminuti; " ;";
        Print #Filnb, TuttiDati(i).Reboots; " ;";
        Print #Filnb, TuttiDati(i).UsedMem; " ;";
        Print #Filnb, TuttiDati(i).FreeMem; " ;";
        Print #Filnb, TuttiDati(i).Flags; " ;";
        Print #Filnb, TuttiDati(i).BattVol; " ;";
        Print #Filnb, TuttiDati(i).BattCurr; " ;";
        Print #Filnb, TuttiDati(i).DacsT; " ;";
        Print #Filnb, TuttiDati(i).BattT; " ;";
        Print #Filnb, TuttiDati(i).DacsP; " ;";
        Print #Filnb, TuttiDati(i).BattP
    Next i
Annulla:
    Me.MousePointer = vbDefault
    Close #Filnb
    DoEvents

End Sub

