VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form fMain 
   Caption         =   "Pegaso I"
   ClientHeight    =   6615
   ClientLeft      =   3405
   ClientTop       =   2340
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bRymodem 
      Caption         =   "Receive Ymodem"
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   5880
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   9120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton bTranslate 
      Caption         =   "Translate Data"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton bTerminal 
      Caption         =   "&Open Terminal"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   375
      Left            =   8520
      TabIndex        =   9
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton bFindModem 
      Caption         =   "&Find Modem"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton bModem 
      Caption         =   "&Modem Call"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton bSetDate 
      Caption         =   "&Set Date&&Time"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton bGetCurrentData 
      Caption         =   "&Get current Date&&Time"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   4575
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label2 
         Caption         =   "Data Files"
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
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton bConnect 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "PEGASO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
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

Private Sub bEnd_Click()
    Unload Me
    End
End Sub

Private Sub bFindModem_Click()
    Me.Hide
    fFindModem.Show
End Sub

Private Sub bModem_Click()
    'apre porta
    'Controlla che ci sia un modem
    'chiama la rubrica
    'chiama modem remoto
End Sub
Private Sub bConnect_Click()
    'Wake up
End Sub

Private Sub bRymodem_Click()
    Dim stringa As String
    'stringa = sGetAppPath
    MSComm1.PortOpen = True
    stringa = YmodemReceiveFile(sGetAppPath)
    MSComm1.PortOpen = False
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
    Dim stringa As String
    Dim lStringa As Long
    Dim Float As Single
    Dim iBlocco As Long
    Dim i As Long
    Dim J As Long
    Dim Filnb As Long
    
    Dim DatePrefix As String        'Year Month day
    
    Dim Hour As Integer
    
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
    FileIn = CmDialog1.Filename
    
    stringa = GetNameFromDir(FileIn)
    DatePrefix = "20" & Left$(stringa, Len(stringa) - 4)
    
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
    'CmDialog1.Filter = "Dati ASCII(*.dat)|*.dat|Dati Sima (*.bin)|*.bin|Tutti i file (*.*)|*.*"
    CmDialog1.Filter = "File Ascii (*.dat)|*.dat|Tutti i file (*.*)|*.*"
    CmDialog1.Filename = Left$(FileIn, Len(FileIn) - 3) + "dat"
    CmDialog1.ShowSave
    FileOut = CmDialog1.Filename
    
    
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
    
    For Hour = 0 To 23  'For every hour record
        On Error GoTo Printall
        'Stringa = bMID(BloccoDati, 1, 51)
        For i = 1 To 51                     'Transer data to the Methane block
            MethaneBlock(i) = BloccoDati(i + 214 * Hour)
            DoEvents
        Next
        For i = 52 To 102                   'Transer data to the H2S block
            H2Sblock(i - 51) = BloccoDati(i + 214 * Hour)
            DoEvents
        Next
        For i = 103 To 213                   'Transer data to the CTD block
            CTDblock(i - 102) = BloccoDati(i + 214 * Hour)
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
            WeekDayn = DateMinsMet / 1440 + 1 'Number of minutes since last Monday 00:00
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
           
            stringa = bMID(CTDblock, (10 + 1 * 18), 4)
            TempCTD = String2long(stringa)
            stringa = bMID(CTDblock, (14 + 1 * 18), 4)
            CondCTD = String2long(stringa)
            stringa = bMID(CTDblock, (18 + 1 * 18), 4)
            PressCTD = String2long(stringa)
            
            'Debug.Print "T="; TempCTD; "C="; CondCTD; "P="; PressCTD
            TuttiDati(TuttiDatiIndex + i).Temp = TempCTD / 10000
            TuttiDati(TuttiDatiIndex + i).Cond = CondCTD / 100000
            TuttiDati(TuttiDatiIndex + i).Press = PressCTD / 1000
            'Print #Filnb, DatePrefix; " "; DateMinsCTD; " "; TempCTD; " "; CondCTD; " "; PressCTD
           
            
        Next i
        TuttiDatiIndex = TuttiDatiIndex + 6
     Next Hour

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
