VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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

Private Sub bTerminal_Click()
    Me.Hide
    frmTerminal.Show
End Sub

Private Sub bTranslate_Click()
    Dim Linea As String     'Variabile dove registro ogni linea di dati ricevuta
    Dim MioFile As String
    Dim dummy As String
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
    
    Dim NBS As Byte 'Number of Binary data Series
    Dim MethaneBlock(71)
    Dim H2Sblock(71)
    Dim CTDblock(71)
    Dim IDsMet1 As Byte
    Dim SmsMet1 As Integer
    Dim NbmMet1 As Byte
    Dim DateMinsMet1 As Integer
    Dim ChansTrans As Byte
    Dim WeekDayn As Byte
    Dim IDsMet2 As Byte
    Dim SmsMet2 As Integer
    Dim NbmMet2 As Byte

    'impostazioni iniziali di CmDialog1 per la scelta del file da leggere
    NewPath sGetAppPath
    On Error GoTo Annulla
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
    CmDialog1.filename = Left$(FileIn, Len(FileIn) - 3) + "dat"
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

    BloccoDati = ""
    TimeOuts = 0
    Bytes = 0
    Intero = 0
    'dati = 0
    iBloccoDati = 0
    ReDim BloccoDati(DFPNT + 100)
    
    Filnb = FreeFile
    Open FileIn For Binary As #1
'    For iBloccoDati = 0 To DFPNT
'        Get #Filnb, , BloccoDati(iBloccoDati)
'    Next
    ReDim BloccoDati(LOF(1) - 1)
    Get #1, , BloccoDati
    Close #1
7
    
'    If iBloccoDati < DFPNT Then
'        Messaggio = "Errore letti" + Str(iBloccoDati) + " dati invece di" + Str(DFPNT)
'        MsgBox (Messaggio)
'        Esci
'    End If

    NBS = BloccoDati(0)
    Debug.Print "NBS="; NBS
    Stringa = bMID(BloccoDati, 1, 71)
    For i = 1 To 71
        MethaneBlock(i) = BloccoDati(i)
        DoEvents
    Next
    'dummy = LeftB(Stringa, 1)
    'Buffer = CByte(dummy)
    Buffer = MethaneBlock(1)
    Debug.Print "Buffer="; Buffer
    IDsMet1 = (Buffer And 240) / 16 'get the left 4 bits and shift them to the right
    Debug.Print "IDsMet1="; IDsMet1
    Debug.Print "sensor= "; SensorType(IDsMet1)
    SmsMet1 = MethaneBlock(1) * 256 + MethaneBlock(2) 'get the two bytes
    SmsMet1 = SmsMet1 And 4095  'erase the left 4 bits
    Debug.Print "SmsMet1="; SmsMet1
    NbmMet1 = MethaneBlock(3)
    Debug.Print "NbmMet1="; NbmMet1
    DateMinsMet1 = MethaneBlock(4) * 256 + MethaneBlock(5)
    WeekDayn = DateMinsMet1 / 1440 + 1
    Debug.Print "WeekDayn="; WeekDayn
    Debug.Print "DateMinsMet1="; DateMinsMet1
    ChansTrans = MethaneBlock(6)
    Debug.Print "ChansTrans="; ChansTrans
    Buffer = MethaneBlock(7)
    IDsMet2 = (Buffer And 240) / 16
    Debug.Print "IDsMet2="; IDsMet2
    Debug.Print "sensor= "; SensorType(IDsMet2)
    SmsMet2 = MethaneBlock(7) * 256 + MethaneBlock(8) 'get the two bytes
    SmsMet2 = SmsMet2 And 4095  'erase the left 4 bits
    Debug.Print "SmsMet2="; SmsMet2
    NbmMet2 = MethaneBlock(9)
    Debug.Print "NbmMet2="; NbmMet2
 
    
    
Annulla:
    Me.MousePointer = vbDefault
    'Imposta la lettura del buffer a tutto il buffer alla volta
    DoEvents

End Sub
