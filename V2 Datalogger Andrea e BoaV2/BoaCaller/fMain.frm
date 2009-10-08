VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form fMain 
   Caption         =   "Caller"
   ClientHeight    =   6885
   ClientLeft      =   825
   ClientTop       =   2040
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10425
   Begin VB.CommandButton bTranscode2 
      Caption         =   "Transcode type 2"
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   2655
      Begin VB.Label lnlines 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "lines of"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lnlinesgot 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Got"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton bDownload 
      Caption         =   "&Download Last File"
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton bTest 
      Caption         =   "&Test"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton bTranscode 
      Caption         =   "Transcode type 1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lMonitor 
      BorderStyle     =   1  'Fixed Single
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   9855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bDownload_Click()
    'Setta la com
    'Apre la Com
    'Fa la telefonata
    'Chiede l'ultimo file aperto
    'Calcola l'ultimo file chiuso
    'Lo scarica
    'Chiude la comunicazione
    'Chiude la porta
    
    
    Dim i As Integer
    Dim Risposta As String
    Dim Tempo0 As Long
    Dim DiffTempo As Long
    Dim Contatore As Long
    Dim Msg As String
    Dim AllOk As Boolean
    Dim ActualID As Long
    Dim LastIDtoDownload As Long
    Dim Stringbuffer As String
    Dim Stringa As String
    Dim Messaggio As String
    Dim nLines As String
    Dim nLinesL As Long
    Dim LastPacket As Boolean
    
    Dim FilePrefix As String
    Dim FileName As String
    Dim FileN As Long
    Dim FileN2 As Long
    Dim FileN3 As Long
    Dim FileN4 As Long
    Dim FileN5 As Long
    Dim FileBuffer As String
    Dim FileBuffer2 As String
    Dim FileBuffer3 As String
    Dim FileBuffer4 As String
    
    Dim RowsNumber As Byte
    Dim RowIndex As Long
    Dim RowsGot As Long
    Dim AllDone As Boolean
    Dim PhoneNumber As String
    Dim Chiamate As Integer
    
    SubName = "bDownload_Click"
    LastPacket = False
    
    
    'Setta la COM
    CommPort = 3
    CommSettings = "115200,n,8,1"
    PhoneNumber = "3480948945"
    AllOk = CallNumber(PhoneNumber)

    Connected = AllOk
    
    bDownload.Enabled = False
    
    If Connected = False Then GoTo Closing  'Si potrebbe fare meglio

    Chiamate = 1
    
    'Chiede l'ultimo file aperto
    AllOk = SendLastFileRequest
    'controllare che non ci sia un errore
    
    'Prende la risposta
    Risposta = InputComTimeOutTerm(20, 0)
    lMonitor.Caption = "Answer lenght->" & Str(Len(Risposta)) + " " + Risposta + vbCr
    
    If Len(Risposta) < 10 Then GoTo Closing     'Cambiare per la lunghezza reale
    
    'Controlla che la risposta sia corretta
    Stringbuffer = Mid(Risposta, 2, 1)
    If Asc(Stringbuffer) = 16 Then
        Stringbuffer = "Response is correct!"
    Else
        Stringbuffer = "Response is wrong->" & Str(Asc(Stringbuffer))
    End If
    
    'Decodifica la risposta
    Stringbuffer = Mid(Risposta, 3, 4)
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    ActualID = Val(Stringbuffer)
    lMonitor.Caption = lMonitor.Caption & "ID of actual file open ->" & ActualID & vbCr
    
    Stringbuffer = Mid(Risposta, 11, 2) 'Year
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringbuffer + "/"
    
    Stringbuffer = Mid(Risposta, 13, 2) 'Month
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringa + Stringbuffer + "/"
    
    Stringbuffer = Mid(Risposta, 15, 2) 'Day
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringa + Stringbuffer + " "
    
    Stringbuffer = Mid(Risposta, 9, 2) 'Hour
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringa + Stringbuffer + ":"
   
    Stringbuffer = Mid(Risposta, 7, 2) 'Minute
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringa + Stringbuffer
    
    lMonitor.Caption = lMonitor.Caption & "Date ->" & Stringa & vbCr
    'Controllare che la data sia quella di oggi
    
    'Calcola l'ultimo file chiuso
    LastIDtoDownload = ActualID - 1                                    '***********************************************
    If LastIDtoDownload <= 0 Then LastIDtoDownload = 30
    
    FilePrefix = ""
    FilePrefix = FilePrefix + Format(Year(Now - 1), "0000")
    FilePrefix = FilePrefix + Format(Month(Now - 1), "00")
    FilePrefix = FilePrefix + Format(Day(Now - 1), "00")
    FilePrefix = FilePrefix + "_" + Format(Hour(Now), "00")
    FilePrefix = FilePrefix + Format(Minute(Now), "00")
    FilePrefix = FilePrefix + Format(Second(Now), "00") + "_" + Trim(Str(LastIDtoDownload))


    
    'chiede informazioni sul numero di linee
    ChiamaFlag = SendFileInfoRequest(LastIDtoDownload)
    'controllare che non ci sia un errore

    'Prende la risposta
    Risposta = InputComTimeOutTerm(20, 0)
    lMonitor.Caption = "Answer lenght->" & Str(Len(Risposta)) + " " + Risposta + vbCr

    'Controlla che la risposta sia corretta
    Stringbuffer = Mid(Risposta, 2, 1)
    If Asc(Stringbuffer) = 18 Then
        Stringbuffer = "Response is correct!"
    Else
        Stringbuffer = "Response is wrong->" & Str(Asc(Stringbuffer))
    End If


    'Interpreta la risposta
    'ID
    Stringbuffer = Mid(Risposta, 3, 4)
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    ActualID = Val(Stringbuffer)
    lMonitor.Caption = lMonitor.Caption & "Answer ID ->" & ActualID & vbCr
    'Lines
    Stringbuffer = Mid(Risposta, 7, 4)
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    nLines = Stringbuffer
    nLinesL = Val(Stringbuffer)
    'lMonitor.Caption = lMonitor.Caption & "ID of actual file open ->" & ActualID & vbCr
    lMonitor.Caption = lMonitor.Caption & "Number of lines ->" & nLinesL & vbCr
    lnlines.Caption = nLines
    

    'Lo scarica
    FileBuffer = ""
    
    'determino quante righe scaricare alla volta
    RowsNumber = 4  'prendo 4 righe alla volta
    RowIndex = 1    'parto dalla prima riga
    RowsGot = 0 'non ho preso ancora righe
    'ciclo di richiesta
    AllDone = False
    
    FileName = App.Path + "\" + FilePrefix + "Logtotale.txt"
    FileN2 = FreeFile
    Open FileName For Output As FileN2

    
    Do
        AllOk = SendPacketDownloadRequest(LastIDtoDownload, RowsNumber, RowIndex)
        'prendo il pacchetto
        'Risposta = InputComTimeOutTerm2(20, 0) 'ricevo un pacchetto di 621 caratteri!!!!
        Risposta = InputComTimeOutTerm(20, 0)
        lMonitor.Caption = "Answer lenght->" & Str(Len(Risposta)) + vbCrLf + Risposta + vbCr
        
        Print #FileN2, Risposta;
        
        If Risposta = "TimeOut" Then
            'Siamo ancora in linea?
            Connected = fMain.MSComm1.CDHolding
        End If
        
        If Connected = True And Len(Risposta) > 50 Then 'se siamo ancora in linea è la risposta è lunga abbastanza...
            'Lo unisco al resto togliendo i due byte iniziali e i due byte finali.
            'FileBuffer = FileBuffer + Mid(Risposta, 4, Len(Risposta) - 4)
            'FileBuffer2 = FileBuffer2 + Mid(Risposta, 2, Len(Risposta) - 4)
            FileBuffer3 = FileBuffer3 + Mid(Risposta, 3, Len(Risposta) - 4) 'il migliore
            'FileBuffer4 = FileBuffer4 + Mid(Risposta, 3, 608)
            
            'calcolo quante linee ho scaricato
            RowsGot = RowsGot + RowsNumber
            lnlinesgot.Caption = Trim(Str(RowsGot))
            
            
            'controllo che non ho finito
            If RowsGot = nLinesL Then
                AllDone = True
                Exit Do
            End If
            If LastPacket = True Then
                AllDone = True
                Exit Do
            End If

            
            'calcolo le prossime righe
            RowIndex = RowIndex + RowsNumber
            If RowIndex + RowsNumber > nLines Then
                'RowsNumber = nLinesL - RowIndex
                LastPacket = True
            End If
        End If
        
        If Connected = False Then
                'aspetta un po
                Connected = CallNumber(PhoneNumber)
                Chiamate = Chiamate + 1
        End If
    
        If Chiamate > 10 Then Exit Do
    Loop Until AllDone = True
    'Chiude la comunicazione
    
    'salva il file
    If AllDone = True Then
        FileName = App.Path + "\" + FilePrefix + "File.txt"
        FileN = FreeFile
        Open FileName For Output As #FileN
        'Print #FileN, FileBuffer
        'Print #FileN, "-------------------------------------------------------------------------------------------------"
        'Print #FileN, FileBuffer2
        'Print #FileN, "-------------------------------------------------------------------------------------------------"
        Print #FileN, FileBuffer3;
        'Print #FileN, "-------------------------------------------------------------------------------------------------"
        Close FileN
    End If
    Close FileN2
    
Closing:
    bDownload.Enabled = True
    'Chiude la porta
    MSComm1.PortOpen = False
    lMonitor.Caption = lMonitor.Caption + vbCrLf + "Closing COM port"
    
End Sub

Private Sub bEnd_Click()
    End
End Sub

Private Sub bTranscode_Click()
    Dim FileName As String
    Dim LineBuffer As String
    Dim Linea As String
    Dim LineCount As Long
    Dim LineLenght As Long
    Dim Lines As Integer
    
    SubName = "bTranscode_Click"
    
    bTranscode.Enabled = False
    
    FileName = App.Path + "\" + "1.log"
    Open FileName For Input As #1
    Line Input #1, LineBuffer
    Close 1
    FileName = App.Path + "\" + "1.txt"
    Open FileName For Output As #1
    LineLenght = Len(LineBuffer)
    Lines = LineLenght / 60
    Debug.Print LineLenght; "->"; Lines; " lines"
    For LineCount = 0 To Lines - 1
        Linea = Mid(LineBuffer, LineCount * 60 + 1, 60)
        'Print #1, Linea
        Linea = ParseLine(Linea)
        Print #1, Linea
    Next LineCount
    Close 1
    bTranscode.Enabled = True
End Sub

Private Sub bTest_Click()
    Dim Risposta As String
    Dim Stringbuffer As String
    Dim ActualID As Long
    Dim answer As Boolean
    Dim i As Long
    Dim j As Integer
    
    Dim FileName As String
    Dim FileName2 As String
    Dim FileN As Long
    Dim FileN2 As Long
    Dim FileBuffer As String
    Dim LineBuffer(611) As Byte

    FileName = App.Path + "\" + "6Logtotale.txt"
    FileN = FreeFile
    Open FileName For Binary As #FileN

    FileName2 = App.Path + "\" + "6file.txt"
    FileN2 = FreeFile
    Open FileName2 For Output As #FileN2
    
    FileBuffer = ""
    For i = 0 To 1
        
        Risposta = ""
        Get #FileN, CLng(i * 612 + 1), LineBuffer
        'Get #FileN, 1, LineBuffer
        For j = 0 To 611
            Risposta = Risposta + Chr(LineBuffer(j))
            'Debug.Print Chr(LineBuffer(j));
        Next j
        'Debug.Print
        Debug.Print Risposta; "*"; Len(Risposta)
        FileBuffer = FileBuffer + Mid(Risposta, 3, Len(Risposta) - 4)
        Debug.Print Mid(Risposta, 4, Len(Risposta) - 4); "*"; Len(Mid(Risposta, 4, Len(Risposta) - 4))
        Risposta = TranscodeLine(Mid(Risposta, 4, Len(Risposta) - 4))
        Debug.Print Risposta
        Risposta = ParseLine2(FileBuffer)
        Debug.Print Risposta
    Next i
    Print #FileN2, FileBuffer
    Close FileN
    Close FileN2
    
    Exit Sub
    
    MSComm1.CommPort = 6
    'MSComm1.Handshaking = comRTS
    MSComm1.PortOpen = True
    
    Risposta = InputComTimeOutTerm4(20, 0)
    
    Exit Sub
    
    answer = SendPacketDownloadRequest(4, 4, 1)
    Exit Sub
    
    MSComm1.CommPort = 6
    MSComm1.Handshaking = comRTS
    MSComm1.PortOpen = True
    
    Do
        DoEvents
        Risposta = InputComTimeOutTerm(10, 0)
'        If Len(Risposta) > 0 Then
            Debug.Print "Risposta->"; Risposta; "   "; Len(Risposta)
            Stringbuffer = Stringbuffer + Risposta
            Debug.Print Stringbuffer; "    "; Len(Stringbuffer)
'        End If
        DoEvents
    Loop Until ActualID = 10
    
    
    answer = SendPacketDownloadRequest(21, 4, 1)
    
    Exit Sub
    
    Risposta = Chr(1)
    'Risposta = Risposta + Chr(16)
    'Risposta = Risposta + "5100B371906001" 'ID e data
    
    Risposta = Risposta + Chr(18)
    Risposta = Risposta + "51005750V" 'ID e numero di linee
    Risposta = Risposta + Chr(0)

    'Decodifica la risposta
    
    Stringbuffer = Mid(Risposta, 3, 4)
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    ActualID = Val(Stringbuffer)
    'lMonitor.Caption = lMonitor.Caption & "ID of actual file open ->" & ActualID & vbCr
    lMonitor.Caption = lMonitor.Caption & "ID of file to download ->" & ActualID & vbCr
    
    Stringbuffer = Mid(Risposta, 7, 4)
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    ActualID = Val(Stringbuffer)
    'lMonitor.Caption = lMonitor.Caption & "ID of actual file open ->" & ActualID & vbCr
    lMonitor.Caption = lMonitor.Caption & "Number of lines ->" & ActualID & vbCr
        
    Exit Sub
    
    
    Stringbuffer = Mid(Risposta, 11, 2) 'Year
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringbuffer + "/"
    
    Stringbuffer = Mid(Risposta, 13, 2) 'Month
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringa + Stringbuffer + "/"
    
    Stringbuffer = Mid(Risposta, 15, 2) 'Day
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringa + Stringbuffer + " "
    
    Stringbuffer = Mid(Risposta, 9, 2) 'Hour
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringa + Stringbuffer + ":"
   
    Stringbuffer = Mid(Risposta, 7, 2) 'Minute
    Stringbuffer = SwapString(Stringbuffer)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    Stringa = Stringa + Stringbuffer
    
    lMonitor.Caption = lMonitor.Caption & "Date ->" & Stringa & vbCr
    'Controllare che la data sia quella di oggi
    Risposta = Hex(ActualID)
    Debug.Print Risposta
    If Len(Risposta) = 2 Then Risposta = "00" + Risposta
    If Len(Risposta) = 3 Then Risposta = "0" + Risposta
    Debug.Print Risposta
    Risposta = SwapString(Risposta)
    Debug.Print Risposta
    
    SendFileInfoRequest (21)
End Sub

Private Sub bTranscode2_Click()
    Dim FileN As Long
    Dim FileN2 As Long
    Dim FileName As String
    Dim LineBuffer As String
    Dim Linea As String
    Dim LineCount As Long
    Dim LineLenght As Long
    Dim Lines As Integer
    
    SubName = "bTranscode2_Click"
    
    bTranscode2.Enabled = False
    
    FileName = App.Path + "\" + "20091006_105558_12File.txt"
    FileN = FreeFile
    Open FileName For Input As #FileN
    Line Input #1, LineBuffer
    Debug.Print Len(LineBuffer)
    Close 1
    'Exit Sub
    FileName = App.Path + "\" + "20091006_105558_12File.txt.csv"
    FileN = FreeFile
    Open FileName For Output As #FileN
    
    FileName = App.Path + "\" + "20091006_105558_12File.txt2.txt"
    FileN2 = FreeFile
    Open FileName For Output As #FileN2

    LineLenght = Len(LineBuffer)
    'LineBuffer = Right(LineBuffer, LineLenght - 2)
    LineLenght = Len(LineBuffer)
    Lines = LineLenght / 152 '76*2
    'Debug.Print Lines; " lines"
    Debug.Print LineLenght; "->"; Lines; " lines"
    For LineCount = 0 To Lines - 1
        Linea = Mid(LineBuffer, LineCount * 152 + 1, 152)
        Print #FileN2, Linea
        Print #FileN2, "-------------------------------------------"
        Print #FileN2,
        'Debug.Print Linea
        'Linea = TranscodeLine(Linea)
        Print #FileN2, Linea
        Linea = ParseLine2(Linea)
        Print #FileN, Linea
    Next LineCount
    Close FileN
    Close FileN2

    bTranscode2.Enabled = True
End Sub
