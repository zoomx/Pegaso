Attribute VB_Name = "Generale"
Option Explicit

Sub Sleeps(seconds As Double)
   'Wait Seconds seconds
   'There is a control for midnight
   Dim TempTime As Double
   TempTime = Timer
   While Timer - TempTime < seconds
      DoEvents
      If Timer < TempTime Then
         TempTime = TempTime - 24# * 3600#
      End If
   Wend
End Sub
'* Use a timer that has greater resolution (generally 1 millisecond).
'  Some of the other timers have values down to 1 millisecond, but you
'  can 't get the precise 1 millisecond resolution.
'
'   'declare
'   Declare Function timeGetTime Lib "MMSYSTEM" () As Long
'
'   'example
'   oldtime& = timeGetTime()
'
'    'code in here
'
'   deltamillisec& = timeGetTime() - oldtime&
Public Sub UnloadAllForms(sFormName As String)
'Unloading All Forms
'There has been a lot of stories about how Visual Basic
'doesn 't unload the forms when you exit the program. This
'is a 'resource killer'.
'This code unloads all of the forms in your program.
'This is a sub, that you would probably use from the
'Form_Unload of your Main form. So here is the code for
'that:
'
'Call UnloadAllForms Me.Name
'
'Also, here is the code if you're calling it from other
'Subs:
'
'Call UnloadAllForms ""

Dim Form As Form
   For Each Form In Forms
      If Form.Name <> sFormName Then
         Unload Form
         Set Form = Nothing
      End If
   Next Form
End Sub
Public Function OpenFile(File2Open As String, FileMode As String) _
     As Integer

'Then there's opening text files. No need to check if it exists or whatever - just call OpenFile with the
'right parameters (Thandle=OpenFile("TempFile","O") for example) and it will do all the error
'checking for you, passing back the file handle if OK, zero if not

     Dim WhatHandle As Integer
     On Local Error GoTo Op_Error
     WhatHandle = FreeFile()

     Select Case FileMode
     Case "I"
     Open File2Open For Input As WhatHandle
     Case "O"
     Open File2Open For Output As WhatHandle
     Case "A"
     Open File2Open For Append As WhatHandle
     Case "B"
     Open File2Open For Binary As WhatHandle
     End Select

     OpenFile = WhatHandle
     Exit Function

Op_Error:
     OpenFile = 0
End Function
Public Function GetDecimal() As String
'Restituisce il separatore decimale
'C'e' anche la API per leggere direttamente dal registro
'di configurazione ma la stringa esiste solamente
'se si modofocano i valori standard
'La API è commentata perchè sembra che in win98 non funzioni sempre
    Dim Decimale As String
    'Decimale = QueryValue(HKEY_USERS, ".Default\Control Panel\International", "sDecimal")
    If Decimale <> "" Then
        Decimale = Left(Decimale, Len(Decimale) - 1)
    Else
        Decimale = Mid(Format(0.5, "0.0"), 2, 1)
    End If
    GetDecimal = Decimale
End Function
Public Function GetMigliaia() As String
'Restituisce il separatore delle migliaia
'C'e' anche la API per leggere direttamente dal registro
'di configurazione ma la stringa esiste solamente
'se si modofocano i valori standard
'La API è commentata perchè sembra che in win98 non funzioni sempre
    Dim Migliaia As String
    'Migliaia = QueryValue(HKEY_USERS, ".Default\Control Panel\International", "sThousand")
    If Migliaia <> "" Then
        Migliaia = Left(Migliaia, Len(Migliaia) - 1)
    Else
        Migliaia = Mid(Format(1000, "#,###"), 2, 1)
    End If

End Function
Public Sub NewPath(Stringa As String)
'Cambia drive e path contemporaneamente
'Modificare per i drive di rete
'Es. NewPath "d:\temp"
    ChDrive (Left(Stringa, 3))
    ChDir (Stringa)
End Sub
Public Sub StampaAscii(Stringa As String)
    'Stampa il valore dei caratteri ASCII di una stringa
    'nella finestra di Debug
    Dim lStringa As Double
    Dim i As Integer
    lStringa = Len(Stringa)
    If lStringa = 0 Then Exit Sub
    Debug.Print "Risposta"; Stringa; " ";
    For i = 1 To lStringa
        Debug.Print Asc(Mid(Stringa, i, 1));
    Next
    Debug.Print
End Sub
Public Function String2Ascii(Stringa As String) As String
'Converte una stringa nei corrispondenti valori ASCII
'Non viene gestito il CHR$(0)
    Dim lStringa As Double
    Dim i As Integer
    Dim StringAscii As String
    lStringa = Len(Stringa)
    If lStringa = 0 Then Exit Function
    For i = 1 To lStringa
        String2Ascii = String2Ascii + Asc(Mid(Stringa, i, 1)) + " "
    Next
End Function
Public Function Char2ascii(Stringa As String) As String
'Trasforma una stringa contenente caratteri ASCII e non
'ASCII in stringa di codici di caratteri ASCII
'Viene gestito anche il chr$(0)
    Dim lStringa As Integer
    Dim tStringa As String
    Dim i As Integer
    
    lStringa = Len(Stringa)
    For i = 1 To lStringa
        If Mid(Stringa, i, 1) = Chr$(0) Then
            tStringa = tStringa + " " + "00"
        Else
            tStringa = tStringa + Str(Asc(Mid(Stringa, i, 1)))
        End If
    Next
    Char2ascii = tStringa
End Function
Public Function CeSpazio(Percorso As String, Nbytes As Long) As Boolean
    Dim iUnita As Integer
    Dim ok As Integer
    Dim BytesLiberi As Long
    Dim ClustersLiberi As Long
    Dim ClustersRichiesti As Long
    Dim Unita As String
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberOfFreeClusters As Long
    Dim TtoalNumberOfClusters As Long
    Dim VaBene As Boolean
    VaBene = False
    'Identifichiamo l'unità
    'Cerchiamo i :
    iUnita = InStr(Percorso, ":")
    'Prendiamo la lettera prima dei :
    Unita = Mid(Percorso, iUnita - 1, 1) + ":\"
    
    ok = GetDiskFreeSpace(Unita, SectorsPerCluster, _
    BytesPerSector, NumberOfFreeClusters, _
    TtoalNumberOfClusters)
    If ok = 0 Then
        CeSpazio = False
        Exit Function
    End If
    BytesLiberi = NumberOfFreeClusters * SectorsPerCluster * BytesPerSector
    ClustersLiberi = NumberOfFreeClusters * SectorsPerCluster
    ClustersRichiesti = Nbytes / SectorsPerCluster / BytesPerSector
    If ClustersRichiesti > ClustersLiberi Then
        VaBene = False
    Else
        VaBene = True
    End If
    CeSpazio = VaBene
End Function
Public Function stripCrLf(Stringa As String) As String
'elimina i Cr e Lf finali in una stringa
    Dim i As Long
    
    For i = 1 To 2
        If Right(Stringa, 1) = vbCr Or Right(Stringa, 1) = vbLf Then
            Stringa = Left(Stringa, Len(Stringa) - 1)
        End If
    Next
    
    stripCrLf = Stringa
End Function
Public Sub FinePerErrore()
    Dim Mes As String
    CloseCom
    Mes = "Errore interno del programma " + App.Title
    Mes = Mes + Str$(Err.Number) + " " + Err.Description
    MsgBox (Mes)
    'Scaricare tutti i forms
    End
End Sub
Public Sub ErrHandler()
'Gestione errore non altrove gestito
    Dim NomeFileErrors As String
    Dim nfile As Integer
    NomeFileErrors = sGetAppPath() + "Errors.log"
    nfile = FreeFile
    Open NomeFileErrors For Append As nfile
    Print #nfile, "Errore in " + App.Title + " del "; Date$, " alle "; Time$
    Print #nfile, "numero "; Err.Number
    Print #nfile, Err.Description
    Print #nfile, Err.Source
    Print #nfile, "applicazione terminata"
    Close nfile
    NomeFileErrors = "Errore nel'applicazione " + App.Title + vbCrLf
    NomeFileErrors = NomeFileErrors + Str(Err.Number) + " " + Err.Description + vbCrLf
    NomeFileErrors = NomeFileErrors + "L'errore è stato salvato nel file errors.log" + vbCrLf
    NomeFileErrors = NomeFileErrors + "L'applicazione verrà chiusa"
    MsgBox (NomeFileErrors)
    'chiude tutti i forms e termina l'applicazione
    'Form_Unload
    End

End Sub

Public Sub ScriviErrore(errore As String)
'Scrive un errore generico sul file errors.log
    Dim NomeFileErrors As String
    Dim nfile As Integer
    NomeFileErrors = sGetAppPath() + "Errors.log"
    nfile = FreeFile
    Open NomeFileErrors For Append As nfile
    Print #nfile, "Errore in "; App.Title; " del "; Date$; " alle "; Time$
    Print #nfile, Err.Description
    Print #nfile, Err.Source
    Print #nfile, "applicazione terminata"
    Close nfile

End Sub
Public Function ScriviErroreSuLog(errore As String) As Boolean
    Dim FileLogName As String
    Dim nf As Long
    
    FileLogName = sGetAppPath + App.EXEName + ".log"
    nf = FreeFile
    Open FileLogName For Append As #nf
    Print #nf, "---------------------------------------------"
    Print #nf, Now, App.EXEName
    Print #nf, "Err number -->"; Err.Number
    Print #nf, Err.Description
    Print #nf, errore
    Close nf
    ScriviErroreSuLog = True
    
End Function

Function sGetAppPath() As String
'*Returns the application path with a trailing \.      *
'*To use, call the function [SomeString=sGetAppPath()] *
Dim sTemp As String
        sTemp = App.Path
        If Right$(sTemp, 1) <> "\" Then sTemp = sTemp + "\"
        sGetAppPath = sTemp
End Function

Public Function Formato(Numero As Double, StringaFormato As String) As String
'Modifica dell'istruzione format
'Tiene conto del fatto che l'istruzione Format mette come
'separatore dei decimali cio' che gli viene indicato
'dalle impostazioni internazionali. Quindi se il separatore non
'è un punto elimina il separatore e ci mette il punto.
'In pratica sostituisce la virgola col punto
'Richiede che in Migliaia e Decimale ci siano i separatori
'corrispondenti, letti dal registro di configurazione
'Si potrebbe evitare di sapere preventivamente di conoscere
'Il separatore cercando la virgola ma questa potrebbe
'essere presente come separatore delle migliaia

    Dim i As Integer
    Dim LungString2 As Integer
    Dim Stringa2 As String
    Stringa2 = Format(Numero, StringaFormato)
    LungString2 = Len(Stringa2)
    'il decimale e' un punto o una virgola?
    If Decimale <> "." Then
        'Si, sostituiamolo con il punto
        i = InStr(Stringa2, Decimale)
        Stringa2 = Left(Stringa2, i - 1) + "." + Right(Stringa2, LungString2 - i)
    End If
    Formato = Stringa2
End Function

Public Function SetInIDE() As Boolean
'Restituisce True (Vero) se si è in ambiente di programmazione
'False se il programma è compilato
    On Error GoTo DivideError
    Debug.Print 1 / 0
    SetInIDE = False
    Exit Function
    
DivideError:
    SetInIDE = True
End Function
Public Sub OpenCom()
    'Apre la porta com
    'Se e' andata bene ComOk e' True altrimenti e' False
    Dim Msg As String

    On Error GoTo ErroreCom
    ComOk = False
    'Apre la porta seriale se non è già aperta
    If fMain.MSComm1.PortOpen = False Then fMain.MSComm1.PortOpen = True
    ComOk = True
    Exit Sub
ErroreCom:
    Select Case Err.Number
        Case 8005  'La Com è già aperta
            Msg = "Errore la porta Com" + Str$(ComPort) + " è già in uso"
            MsgBox Msg, , "Errore"
            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case 8002
            Msg = "Errore la porta Com" + Str$(ComPort) + " non esiste!"
            MsgBox Msg, , "Errore"
            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case Else
            ErrHandler
            Exit Sub
    End Select

End Sub

Public Sub CloseCom()
    'Chiude la porta seriale se non è già chiusa
    fMain.MSComm1.InBufferCount = 0
    If fMain.MSComm1.PortOpen = True Then fMain.MSComm1.PortOpen = False
End Sub
Public Sub WaitCom()
'Aspetta che sulla COM ci siano dei caratteri.
'Senza TIMEOUT!
    Do
        DoEvents
    Loop Until fMain.MSComm1.InBufferCount >= 1
End Sub

Public Function InputComTimeOut(TimeOut As Integer) As String
'Attende un input il cui terminatore e' vbLF
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As String
    Dim dummy As String
    
        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents

        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            dummy = ""
            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until dummy = vbLf Or (Timer > TimeStop)
                DoEvents
                dummy = fMain.MSComm1.Input
                Linea = Linea + dummy
            Loop
        Else
            Linea = "TimeOut"
        End If
        InputComTimeOut = Linea

End Function

Public Function InputComTimeOutTerm(TimeOut As Integer, Terminator As Byte) As String
'Attende un input il cui terminatore e' Terminator
'Con TIMEOUT

        Dim TimeStop As Long
        Dim Linea As String
        Dim dummy As String

        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            dummy = ""
            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until dummy = Chr(Terminator) Or (Timer > TimeStop)
                DoEvents
                dummy = fMain.MSComm1.Input
                Linea = Linea + dummy
            Loop
        Else
            Linea = "TimeOut"
        End If
        InputComTimeOutTerm = Linea

End Function
Public Function InputComTimeOutBin(TimeOut As Integer, NumByte As Integer) As String
'Attende un input binario senza terminatore
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As String
    'Dim dummy As String
    
        'fMain.MSComm1.InputMode = comInputModeBinary

        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = NumByte
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= NumByte) Or (Timer > TimeStop)
'        If fMain.MSComm1.InBufferCount >= NumByte Then
'            linea = ""
'        Else
'            linea = "TimeOut"
'        End If
        Linea = fMain.MSComm1.Input
        If Linea <> "" Then
            Debug.Print "inputCom--->"; Char2ascii(Linea)
        End If
        InputComTimeOutBin = Linea

End Function

Public Function InputComTimeOutBin2(TimeOut As Integer, NumByte As Integer) As String
'Attende un input binario senza terminatore
'Con TIMEOUT sul singolo carattere!!!
    Dim TimeStop As Long
    Dim Linea As String
    Dim TimerOut As Boolean
    Dim BloccoDati(32768) As Byte
    Dim Blocco() As Byte
    Dim iBloccoDati As Long
    Dim TimeOuts As Integer
    Dim i As Long
    

    iBloccoDati = 0
    'ReDim BloccoDati(1000)
    Do
        DoEvents
        TimeStop = Timer + 2
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount = 0 Then
            TimeOuts = TimeOuts + 1
        Else
            'dati = dati + fMain.MSComm1.InBufferCount
            Blocco = fMain.MSComm1.Input
            For i = LBound(Blocco) To UBound(Blocco)
                BloccoDati(iBloccoDati) = Blocco(i)
                iBloccoDati = iBloccoDati + 1
            Next i
            TimeOuts = 0
        End If
        If TimeOuts > 3 Then Exit Do
        
        
        DoEvents
        
    Loop Until TimeOuts > 5
    
    For i = 0 To iBloccoDati - 1
        InputComTimeOutBin2 = InputComTimeOutBin2 + Chr$(BloccoDati(i))
    Next i

End Function

Public Function bMID(matrice() As Byte, inizio As Long, lunghezza As Long) As String
'Estrae una stringa da un vettore di bytes
'Sintassi come istruzione MID
    Dim Stringa As String
    Dim i As Long

    For i = inizio To inizio + lunghezza - 1
        Stringa = Stringa + Chr(matrice(i))
        DoEvents
    Next
    bMID = Stringa
End Function

