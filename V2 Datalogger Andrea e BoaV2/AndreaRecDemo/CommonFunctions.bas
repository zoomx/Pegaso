Attribute VB_Name = "CommonFunctions"
Option Explicit

Function sGetAppPath() As String
'*Returns the application path with a trailing \.      *
'*To use, call the function [SomeString=sGetAppPath()] *
Dim sTemp As String
        sTemp = App.Path
        If Right$(sTemp, 1) <> "\" Then sTemp = sTemp + "\"
        sGetAppPath = sTemp
End Function


Public Sub ErrHandler()
'Gestione errore non altrove gestito
    Dim NomeFileErrors As String
    Dim nFile As Integer
    NomeFileErrors = sGetAppPath() + "Errors.log"
    nFile = FreeFile
    Open NomeFileErrors For Append As nFile
    Print #nFile, "Errore in " + App.Title + " del "; Date$, " alle "; Time$
    Print #nFile, "numero "; Err.Number
    Print #nFile, Err.Description
    Print #nFile, Err.Source
    Print #nFile, "applicazione terminata"
    Close nFile
    NomeFileErrors = "Errore nel'applicazione " + App.Title + vbCrLf
    NomeFileErrors = NomeFileErrors + Str(Err.Number) + " " + Err.Description + vbCrLf
    NomeFileErrors = NomeFileErrors + "L'errore è stato salvato nel file errors.log" + vbCrLf
    NomeFileErrors = NomeFileErrors + "L'applicazione verrà chiusa"
    MsgBox (NomeFileErrors)
    'chiude tutti i forms e termina l'applicazione
    'Form_Unload
    End

End Sub

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

Public Function InputComTimeOut(TimeOut As Integer) As String
'Attende un input il cui terminatore e' vbLF
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As String
    Dim Dummy As String
    
        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents

        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until Dummy = vbLf Or (Timer > TimeStop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
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
        Dim Dummy As String

        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until Dummy = Chr(Terminator) Or (Timer > TimeStop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
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
        'Debug.Print Timer; " "; TimeStop
        fMain.MSComm1.InputLen = NumByte
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= NumByte) Or (Timer > TimeStop)
'        If fMain.MSComm1.InBufferCount >= NumByte Then
'            linea = ""
'        Else
'            linea = "TimeOut"
'        End If
        'Debug.Print fMain.MSComm1.InBufferCount
        Linea = fMain.MSComm1.Input
'        If Linea <> "" Then
'            Debug.Print "inputCom--->"; Char2ascii(Linea)
'        End If
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
Public Function InputComTimeOutBin3(TimeOut As Integer) As Byte
'Attende un input binario di un carattere senza terminatore
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As Byte
    Dim Stringa As String
    'Dim dummy As String
    
        'fMain.MSComm1.InputMode = comInputModeBinary

        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount = 0 Then
            Linea = 32
        Else
            Linea = Asc(fMain.MSComm1.Input)
        End If
        
'        If Linea <> "" Then
'            'Debug.Print "inputCom--->"; Char2ascii(Linea)
'        End If
        InputComTimeOutBin3 = Linea
        Debug.Print Linea

End Function

Public Function MandaComando(comando As String, TmOut As Integer) As String
'Manda un comando al modem e attende la risposta con Time Out
    Dim Linea As String
    Dim Dummy As String
    Dim TimeStop As Long
    
        If fMain.MSComm1.PortOpen = False Then
            Linea = "Porta non aperta"
            GoTo fine
        End If
        fMain.MSComm1.InBufferCount = 0
        fMain.MSComm1.Output = "AT" + comando + vbCrLf
        TimeStop = Timer + TmOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            TimeStop = Timer + TmOut ' Imposta l'ora di fine
            Do Until Dummy = vbLf Or (Timer > TimeStop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
            Loop
            'StampaAscii (Linea)
            Linea = ""
            Dummy = ""
            TimeStop = Timer + TmOut ' Imposta l'ora di fine
            Do Until Dummy = vbLf Or (Timer > TimeStop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
            Loop
            'StampaAscii (Linea)
        End If
fine:
    fMain.MSComm1.InBufferCount = 0
    If Len(Linea) > 2 Then
        Linea = Left(Linea, Len(Linea) - 2)
    End If
    'Debug.Print "linea depurata-->"; Linea
    MandaComando = Linea
End Function

