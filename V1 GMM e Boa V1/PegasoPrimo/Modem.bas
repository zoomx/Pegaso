Attribute VB_Name = "Modem"
Option Explicit
'Aggiunte per la comunicazione via modem
Type TipoConnessione
    Locale As Boolean
    nTelefono As String
    Manuale As Boolean
    Ora As String
    Password As String
    ComPort As Integer
    ModemString As String
    PortConfiguration As String
End Type

'Variabile che contiene il tipo di connessione
Public CfgCon As TipoConnessione

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
        fMain.MSComm1.Output = "AT" + comando + vbCr
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

Public Function MandaComando2(comando As String) As String
'Manda un comando al modem e attende la risposta senza Time Out
    Dim Linea As String
    Dim Dummy As String
    Dim TimeStop As Long

        If fMain.MSComm1.PortOpen = False Then
            Linea = "Porta non aperta"
            GoTo fine
        End If
        fMain.MSComm1.InBufferCount = 0
        fMain.MSComm1.Output = "AT" + comando + vbCr
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents
        Loop Until fMain.MSComm1.InBufferCount >= 1
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            TimeStop = Timer + 1 ' Imposta l'ora di fine
            Do Until Dummy = vbLf Or (Timer > TimeStop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
            Loop
        End If
        Do
            DoEvents
        Loop Until fMain.MSComm1.InBufferCount >= 1
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            'timestop = Timer + 1 ' Imposta l'ora di fine
            Do Until Dummy = vbLf 'Or (Timer > timestop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
            Loop
        End If

fine:
    fMain.MSComm1.InBufferCount = 0
    If Len(Linea) > 2 Then
        Linea = Left(Linea, Len(Linea) - 2)
    End If
    MandaComando2 = Linea
End Function
