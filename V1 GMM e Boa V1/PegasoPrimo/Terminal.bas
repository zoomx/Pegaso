Attribute VB_Name = "Terminal"
Option Explicit
Public hLogFile As Integer ' Gestore del file registro aperto.
Public StartTime As Date   ' Memorizza il momento di partenza del timer della porta
Public Echo As Boolean        ' Flag per l'eco attivato/disattivato.
Public Ret As Integer
Public isTerminal As Boolean
Public CancelSend As Boolean

' Questa routine aggiunge dati alla proprietà Text del controllo Term.
' Filtra anche i caratteri di controllo, come BACKSPACE,
' ritorno a capo e avanzamento riga e scrive i dati
' in un file registro aperto.
' I caratteri BACKSPACE cancellano i caratteri a sinistra,
' sia nella proprietà Text che in una stringa passata.
' I caratteri di avanzamento riga vengono aggiunti dopo tutti i
' caratteri di ritorno a capo. Vengono inoltre controllate
' le dimensioni della proprietà Text del controllo Term in modo
' che non siano mai maggiori del valore di MAXTERMSIZE.
Public Static Sub ShowData(Term As Control, Data As String)
    On Error GoTo Handler
    Const MAXTERMSIZE = 16000
    Dim TermSize As Long, i
    
    ' Controlla che le dimensioni del testo esistente non siano
    ' troppo grandi.
    TermSize = Len(Term.Text)
    If TermSize > MAXTERMSIZE Then
       Term.Text = Mid$(Term.Text, 4097)
       TermSize = Len(Term.Text)
    End If

    ' Punta alla fine dei dati di Term.
    Term.SelStart = TermSize

    ' Filtra/gestisce i caratteri BACKSPACE.
    Do
       i = InStr(Data, Chr$(8))
       If i Then
          If i = 1 Then
             Term.SelStart = TermSize - 1
             Term.SelLength = 1
             Data = Mid$(Data, i + 1)
          Else
             Data = Left$(Data, i - 2) & Mid$(Data, i + 1)
          End If
       End If
    Loop While i

    ' Elimina i caratteri di avanzamento riga.
    Do
       i = InStr(Data, Chr$(10))
       If i Then
          Data = Left$(Data, i - 1) & Mid$(Data, i + 1)
       End If
    Loop While i

    ' Verifica che tutti i caratteri di ritorno a capo siano
    ' seguiti da un carattere di avanzamento riga.
    i = 1
    Do
       i = InStr(i, Data, Chr$(13))
       If i Then
          Data = Left$(Data, i) & Chr$(10) & Mid$(Data, i + 1)
          i = i + 1
       End If
    Loop While i

    ' Assegna i dati filtrati alla proprietà SelText.
    Term.SelText = Data
  
    ' Se richiesto, registra i dati nel file.
    If hLogFile Then
       i = 2
       Do
          Err = 0
          Put hLogFile, , Data
          If Err Then
             i = MsgBox(Error$, 21)
             If i = 2 Then
                'frmTerminal.mnuCloseLog_Click
                'frmterminal.mnuCloseLog.
             End If
          End If
       Loop While i <> 2
    End If
    Term.SelStart = Len(Term.Text)
Exit Sub

Handler:
    MsgBox Error$
    Resume Next
End Sub
' Richiama questa funzione per avviare il timer della durata della connessione
Public Sub StartTiming()
    StartTime = Now
    frmTerminal.Timer1.Enabled = True
End Sub
' Richiama questa funzione per interrompere il calcolo della durata della connessione
Public Sub StopTiming()
    frmTerminal.Timer1.Enabled = False
    frmTerminal.sbrStatus.Panels("ConnectTime").Text = ""
End Sub

