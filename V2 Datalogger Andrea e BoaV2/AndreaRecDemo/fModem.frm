VERSION 5.00
Begin VB.Form fModem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connessione Remota"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "fModem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bRubrica 
      Caption         =   "&Rubrica"
      Height          =   500
      Left            =   1800
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Composizione"
      Height          =   915
      Left            =   150
      TabIndex        =   10
      Top             =   1170
      Width           =   2685
      Begin VB.OptionButton Option2 
         Caption         =   "A toni"
         Height          =   315
         Left            =   690
         TabIndex        =   12
         Top             =   450
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton oImpulsi 
         Caption         =   "A impulsi"
         Height          =   345
         Left            =   690
         TabIndex        =   11
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.CommandButton bAnnulla 
      Caption         =   "&Annulla"
      Height          =   500
      Left            =   3480
      TabIndex        =   8
      Top             =   2400
      Width           =   1005
   End
   Begin VB.CommandButton bChiama 
      Caption         =   "&Chiama"
      Height          =   500
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Frame Frame4 
      Caption         =   "Configurazione porta"
      Height          =   585
      Left            =   2940
      TabIndex        =   5
      Top             =   1500
      Width           =   1605
      Begin VB.TextBox txtPortSettings 
         Height          =   285
         Left            =   210
         TabIndex        =   6
         Text            =   "19200,n,8,1"
         Top             =   210
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "N. Telefono"
      Height          =   915
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   2685
      Begin VB.TextBox txtNumero 
         Height          =   285
         Left            =   420
         TabIndex        =   2
         Top             =   360
         Width           =   1845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modem su COM"
      Height          =   795
      Left            =   2940
      TabIndex        =   0
      Top             =   90
      Width           =   1605
      Begin VB.Label lCOM 
         Caption         =   "COM"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Stringa inizializzazione modem"
      Height          =   885
      Left            =   960
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtModemString 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   330
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   2115
      Left            =   60
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "fModem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
    
    fModem.txtNumero.SetFocus

End Sub

Private Sub Form_Load()
    Dim i As Long
    
    ChiamaFlag = False
    'Nasconde la label1
    Label1.Visible = False
    'Mostra gli altri controlli
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = True
    Frame5.Visible = True
    
    txtModemString.Text = ModemString
    'Stabiliamo una porta COM iniziale
    lCOM.Caption = "COM" & ComPort
    
    Messaggio = sReadINI("Modem", "UltimoNumero", FileIni)
    fModem.txtNumero.Text = Messaggio
    Messaggio = sReadINI("Modem", "UltimoSettings", FileIni)
    If Messaggio = "" Then Messaggio = "57600,n,8,1"
    txtPortSettings = Messaggio '+ ",n,8,1"
    i = Val(sReadINI("Modem", "UltimaCom", FileIni))
    If i = 0 Then i = 1
    ComPort = i
    fMain.MSComm1.RThreshold = 0
    fMain.MSComm1.SThreshold = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If ChiamaFlag = True Then
            OpenCom
            fMain.MSComm1.Output = Chr$(3)
            CloseCom
        End If
        ChiamaFlag = False
        Me.Hide
        Unload Me
        fMain.Show
    End If
End Sub


Private Sub bAnnulla_Click()
    If ChiamaFlag = True Then
        OpenCom
        'fMain.MSComm1.Output = Chr$(3)
        CloseCom
    End If
    ChiamaFlag = False
    Me.Hide
    Unload Me
    fMain.Show
End Sub

Private Sub bRubrica_Click()
    Load fRubrica
    Me.Hide
    fRubrica.Show
End Sub

Private Sub bChiama_Click()
    Dim i As Integer
    Dim risposta As String
    Dim Tempo0 As Long
    Dim DiffTempo As Long
    Dim Contatore As Long
    Dim Msg As String
    
    'Chiude la porta se aperta
    CloseCom
    On Error GoTo GestioneErroreCom
    'Seleziona la porta
    fMain.MSComm1.CommPort = ComPort
    'Setta la porta
    fMain.MSComm1.Settings = txtPortSettings
    'Apre la porta
    'Altri settaggi com indispensabili per un corretto
    'funzionamento del modem
    ChiamaFlag = True
    bChiama.Enabled = False
    
    fMain.MSComm1.Handshaking = comRTS
    fMain.MSComm1.RTSEnable = True
    
    fMain.lblStatus.Caption = "Opening COM port"
    OpenCom
    

    
    
    'controlla che la porta sia valida o esistente
    If ComOk = False Then
        MsgBox ("COM port not valid or in use!!")
        bChiama.Enabled = True
        ChiamaFlag = False
        Exit Sub
    End If
    fMain.lblStatus.Caption = "COM port opened"
    
    
    'Registra il numero di telefono immesso
    Tempo0 = WriteINI("Modem", "UltimoNumero", txtNumero, FileIni)
    If Tempo0 = 0 Then
        Messaggio = "Errore!" + vbCrLf
        Messaggio = Messaggio + "Impossibile registrare il" + vbCrLf
        Messaggio = Messaggio + "numero di telefono sul file" + vbCrLf
        Messaggio = Messaggio + "MH4.ini"
        MsgBox Messaggio, vbOKOnly + vbCritical, "MH4 Errore!"
        ScriviErroreSuLog Messaggio
    End If
    
    'Registra l'ultima velocita' della porta usata
    Tempo0 = WriteINI("Modem", "UltimoSettings", txtPortSettings, FileIni)
    
    'Registra l'ultima com usata
    Tempo0 = WriteINI("Modem", "UltimaCom", ComPort, FileIni)

    If fDebug Then
        Print #fdn, "Chiamata al numero "; txtNumero; " su COM"; ComPort; " "; txtPortSettings
        
    End If

    'Rivela la label1
    Label1.Visible = True
    
    'nasconde gli altri controlli
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    
    'Manda una serie di return per sincronizzare
    'il modem con la porta
    'Label1.Caption = "Sincronizzazione modem" + vbCrLf
    For i = 1 To 5
        DoEvents
        fMain.MSComm1.Output = vbCr
        Sleep (10)
    Next
    'Manda un reset al modem
    'Label1.Caption = Label1.Caption + "Reset modem" + vbCr
    
    'Cambiato da at&f ad atz
    '''''''fMain.MSComm1.Output = "AT&F" + vbCr
    'fMain.MSComm1.Output = "ATZ0" + vbCr
    DoEvents
    'Sleep (50)
    If ChiamaFlag = False Then
        If fDebug Then Print #fdn, "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If

    ''''Risposta = UCase(InputComTimeOut(2))
    ''''Debug.Print "1z "; Risposta;
    ''''Risposta = UCase(InputComTimeOut(2))
    ''''Debug.Print "2 "; Risposta;

    'Label1.Caption = Label1.Caption + "Settaggio modem" + vbCrLf

    'azzera il buffer di input
    'fMain.MSComm1.InBufferCount = 0
    'Manda la stringa di inizializzazione
'    For i = 1 To Len(txtModemString.Text)
'        fMain.MSComm1.Output = Mid(txtModemString.Text, i, 1)
'        Label1.Caption = Label1.Caption + Mid(txtModemString.Text, i, 1)
'        DoEvents
'        Sleep (100)
'    Next
'    fMain.MSComm1.Output = vbCr
'Label1.Caption = Label1.Caption + vbCrLf

    If ChiamaFlag = False Then
        If fDebug Then Print #fdn, "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If

    fMain.lblStatus.Caption = "Setting Modem"
    fMain.MSComm1.Output = "ATL3" + vbCr
    'Sleep (100)

    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "1l3 "; risposta
    Debug.Print "1e0 "; risposta;
    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "2 "; risposta
    Debug.Print "2 "; risposta;

    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If

    fMain.MSComm1.Output = "ATM2" + vbCr
    'Sleep (100)

    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "1m2 "; risposta
    Debug.Print "1e0 "; risposta;
    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "2 "; risposta
    Debug.Print "2 "; risposta;

    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If


    fMain.MSComm1.Output = "AT E0" + vbCr
    'Sleep (100)

    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "1e0 "; risposta
    Debug.Print "1e0 "; risposta;
    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "2 "; risposta
    Debug.Print "2 "; risposta;

    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If

    fMain.MSComm1.Output = "AT X3" + vbCr
    'Sleep (100)
    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "1x1 "; risposta
    Debug.Print "1x1 "; risposta;
    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "2 "; risposta
    Debug.Print "2 "; risposta;

    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If

    fMain.MSComm1.Output = "ATV1" + vbCr
    'Sleep (100)
    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "1v1 "; risposta;
    Debug.Print "1v1 "; risposta;
    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "2 "; risposta;
    Debug.Print "2 "; risposta;

    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If

    fMain.MSComm1.Output = "ATQ0" + vbCr
    'Sleep (200)
    'azzera il buffer di input
    'fMain.MSComm1.InBufferCount = 0

    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "1q0 "; risposta
    Debug.Print "1q0 "; risposta;
    risposta = UCase(InputComTimeOut(2))
    If fDebug Then Print #fdn, "2 "; risposta
    Debug.Print "2 "; risposta;

    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If

    DoEvents

    
    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If


    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If


    i = InStr(risposta, "OK")
    If i = 0 Then
        MsgBox ("Modem doesn't answer!")
        fMain.lblStatus.Caption = "Modem doesn't answer! Aborted"
        Exit Sub
    End If

    'Label1.Caption = Label1.Caption + "Il modem ha risposto! " + Risposta + vbCrLf
    Label1.Caption = Label1.Caption + "Calling " _
    + txtNumero.Text + vbCrLf
    fMain.lblStatus.Caption = "Calling " & txtNumero.Text
    'msgbog (Risposta)
    'azzera il buffer di input
    Sleep (100)
    fMain.MSComm1.InBufferCount = 0
    'Chiama
    If oImpulsi.Value = True Then
        fMain.MSComm1.Output = "ATDP" + Trim(txtNumero.Text) + vbCr
    Else
        fMain.MSComm1.Output = "ATDT" + Trim(txtNumero.Text) + vbCr
    End If
    
    Label1.Caption = Label1.Caption + "Waiting for remote modem" + vbCrLf
    fMain.lblStatus.Caption = "Waiting for remote modem"
    
    Tempo0 = Timer
Retry:
    risposta = UCase(InputComTimeOut(10))
    Debug.Print "3 "; risposta;
    If fDebug Then Print #fdn, "3 "; risposta
    
    If ChiamaFlag = False Then
        Debug.Print "ANNULLATO!"
        fMain.lblStatus.Caption = "Aborted"
        Exit Sub
    End If
    
    If Left(risposta, 1) < " " Then GoTo Retry
    If risposta = "TIMEOUT" Then
        'controllo timeout
        If Timer < Contatore Then       '??????Contatore non è inizializzato!!!
            DiffTempo = Timer + 86400 - Tempo0
        Else
            DiffTempo = Timer - Tempo0
        End If
        Debug.Print DiffTempo
        Label1.Caption = Label1.Caption & Str(DiffTempo) & " "
        GoTo Retry
    End If
    
    'Risposta = Left(Risposta, Len(Risposta) - 2)
    Messaggio = risposta
    risposta = Left(risposta, 4)
    Label1.Caption = Label1.Caption & vbCrLf
    fMain.lblStatus.Caption = risposta
    Select Case risposta
        Case "CONNECT"
            'MsgBox "Connect" + Risposta
            Label1.Caption = Label1.Caption + "Remote Modem Connected" + vbCrLf
            Me.Hide
            Unload Me
            fMain.lblStatus.Caption = "Connected"
            fMain.Show
            Exit Sub
        Case "CONN"
            'MsgBox "Conn" + Risposta
            Label1.Caption = Label1.Caption + "Remote Modem Connected" + vbCrLf
            'Connetti
            fMain.bConnect.Enabled = True
            Me.Hide
            fMain.lblStatus.Caption = "Connected"
            Unload Me
            fMain.Show
            Exit Sub

        Case "BUSY"
            MsgBox "Line Busy" '+ Risposta
            GoTo Fallimento
        Case "NO CARRIER"
            MsgBox "No Carrier!" '+ Risposta
            GoTo Fallimento
        Case "NO C"
            MsgBox "No Carrier!" '+ Risposta
            GoTo Fallimento
        Case "NO DIALTONE"
            MsgBox "No Dial Tone!" + Messaggio
            GoTo Fallimento
        Case "NO D"
            MsgBox "No Dial Tone!" + Messaggio
            GoTo Fallimento
        Case "DELA"
            MsgBox "Modem in Delayed mode! " + Messaggio
            GoTo Fallimento
        Case "ATDT"
            GoTo Retry
        Case Else
            MsgBox "Wrong answer from local modem! " + Messaggio
            GoTo Fallimento
    End Select
    
Exit Sub

GestioneErroreCom:
    Select Case Err.Number
        Case 380  'Settaggi porta errati
            Msg = "Settaggi porta COM errati!"
            MsgBox Msg, , "Errore"
            Err.Clear   ' Cancella i campi dell'oggetto
            'ComOk = False
            Messaggio = "57600,n,8,1"
            txtPortSettings = Messaggio
            Exit Sub
        Case Else
            ErrHandler
            Exit Sub
    End Select

Fallimento:
'fMain.bFine.Enabled = True
'fMain.bConnetti.Enabled = True
'fMain.bRemota.Enabled = True
ChiudiLinea
CloseCom
'Close
Me.Hide
Unload Me
fMain.Show
Exit Sub

End Sub


Private Sub Timeout1()
    'Prova a far ripartire il programma
    Dim Mes As String
    fMain.MSComm1.Output = Chr$(18)
    Mes = "Errore nella comunicazione" + vbCr + "la centralina Andrea non risponde!"
    MsgBox (Mes)
    'UnloadAllForms (Me.Name)
    fMain.MousePointer = vbNormal
    Me.Hide
    Unload Me
    fMain.Show
    'fmain.StatusBar1.Panels(3).Text = "Errore nella comunicazione"
End Sub

Public Sub AbilitaTasti()
    'Abilita i tasti del form principale
    'fMain.bScarica.Enabled = True
    'fMain.bProgramma.Enabled = True
    'fMain.bConnetti.Enabled = True
    'fMain.bTestSensori.Enabled = True
    'fMain.bRemota.Enabled = False
    'fMain.bOrarioModem.Enabled = True
End Sub

Public Sub ChiudiLinea()
    OpenCom
'    fMain.MSComm1.Output = CTRLC
    Sleep (250)
    fMain.MSComm1.Output = "+++"
    Sleep (250)
    fMain.MSComm1.Output = "ATH0"
    'fMain.bRemota.Enabled = True
    'fMain.bConnetti.Enabled = True
End Sub

Private Sub txtPortSettings_DblClick()
    'Me.Hide
    'fVelComModem.Show
End Sub

