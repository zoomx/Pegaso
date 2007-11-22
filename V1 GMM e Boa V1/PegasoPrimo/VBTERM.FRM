VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTerminal 
   Caption         =   "Terminale"
   ClientHeight    =   4935
   ClientLeft      =   2940
   ClientTop       =   2055
   ClientWidth     =   7155
   ForeColor       =   &H00000000&
   Icon            =   "VBTERM.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7155
   Begin ComctlLib.Toolbar tbrToolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OpenLogFile"
            Description     =   "Apri file registro..."
            Object.ToolTipText     =   "Apri file registro..."
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "CloseLogFile"
            Description     =   "Chiudi file registro"
            Object.ToolTipText     =   "Chiudi file registro"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DialPhoneNumber"
            Description     =   "Componi numero telefonico..."
            Object.ToolTipText     =   "Componi numero telefonico..."
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "HangUpPhone"
            Description     =   "Interrompi comunicazione"
            Object.ToolTipText     =   "Interrompi comunicazione"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Properties"
            Description     =   "Proprietà..."
            Object.ToolTipText     =   "Proprietà..."
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "TransmitTextFile"
            Description     =   "Trasmetti file di testo..."
            Object.ToolTipText     =   "Trasmetti file di testo..."
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   4000
         TabIndex        =   2
         Top             =   75
         Width           =   240
         Begin VB.Image imgConnected 
            Height          =   240
            Left            =   0
            Picture         =   "VBTERM.frx":030A
            Stretch         =   -1  'True
            ToolTipText     =   "Toggles Port"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgNotConnected 
            Height          =   240
            Left            =   0
            Picture         =   "VBTERM.frx":0454
            Stretch         =   -1  'True
            ToolTipText     =   "Toggles Port"
            Top             =   0
            Width           =   240
         End
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   210
      Top             =   3645
   End
   Begin VB.TextBox txtTerm 
      Height          =   3690
      Left            =   1245
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1140
      Width           =   5790
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   165
      Top             =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   45
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin MSComDlg.CommonDialog OpenLog 
      Left            =   105
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "LOG"
      DialogTitle     =   "Apri file registro per le comunicazioni"
      Filter          =   "File registro (*.log)|*.log;"
      FilterIndex     =   501
   End
   Begin ComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4620
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Stato:"
            TextSave        =   "Stato:"
            Key             =   "Status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Stato porta comunicazioni"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8389
            MinWidth        =   2
            Text            =   "Impostazioni:"
            TextSave        =   "Impostazioni:"
            Key             =   "Settings"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Impostazioni porta comunicazioni"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1244
            TextSave        =   ""
            Key             =   "ConnectTime"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Tempo di connessione"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VBTERM.frx":059E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VBTERM.frx":08B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VBTERM.frx":0BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VBTERM.frx":0EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VBTERM.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VBTERM.frx":1520
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenLog 
         Caption         =   "&Apri file registro..."
      End
      Begin VB.Menu mnuCloseLog 
         Caption         =   "&Chiudi file registro"
         Enabled         =   0   'False
      End
      Begin VB.Menu M3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendText 
         Caption         =   "&Trasmetti file di testo..."
         Enabled         =   0   'False
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu mnuPort 
      Caption         =   "&Porta comunicazioni"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Apri porta"
      End
   End
   Begin VB.Menu mnuMSComm 
      Caption         =   "&MSComm"
      Begin VB.Menu mnuInputLen 
         Caption         =   "&InputLen..."
      End
      Begin VB.Menu mnuRThreshold 
         Caption         =   "&RThreshold..."
      End
      Begin VB.Menu mnuSThreshold 
         Caption         =   "&SThreshold..."
      End
      Begin VB.Menu mnuParRep 
         Caption         =   "P&arityReplace..."
      End
      Begin VB.Menu mnuDTREnable 
         Caption         =   "&DTREnable"
      End
      Begin VB.Menu Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHCD 
         Caption         =   "&CDHolding..."
      End
      Begin VB.Menu mnuHCTS 
         Caption         =   "CTSH&olding..."
      End
      Begin VB.Menu mnuHDSR 
         Caption         =   "DSRHo&lding..."
      End
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------
' VBTerm - Programma dimostrativo per il
' controllo ActiveX MSComm.
'
' Copyright (c) 1994, Crescent Software, Inc.
' sviluppato da Don Malin e Carl Franklin.
'
' Aggiornato da Mike Maddox
'--------------------------------------------------
Option Explicit
                        
Dim Temp As String      ' Stringa temporanea.

Private Sub Form_Load()
    Dim CommPort As String, Handshaking As String, Settings As String
    Dim OpenFlag As Boolean
    On Error Resume Next
    
    isTerminal = True
    
    ' Imposta il colore predefinito per il terminale
    txtTerm.SelLength = Len(txtTerm)
    txtTerm.SelText = ""
    txtTerm.ForeColor = vbBlue
       
    ' Imposta il titolo
    'App.Title = "Terminale Visual Basic"
    
    ' Imposta la spia per l'indicazione dello stato
    imgNotConnected.ZOrder
       
    ' Centra il form
    frmTerminal.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    ' Carica le impostazioni del registro di configurazione
    
'    Settings = GetSetting(App.Title, "Properties", "Settings", "") ' fmain.MSComm1.Settings]\
'    If Settings <> "" Then
'        MSComm1.Settings = Settings
'        If Err Then
'            MsgBox Error$, 48
'            Exit Sub
'        End If
'    End If
    
    'CommPort = GetSetting(App.Title, "Properties", "CommPort", "") ' fmain.MSComm1.CommPort
    CommPort = fMain.MSComm1.CommPort
    'If CommPort <> "" Then fmain.MSComm1.CommPort = CommPort 'La porta è già aperta!
    
    Handshaking = fMain.MSComm1.Handshaking
'    Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "") 'fmain.MSComm1.Handshaking
'    If Handshaking <> "" Then
'        fmain.MSComm1.Handshaking = Handshaking
'        If Err Then
'            MsgBox Error$, 48
'            Exit Sub
'        End If
'    End If
    
'    Echo = GetSetting(App.Title, "Properties", "Echo", "") ' Eco
    Echo = True
    On Error GoTo 0
    OpenFlag = fMain.MSComm1.PortOpen

    mnuOpen.Checked = OpenFlag
    mnuSendText.Enabled = OpenFlag
    tbrToolBar.Buttons("TransmitTextFile").Enabled = OpenFlag

    imgConnected.ZOrder
    sbrStatus.Panels("Settings").Text = "Impostazioni: " & fMain.MSComm1.Settings
    'fMain.MSComm1.CommEvent
    fMain.MSComm1.RThreshold = 1
    fMain.MSComm1.SThreshold = 1
    StartTiming

End Sub

Private Sub Form_Resize()
   ' Ridimensiona la finestra del terminale VB
   txtTerm.Move 0, tbrToolBar.Height, frmTerminal.ScaleWidth, frmTerminal.ScaleHeight - sbrStatus.Height - tbrToolBar.Height
   
   ' Posiziona la spia per l'indicazione dello stato
   Frame1.Left = ScaleWidth - Frame1.Width * 1.5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Counter As Long

'    If fmain.MSComm1.PortOpen Then
'       ' Attende 10 secondi per la trasmissione dei dati.
'       Counter = Timer + 10
'       Do While fmain.MSComm1.OutBufferCount
'          Ret = DoEvents()
'          If Timer > Counter Then
'             Select Case MsgBox("Impossibile trasmettere i dati", 34)
'                ' Annulla.
'                Case 3
'                   Cancel = True
'                   Exit Sub
'                ' Riprova.
'                Case 4
'                   Counter = Timer + 10
'                ' Ignora.
'                Case 5
'                   Exit Do
'             End Select
'          End If
'       Loop
'
'       fmain.MSComm1.PortOpen = 0
'    End If
'
'    ' Se il file registro è aperto, ne cancella il contenuto e lo chiude.
'    If hLogFile Then mnuCloseLog_Click
'    End
    fMain.MSComm1.RThreshold = 0
    fMain.MSComm1.SThreshold = 0
    isTerminal = False
    Unload Me
    fMain.Show
End Sub

Private Sub imgConnected_Click()
    ' Richiama la routine mnuOpen_Click per passare dalla modalità di
    ' connessione a quella di disconnessione e viceversa.
    Call mnuOpen_Click
End Sub

Private Sub imgNotConnected_Click()
    ' Richiama la routine mnuOpen_Click per passare dalla modalità di
    ' connessione a quella di disconnessione e viceversa.
    Call mnuOpen_Click
End Sub

Private Sub mnuCloseLog_Click()
    ' Chiude il file registro.
    Close hLogFile
    hLogFile = 0
    mnuOpenLog.Enabled = True
    tbrToolBar.Buttons("OpenLogFile").Enabled = True
    mnuCloseLog.Enabled = False
    tbrToolBar.Buttons("CloseLogFile").Enabled = False
    frmTerminal.Caption = "Terminale"
End Sub


' Inverte l'impostazione della proprietà DTREnabled.
Private Sub mnuDTREnable_Click()
    ' Inverte l'impostazione della proprietà DTREnabled
    fMain.MSComm1.DTREnable = Not fMain.MSComm1.DTREnable
    mnuDTREnable.Checked = fMain.MSComm1.DTREnable
End Sub


Private Sub mnuFileExit_Click()
    ' Utilizza Form_Unload in quanto contiene codice per verificare se dei dati
    ' non sono stati inviati e se esiste un file registro aperto.
    Form_Unload Ret
    isTerminal = False
End Sub



' Inverte l'impostazione della proprietà DTREnable per interrompere la comunicazione.
Private Sub mnuHangup_Click()
    On Error Resume Next
    
    fMain.MSComm1.Output = "ATH"      ' Invia la stringa di interruzione
    Ret = fMain.MSComm1.DTREnable     ' Salva l'impostazione corrente.
    fMain.MSComm1.DTREnable = True    ' Imposta la proprietà DTREnable su True.
    fMain.MSComm1.DTREnable = False   ' Imposta la proprietà DTREnable su False.
    fMain.MSComm1.DTREnable = Ret     ' Ripristina le impostazioni precedenti.
    'mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    
    ' Se la porta è ancora aperta, la chiude
    'If fmain.MSComm1.PortOpen Then fmain.MSComm1.PortOpen = False
    
    ' Notifica all'utente che è stato generato un errore
    If Err Then MsgBox Error$, 48
    
    mnuSendText.Enabled = False
    tbrToolBar.Buttons("TransmitTextFile").Enabled = False
    'mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    'mnuDial.Enabled = True
    tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
    sbrStatus.Panels("Settings").Text = "Impostazioni: "
    
    ' Disattiva la spia e nasconde il segno di spunta a sinistra
    ' della voce di menu Apri porta
    mnuOpen.Checked = False
    imgNotConnected.ZOrder
            
    ' Interrompe il timer della porta
    StopTiming
    sbrStatus.Panels("Status").Text = "Stato: "
    On Error GoTo 0
End Sub

' Visualizza il valore della proprietà CDHolding.
Private Sub mnuHCD_Click()
    If fMain.MSComm1.CDHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CDHolding = " + Temp
End Sub

' Visualizza il valore della proprietà CTSHolding.
Private Sub mnuHCTS_Click()
    If fMain.MSComm1.CTSHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CTSHolding = " + Temp
End Sub

' Visualizza il valore della proprietà DSRHolding.
Private Sub mnuHDSR_Click()
    If fMain.MSComm1.DSRHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "DSRHolding = " + Temp
End Sub

' Questa routine imposta la propretà InputLen, che determina quanti
' byte di dati vengono letti ogni volta che Input viene utilizzata
' per recuperare dati dal buffer di input.
' L'impostazione di InputLen su 0 specifica che
' deve essere letto l'intero contenuto del buffer.
Private Sub mnuInputLen_Click()
    On Error Resume Next

    Temp = InputBox$("Immettere il nuovo valore per InputLen:", "InputLen", Str$(fMain.MSComm1.InputLen))
    If Len(Temp) Then
        fMain.MSComm1.InputLen = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub

' Inverte lo stato della porta (aperta o chiusa).
Private Sub mnuOpen_Click()
    On Error Resume Next
    Dim OpenFlag
'
'    fmain.MSComm1.PortOpen = Not fmain.MSComm1.PortOpen
'    If Err Then MsgBox Error$, 48
'

'    CloseCom
'    fMain.MSComm1.Settings = "19200,n,8,1"
'    fMain.MSComm1.CommPort = "1"
'    fMain.MSComm1.Handshaking = comRTS
'    fMain.MSComm1.RTSEnable = True
    OpenCom
    isTerminal = True
    OpenFlag = fMain.MSComm1.PortOpen
'
    mnuOpen.Checked = OpenFlag
    mnuSendText.Enabled = OpenFlag
    tbrToolBar.Buttons("TransmitTextFile").Enabled = OpenFlag
'
'    If fmain.MSComm1.PortOpen Then
'        ' Attiva il pulsante e la voce di menu per la composizione del numero
'        mnuDial.Enabled = True
'        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
'
'        ' Attiva il pulsante e la voce di menu per l'interruzione della comunicazione
'        mnuHangUp.Enabled = True
'        tbrToolBar.Buttons("HangUpPhone").Enabled = True
'
        imgConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Impostazioni: " & fMain.MSComm1.Settings
        StartTiming
'    Else
'        ' Attiva il pulsante e la voce di menu per la composizione del numero
'        mnuDial.Enabled = True
'        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
'
'        ' Disattiva il pulsante e la voce di menu per l'interruzione della comunicazione
'        mnuHangUp.Enabled = False
'        tbrToolBar.Buttons("HangUpPhone").Enabled = False
'
'        imgNotConnected.ZOrder
'        sbrStatus.Panels("Settings").Text = "Impostazioni: "
'        StopTiming
'    End If
'
End Sub

Private Sub mnuOpenLog_Click()
   Dim Replace
   On Error Resume Next
   OpenLog.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer
   OpenLog.CancelError = True
      
   ' Richiede all'utente il nome del file registro.
   OpenLog.DialogTitle = "Apri file registro per le comunicazioni"
   OpenLog.Filter = "File registro (*.LOG)|*.log|Tutti i file (*.*)|*.*"
   
   Do
      OpenLog.FileName = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then Exit Sub
      Temp = OpenLog.FileName

      ' Se il file esiste già, chiede all'utente se desidera sovrascrivere
      ' il file esistente o aggiungervi il nuovo file.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
      If Ret Then
         Replace = MsgBox("Sostituire il file esistente - " + Temp + "?", 35)
      Else
         Replace = 0
      End If
   Loop While Replace = 2

   ' L'utente ha scelto il pulsante Sì, quindi il file viene eliminato.
   If Replace = 6 Then
      Kill Temp
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
   End If

   ' Apre il file registro.
   hLogFile = FreeFile
   Open Temp For Binary Access Write As hLogFile
   If Err Then
      MsgBox Error$, 48
      Close hLogFile
      hLogFile = 0
      Exit Sub
   Else
      ' Passa alla fine del file in modo che sia possibile aggiungervi i nuovi dati.
      Seek hLogFile, LOF(hLogFile) + 1
   End If

   frmTerminal.Caption = "Terminale - " + OpenLog.FileTitle
   mnuOpenLog.Enabled = False
   tbrToolBar.Buttons("OpenLogFile").Enabled = False
   mnuCloseLog.Enabled = True
   tbrToolBar.Buttons("CloseLogFile").Enabled = True
End Sub

' Questa routine imposta la proprietà ParityReplace, che contiene
' il carattere che sostituirà tutti i caratteri non corretti
' che vengono ricevuti a causa di un errore di parità.
Private Sub mnuParRep_Click()
    On Error Resume Next

    Temp = InputBox$("Immettere il carattere sostitutivo", "ParityReplace", fMain.MSComm1.ParityReplace)
    fMain.MSComm1.ParityReplace = Left$(Temp, 1)
    If Err Then MsgBox Error$, 48
End Sub

' Questa routine imposta la proprietà RThreshold, che determina
' quanti byte possono arrivare al buffer di ricezione prima che
' l'evento OnComm venga generato e la proprietà CommEvent venga
' impostata su comEvReceive.
Private Sub mnuRThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Immettere il nuovo valore per la proprietà RThreshold:", "RThreshold", Str$(fMain.MSComm1.RThreshold))
    If Len(Temp) Then
        fMain.MSComm1.RThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If

End Sub

Private Sub mnuSendText_Click()
   Dim hSend, BSize, LF&
   
   Exit Sub
   
   On Error Resume Next
   
   mnuSendText.Enabled = False
   tbrToolBar.Buttons("TransmitTextFile").Enabled = False
   
   ' Richiede all'utente il nome del file.
   OpenLog.DialogTitle = "Invia file di testo"
   OpenLog.Filter = "File di testo (*.TXT)|*.txt|Tutti i file (*.*)|*.*"
   Do
      OpenLog.CancelError = True
      OpenLog.FileName = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then
        mnuSendText.Enabled = True
        tbrToolBar.Buttons("TransmitTextFile").Enabled = True
        Exit Sub
      End If
      Temp = OpenLog.FileName

      ' Se il file non esiste torna indietro.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         mnuSendText.Enabled = True
         tbrToolBar.Buttons("TransmitTextFile").Enabled = True
         Exit Sub
      End If
      If Ret Then
         Exit Do
      Else
         MsgBox Temp + " non è stato trovato", 48
      End If
   Loop

   ' Apre il file registro.
   hSend = FreeFile
   Open Temp For Binary Access Read As hSend
   If Err Then
      MsgBox Error$, 48
   Else
      ' Visualizza la finestra di dialogo di annullamento dell'operazione.
      CancelSend = False
      Me.Caption = "Trasmissione in corso del file di testo - " + Temp
      'frmCancelSend.Show
      
      ' Legge il file in blocchi pari alle dimensioni del buffer di trasmissione.
      BSize = fMain.MSComm1.OutBufferSize
      LF& = LOF(hSend)
      Do Until EOF(hSend) Or CancelSend
         ' Non legge più del necessario alla fine.
         If LF& - Loc(hSend) <= BSize Then
            BSize = LF& - Loc(hSend) + 1
         End If
      
         ' Legge un blocco di dati.
         Temp = Space$(BSize)
         Get hSend, , Temp
      
         ' Trasmette il blocco.
         fMain.MSComm1.Output = Temp
         If Err Then
            MsgBox Error$, 48
            Exit Do
         End If
      
         ' Attende che tutti i dati vengano inviati.
         Do
            Ret = DoEvents()
         Loop Until fMain.MSComm1.OutBufferCount = 0 Or CancelSend
      Loop
   End If
   
   Close hSend
   mnuSendText.Enabled = True
   tbrToolBar.Buttons("TransmitTextFile").Enabled = True
   CancelSend = True
   'frmCancelSend.Hide
End Sub


' Questa routine imposta la proprietà  SThreshold, che determina
' quanti caratteri (al massimo) debbono essere attesi
' nel buffer di uscita  prima che la proprietà CommEvent
' venga impostata su comEvSend e che l'evento OnComm venga generato.
Private Sub mnuSThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Immettere il nuovo valore per SThreshold", "SThreshold", Str$(fMain.MSComm1.SThreshold))
    If Len(Temp) Then
        fMain.MSComm1.SThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub

Private Sub Timer2_Timer()
sbrStatus.Panels("Status").Text = "Stato: "
Timer2.Enabled = False

End Sub

' I tasti intercettati in questa fase vengono inviati
' al controllo MSComm dove vengono ripetuti tramite
' l'evento OnComm (comEvReceive) e visualizzati
' con la routine ShowData.
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    ' Se la porta è aperta...
    If fMain.MSComm1.PortOpen Then
        ' Invia alla porta le sequenze di tasti.
        fMain.MSComm1.Output = Chr$(KeyAscii)
        
        ' Non è necessario visualizzare il tasto nel
        ' controllo a meno che l'eco non sia attivato.
        ' I modem in genere ripetono il carattere.
        If Not Echo Then
            ' Si posiziona alla fine del terminale
            txtTerm.SelStart = Len(txtTerm)
            KeyAscii = 0
        End If
    End If
     
End Sub


Private Sub tbrToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "OpenLogFile"
    Call mnuOpenLog_Click
Case "CloseLogFile"
    Call mnuCloseLog_Click
Case "DialPhoneNumber"
    'Call mnuDial_Click
Case "HangUpPhone"
    Call mnuHangup_Click
'Case "Properties"
'    Call mnuProperties_Click
Case "TransmitTextFile"
    Call mnuSendText_Click
End Select
End Sub

Private Sub Timer1_Timer()
    ' Visualizza la durata della connessione
    sbrStatus.Panels("ConnectTime").Text = Format(Now - StartTime, "hh:nn:ss") & " "
End Sub
