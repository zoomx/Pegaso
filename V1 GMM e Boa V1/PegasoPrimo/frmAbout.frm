VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informazioni su Poseidon"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   33712
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   3075
      Width           =   1245
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   360
      Picture         =   "frmAbout.frx":110C
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5160
      Picture         =   "frmAbout.frx":1D4E
      Top             =   1800
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "Descrizione applicazione"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Pegaso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versione "
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Avviso: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   4
      Top             =   2625
      Visible         =   0   'False
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opzioni di protezione per la chiave del registro di configurazione...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipi di primo livello per la chiave del registro di configurazione...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Stringa Unicode con terminazione Null
Const REG_DWORD = 4                      ' Numero a 32 bit

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef pHkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub CmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Informazioni su " & App.Title
    lblVersion.Caption = "Versione " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription = App.FileDescription
End Sub

Public Sub StartSysInfo()
    On Local Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Tenta di leggere dal registro di configurazione le informazioni del sistema
    ' sul nome e il percorso dell'applicazione...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Tenta di leggere dal registro di configurazione le informazioni del sistema
    ' relative solo al percorso dell'applicazione...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Convalida l'esistenza di una versione del file a 32 bit conosciuta
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Errore - Impossibile trovare il file...
        Else
            GoTo SysInfoErr
        End If
    ' Errore - Impossibile trovare la voce del registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Le informazioni sul sistema non sono disponibili in questa fase", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contatore per il ciclo
    Dim rc As Long                                          ' Codice restituito
    Dim hKey As Long                                        ' Handle a una chiave del registro di configurazione aperta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo di dati di una chiave del registro di configurazione
    Dim tmpVal As String                                    ' Variabile per la memorizzazione temporanea del valore di una chiave del registro di configurazione
    Dim KeyValSize As Long                                  ' Dimensioni della variabile per la chiave del registro di configurazione
    '------------------------------------------------------------------
    ' Apre la chiave del registro sotto KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Apre la chiave del registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestisce gli errori...
    
    tmpVal = String$(1024, 0)                             ' Assegna lo spazio per la variabile
    KeyValSize = 1024                                       ' Definisce le dimensioni della variabile
    
    '---------------------------------------------------------------
    ' Recupera il valore della chiave del registro di configurazione...
    '---------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Recupera/crea il valore della chiave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestisce gli errori
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 aggiunge una stringa con terminazione Null...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Trova Null, estrae dalla stringa
    Else                                                    ' WinNT non aggiunge la terminazione Null alle stringhe...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Non trova Null, estrae solo la stringa
    End If
    '----------------------------------------------------------------
    ' Determina il tipo del valore della chiave per la conversione...
    '----------------------------------------------------------------
    Select Case KeyValType                                  ' Esamina i tipi di dati...
    Case REG_SZ                                             ' Tipo di dati String per la chiave del registro
        KeyVal = tmpVal                                     ' Copia il valore String
    Case REG_DWORD                                          ' Tipo di dati Double Word per la chiave del registro
        For i = Len(tmpVal) To 1 Step -1                    ' Converte ogni bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Crea il valore carattere per carattere.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Converte Double Word in String
    End Select
    
    GetKeyValue = True                                      ' Operazione riuscita
    rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro
    Exit Function                                           ' Esce
    
GetKeyError:      ' Svuota in seguito a un errore...
    KeyVal = ""                                             ' Imposta su una stringa vuota il valore restituito
    GetKeyValue = False                                     ' Operazione non riuscita
    rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro
End Function

Private Sub mnuWeb_Click()
    Dim ThisApp As String
    Dim ThisVer As String
    ThisApp = App.ProductName '"Title of App"
    ThisVer = App.Major + App.Minor '"Version of App"
    Dim iRet As Long
    Exit Sub
    Dim response As Integer
    response = MsgBox("Hai scelto 'Visita il sito Web', che" & _
    " lancerà il tuo Browser e ti collegherà al sito Web SIMA." & vbCrLf & "" & _
    vbCrLf & "Vuoi continuare?", 4, ThisApp & ThisVer)

    Select Case response

    ' Yes response.
      Case vbYes:
        iRet = ShellExecute(Me.hwnd, vbNullString, _
        "http://www.geocities.com/", vbNullString, _
        "c:\", SW_SHOWNORMAL)
    ' No response.
      Case vbNo:
        Exit Sub
    End Select
End Sub

Private Sub iSimaLogo_Click()
    mnuWeb_Click
End Sub

