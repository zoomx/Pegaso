Attribute VB_Name = "VariabiliGlobali"
Option Explicit

Public Const TmOut As Integer = 10 'Timeout comunicazioni
Public Const None As Integer = 0
Public ComPort As Integer
Public ModemString As String
Public PAnno As String
Public PMese As String
Public PGiorno As String
Public POra As String
Public PMinuti As String
Public Programmato As Boolean
Public Scaricato As Boolean    'Serve per sapere se si e' scaricato o meno. Deve essere globale
Public ProgLoaded As Boolean
Public ProgChanged As Boolean
Public ProgSaved As Boolean
Public Messaggio As String
'Public Intero As Integer
Public Lungo As Long
'Public Float As Single
'Public Dfloat As Double
'Public Stringa As String
Public ComOk As Boolean
Public Collegato As Boolean
Public FileOut As String
Public PathOut As String
Public DriveOut As String
Public comando As String
Public Esci As Boolean
Public DataProgrammazione As Date

Public filename As String
Public Const Vero As Boolean = True
Public Const Falso As Boolean = False
Public Stazione As String
Public Intervallo As Long  'Intevallo di campionamento in secondi
Public CTRLC As String
Public ConnessioneRemota As Boolean

Public FileIni As String    'Nome file di inizializzazione
Public SE As String         'Separatore di elenco
Public Decimale As String
Public fDebug As Boolean    'Se e' vero stampa sul file di log
Public fdn As Integer       'E' il numero di file del file di log
Public lDebug As Boolean    'Se è vero fa comparire piu' pulsanti e menu speciali
                            'visibili in precedenti versioni e adesso nascosti
Public TipoFile As String   'Indica il tipo di file da salvare ASCII Binario Excel..
Public InitDirData As String    'Indica la dir iniziale per salvare i dati
Public LastFileSaved As String  'Ultimo file salvato
Public GMTshift As Integer  'indica lo shift rispetto all'orario GMT

Public Versione As String
Public ChiamaFlag As Boolean


'MM commands
Public Const WakeUp As String = "W"
Public Const Directory As String = "DD"
Public Const GetFile As String = "GF"
Public Const GetDateTime As String = ""
Public Const SetDateTime As String = ""

Public SensorType(7) As String
