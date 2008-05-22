Attribute VB_Name = "Init"
Option Explicit

Private Sub Main()
    Dim i As Integer
  'show the splash screen
   frmSplash.Show
   'Execute Init instructions
   Init
   DoEvents
   'Call Sleep(2000)
  'show the main application
   fMain.Show
   DoEvents
  'perform any other startup functions as required by your program
    i = Val(sReadINI("Modem", "UltimaCom", FileIni))
    If i = 0 Then
        'There is not a COM port defined
        fMain.bModem.Enabled = False
    End If

  'unload the splash screen and free its memory
   Unload frmSplash
   Set frmSplash = Nothing
End Sub

Public Sub Init()
    Dim nfile As Integer
    Dim rint As Integer
    Dim Path As String
    Dim i As Long
    
    Path = sGetAppPath()

    Versione = "0.0.1"

    FileIni = sGetAppPath + "PegasoPrimo.ini"

    
    
    SE = ";"    'Il separatore di elenco è la virgola
    frmSplash.lblWarning = ""

    'Riempimento tabella SensorType
    SensorType(0) = "Empty0"
    SensorType(1) = "Empty1"
    SensorType(2) = "Empty2"
    SensorType(3) = "First Methane Sensor"
    SensorType(4) = "Second Methane Sensor"
    SensorType(5) = "Third Methane Sensor"
    SensorType(6) = "H2S Sensor"
    SensorType(7) = "CTD"

    CTRLC = Chr(3)
    fdn = 0
    'Apre il file di log
    If fDebug Then
        filename = sGetAppPath + "log.txt"
        fdn = FreeFile
        Open filename For Append As #fdn
        Print #fdn,
        Print #fdn, "-----------------------------------------------------"
        Print #fdn, Versione
        Print #fdn, Date, Time
    End If

End Sub

