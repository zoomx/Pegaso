Attribute VB_Name = "FunzioniBoa"
  Option Explicit
  
  ' Datalog Download...
  Const FILE_LAST_REQ = 15                        ' 0X0F
  Const FILE_LAST_RESP = 16                       ' 0X10
  Const FILE_INFO_REQ = 17                        ' 0X11
  Const FILE_INFO_RESP = 18                       ' 0X12
  Const FILE_SYNCHRO_REQ = 19                     ' 0X13
  Const FILE_SYNCHRO_RESP = 20                    ' 0X14
  Const FILE_DOWNLOAD_REQ = 21                    ' 0X15
  Const FILE_DOWNLOAD_RESP = 22                   ' 0X16
  Const DATALOG_ENDIS_REQ = 23                    ' 0X17
  Const DATALOG_ENDIS_RESP = 24                   ' 0X18
  Const DATALOG_SETFREQ_REQ = 25                  ' 0X19
  Const DATALOG_SETFREQ_RESP = 26                 ' 0X1A
  Const DATALOG_STATE_REQ = 27                    ' 0X1B
  Const DATALOG_STATE_RESP = 28                   ' 0X1C
  
  Const SOF = 1                                  ' start of frame
  Const EOF = 0                                  ' end of frame

Type TDaylyFileInfo
    File_ID As Double  'word
    File_Day As Byte
    File_Month As Byte
    File_Year As Byte
    File_Hour As Byte
    File_Min As Byte
End Type

Type TDatalogFields
    DL_Day As Byte
    DL_Month As Byte
    DL_Year As Byte
    DL_Hour As Byte
    DL_Min As Byte
    DL_Sec As Byte

    '....Temperature
    DL_T1 As Double         'SmallInt T1 (signed int16 moltiplicato x 10)
    DL_T2 As Double         'SmallInt T2 (signed int16 moltiplicato x 10)
    DL_T3 As Double         'SmallInt T3 (signed int16 moltiplicato x 10)

    '....Meteo
    DL_MeteoWindSpeed As Double     'Word (unsigned int16 moltiplicato x 10)
    DL_MeteoWindDirection As Double 'Word (unsigned int16 moltiplicato x 10)

    '....GPS
    DL_GPS_Boa_Lat As Double    'integer (unsigned int32 moltiplicato x 10^7)
    DL_GPS_Boa_Lon As Double    'integer (unsigned int32 moltiplicato x 10^7)
    DL_GPS_Boa_Lat_Dir As Byte  'Char
    DL_GPS_Boa_Lon_Dir As Byte  'Char
    DL_GPSSatUsed As Byte       'byte

    '....POWER
    DL_MonitorBattery_1_12V As Byte             '  (unsigned int8 molt x 10)
    DL_MonitorBattery_2_12V As Byte             '  (unsigned int8 molt x 10)
    DL_MonitorBattery_3_12V As Byte             '  (unsigned int8 molt x 10)

    '....Accelerometers
    DL_AX As Integer     'SmallInt T1 (signed int16 moltiplicato x 10)
    DL_AY As Integer     'SmallInt T2 (signed int16 moltiplicato x 10)
    DL_AZ As Integer     'SmallInt T3 (signed int16 moltiplicato x 10)

    '....H2O SENSORS
    DL_H2O_1 As Byte             '  (unsigned int8 molt x 10)
    DL_H2O_2 As Byte             '  (unsigned int8 molt x 10)
End Type

'var
' Last Closed File INFO
    LCF_ID As Double    'Word = 0
    LCF_Year As Byte    'byte = 0
    LCF_Month As Byte   'byte = 0
    LCF_Day As Byte     'byte = 0
    LCF_Hour As Byte    'byte = 0
    LCF_Minute As Byte  'byte = 0

' File Info Req & download
    Selected_File_ID As Double          'Word = 0
    SelectedFileLinesNumber As Double   'Word = 0

' Datalog Enable/Disable
    Datalog_Enable As Byte  'byte = 3
    
'Datalog Sample time (minute)
    DatalogSampleTime As Byte   'byte = 0

    Dim FilesNameArray(30) As String      'array(1..30) of string

'Datalog Synchronization
    Dim FileSynchroVector(30) As TDaylyFileInfo     'array(1..30) of TDaylyFileInfo

'Download
    LinesToDownloadNumber As Byte   'byte = 4
    LineTostartDownload As Double   'Word = 0
    File_Line_Downloaded As Double  'Word = 0

  'alla freq massima di una scrittura al minuto posso aver scritto su SD 60*24 = 1440 righe.
    Dim DataDownloadedVector(1440) As TDatalogFields    'array(0..1440) of TDatalogFields
  'Sapendo che ogni riga occupa 68 byte ne risulta una occupazione di memoria di 10183680 byte in delphi

    prova As Byte 'byte = 0
    Namefile As String
    text As String
    BuoyFile As Variant   'TextFile
    i As Integer    'Integer


'==============================================================================
'  SendLastFileRequest()
'
'    Richiedo al Datalogger qual'è l'ultimo file chiuso ed aggiornato
'    Il datalogger mi risponde fornendomi:
'    - ID del file
'    - Day, Month, Hour, Min, Sec  della chiusura del file
'
'==============================================================================
'Procedure SendLastFileRequest()
'begin
'
'  Form1.DATA_REQUEST = False
'  Form1.SendRequest (FILE_LAST_REQ)
'
'End
'******************************************************************************




'Procedure ReceiveLastFile()
'Var
  DataTmp As String
'begin

  '| 2 | 3 | 4 | 5 |         (4 sent byte per l'ID dell'ultimo file chiuso)
  DataTmp = Form1.Buffer(3) + Form1.Buffer(2) + Form1.Buffer(5) + Form1.Buffer(4)
  HexToBin(PChar(DataTmp), @LCF_ID, 4)

  '| 14 | 15 |                (2 sent byte per Last Closed File Minute)
  DataTmp = Form1.Buffer(7) + Form1.Buffer(6)
  HexToBin(PChar(DataTmp), @LCF_Minute, 2)

  '| 12 | 13 |                (2 sent byte per Last Closed File Hour)
  DataTmp = Form1.Buffer(9) + Form1.Buffer(8)
  HexToBin(PChar(DataTmp), @LCF_Hour, 2)

  '| 10 | 11 |                (2 sent byte per Last Closed File Day)
  DataTmp = Form1.Buffer(11) + Form1.Buffer(10)
  HexToBin(PChar(DataTmp), @LCF_Year, 2)

  '| 8 | 9 |                  (2 sent byte per Last Closed File Month)
  DataTmp = Form1.Buffer(13) + Form1.Buffer(12)
  HexToBin(PChar(DataTmp), @LCF_Month, 2)

  '| 6 | 7 |                  (2 sent byte per Last Closed File Year)
  DataTmp = Form1.Buffer(15) + Form1.Buffer(14)
  HexToBin(PChar(DataTmp), @LCF_Day, 2)
'End
'******************************************************************************



'==============================================================================
'  SendFileSynchroRequest()
'
'    Richiedo al Datalogger l'intero file di sincronizzazione
'    Il datalogger mi risponde fornendomi:
'    - ID di ogni file in memoria
'    - Day, Month, Hour, Min, Sec  di ciascun file
'
'==============================================================================
'Procedure SendFileSynchroRequest()
'begin
'
'  Form1.DATA_REQUEST = False
'  sleep (500)
'  Form1.SendRequest (FILE_SYNCHRO_REQ)
'
'End




'==============================================================================
'  SendFileInfoRequest()
'
'    Richiedo le informazioni di un file passando come parametro l'ID
'    Ricevo come risultato l'ID del file ed il numero di line
'    di cui è costituito
'
'==============================================================================
'procedure SendFileInfoRequest(File_ID: Word)
'Var
  Checksum As Byte
  s As String

'begin

  Form1.DATA_REQUEST = False
  'sleep (500)
  Form1.CommPortDriver1.Sendbyte (SOF)
  Form1.CommPortDriver1.Sendbyte (FILE_INFO_REQ)
  Checksum = SOF + FILE_INFO_REQ


  s = inttohex(File_ID, 4)
  Form1.CommPortDriver1.Sendchar (s(4))
  Form1.CommPortDriver1.Sendchar (s(3))
  Form1.CommPortDriver1.Sendchar (s(2))
  Form1.CommPortDriver1.Sendchar (s(1))
  Checksum = Checksum + byte(s(1)) + byte(s(2))+ byte(s(3)) + byte(s(4))

  Checksum = 256 - (Checksum Mod 256)

  Form1.CommPortDriver1.Sendbyte (Checksum)
  If Checksum <> 0 Then
    Form1.CommPortDriver1.Sendbyte (EOF)
  Else
    exit

'End





'==============================================================================
'  SendPacketDownloadRequest()
'
'    Sapendo quante linee contiene il file che voglio scaricare
'    suddivido il download in tanti pacchetti.
'    Ogni volta che ricevo un pacchetto richiedo il successivo, fino
'    alla fine del file.
'
'==============================================================================
'procedure SendPacketDownloadRequest(File_ID: Word RowsNumber: byte RowIndex: Word)
'Var
  Checksum As Byte
  s As String
'begin

  inc (prova)
  Form1.edit8.text = inttostr(prova)
  

  Form1.CommPortDriver1.Sendbyte (SOF)
  Form1.CommPortDriver1.Sendbyte (FILE_DOWNLOAD_REQ)
  Checksum = SOF + FILE_DOWNLOAD_REQ

  s = inttohex(Selected_File_ID, 4)
  Form1.CommPortDriver1.Sendchar (s(4))
  Form1.CommPortDriver1.Sendchar (s(3))
  Form1.CommPortDriver1.Sendchar (s(2))
  Form1.CommPortDriver1.Sendchar (s(1))
  Checksum = Checksum + byte(s(1)) + byte(s(2))+ byte(s(3)) + byte(s(4))

  s = inttohex(RowsNumber, 2)
  Form1.CommPortDriver1.Sendchar (s(2))
  Form1.CommPortDriver1.Sendchar (s(1))
  Checksum = Checksum + byte(s(1)) + byte(s(2))

  s = inttohex(RowIndex, 4)
  Form1.CommPortDriver1.Sendchar (s(4))
  Form1.CommPortDriver1.Sendchar (s(3))
  Form1.CommPortDriver1.Sendchar (s(2))
  Form1.CommPortDriver1.Sendchar (s(1))
  Checksum = Checksum + byte(s(1)) + byte(s(2))+ byte(s(3)) + byte(s(4))

  Checksum = 256 - (Checksum Mod 256)

  Form1.CommPortDriver1.Sendbyte (Checksum)
  If Checksum <> 0 Then
    Form1.CommPortDriver1.Sendbyte (EOF)
  Else
    exit

'End





'Procedure ReceiveFileInfo()
'Var
  DataTmp As String
'begin

  '| 2 | 3 | 4 | 5 |             (4 sent byte File ID)
  DataTmp = Form1.Buffer(3) + Form1.Buffer(2) + Form1.Buffer(5) + Form1.Buffer(4)
  HexToBin(PChar(DataTmp), @Selected_File_ID, 4)

  DataTmp = Form1.Buffer(7) + Form1.Buffer(6) + Form1.Buffer(9) + Form1.Buffer(8)
  HexToBin(PChar(DataTmp), @SelectedFileLinesNumber, 4)
'End





'Procedure ReceiveFileSynchro()
'Var
  i As Byte
  shift As Integer  'Integer
  DataTmp As String
'begin

'  Form1.Memo1.Clear

  For i = 1 To 30 'do
  'begin

    shift = (i - 1) * 14

    '| 2 | 3 | 4 | 5 |         (4 sent byte File ID)
    DataTmp = Form1.Buffer(3 + shift) + Form1.Buffer(2 + shift) + Form1.Buffer(5 + shift) + Form1.Buffer(4 + shift)
    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_ID, 4)

    '| 6 | 7 |                     (2 sent byte per la File Minute)
    DataTmp = Form1.Buffer(7 + shift) + Form1.Buffer(6 + shift)
    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Day, 2)

    '| 8 | 9 |                     (2 sent byte per la File Hour)
    DataTmp = Form1.Buffer(9 + shift) + Form1.Buffer(8 + shift)
    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Month, 2)

    '| 10 | 11 |                     (2 sent byte per la File Day)
    DataTmp = Form1.Buffer(11 + shift) + Form1.Buffer(10 + shift)
    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Year, 2)

    '| 12 | 13 |                     (2 sent byte per la File Month)
    DataTmp = Form1.Buffer(13 + shift) + Form1.Buffer(12 + shift)
    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Hour, 2)

    '| 14 | 15 |                   (2 sent byte per la File Year)
    DataTmp = Form1.Buffer(15 + shift) + Form1.Buffer(14 + shift)
    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Min, 2)


   Next 'End

    'Viene riempita la lista dei file disponibili
    If Form1.TreeViewDatalogDate.Items.Count > 0 Then
      Form1.TreeViewDatalogDate.Items.Clear
      For i = 30 To 1   'downto 1 do
      'begin
        If (FileSynchroVector(i).File_ID <= 30) Then
        'begin
          Form1.TreeViewDatalogDate.Items.AddFirst(Form1.TreeViewDatalogDate.TopItem,
          format('File %d: del %d.%d.%d alle %d:%d',
          (FileSynchroVector(i).File_ID, FileSynchroVector(i).File_Day,
          FileSynchroVector(i).File_Month, FileSynchroVector(i).File_Year,
          FileSynchroVector(i).File_Hour, FileSynchroVector(i).File_Min)))
        'End
        Else
          Form1.TreeViewDatalogDate.Items.AddFirst(Form1.TreeViewDatalogDate.TopItem, '------------')

      Next 'End
'End






'==============================================================================
'  SendDatalogRequest()
'    cambiare la descrizione!!!!
'
'==============================================================================
'procedure SendDatalogRequest(File_ID: Word LineNumber, LineAddress: byte )
'Var
  Checksum As Byte
  s As String
'begin

  Form1.CommPortDriver1.Sendbyte (SOF)
  Form1.CommPortDriver1.Sendbyte (FILE_DOWNLOAD_REQ)
  Checksum = SOF + FILE_DOWNLOAD_REQ

  s = inttohex(File_ID, 4)
  Form1.CommPortDriver1.Sendchar (s(4))
  Form1.CommPortDriver1.Sendchar (s(3))
  Form1.CommPortDriver1.Sendchar (s(2))
  Form1.CommPortDriver1.Sendchar (s(1))
  Checksum = Checksum + byte(s(1)) + byte(s(2))+ byte(s(3)) + byte(s(4))

  s = inttohex(LineNumber, 2)
  Form1.CommPortDriver1.Sendchar (s(2))
  Form1.CommPortDriver1.Sendchar (s(1))
  Checksum = Checksum + byte(s(1)) + byte(s(2))

  s = inttohex(LineAddress, 2)
  Form1.CommPortDriver1.Sendchar (s(2))
  Form1.CommPortDriver1.Sendchar (s(1))
  Checksum = Checksum + byte(s(1)) + byte(s(2))

  Checksum = 256 - (Checksum Mod 256)

  Form1.CommPortDriver1.Sendbyte (Checksum)
  If Checksum <> 0 Then
    Form1.CommPortDriver1.Sendbyte (EOF)
  Else
    exit

'End




'==============================================================================
'  SendDatalogStatusReq()
'
'
'==============================================================================
'Procedure SendDatalogStatusReq()
'Var
  Checksum As Byte
'begin
  Form1.DATA_REQUEST = False
  sleep (500)

  Form1.CommPortDriver1.Sendbyte (SOF)
  Form1.CommPortDriver1.Sendbyte (DATALOG_STATE_REQ)
  Checksum = SOF + DATALOG_STATE_REQ

  Checksum = 256 - (Checksum Mod 256)

  Form1.CommPortDriver1.Sendbyte (Checksum)

  If Checksum <> 0 Then
    Form1.CommPortDriver1.Sendbyte (EOF)
  Else
    exit

'End




'==============================================================================
'  SetDatalogStatus()
'
'
'==============================================================================
'Procedure SetDatalogStatus()
'Var
  Checksum As Byte
  s As String

'begin
  Form1.DATA_REQUEST = False
  sleep (500)
  Form1.CommPortDriver1.Sendbyte (SOF)
  Form1.CommPortDriver1.Sendbyte (DATALOG_ENDIS_REQ)
  Checksum = SOF + DATALOG_ENDIS_REQ


  s = inttohex(Datalog_Enable, 2)
  Form1.CommPortDriver1.Sendchar (s(2))
  Form1.CommPortDriver1.Sendchar (s(1))
  Checksum = Checksum + byte(s(1)) + byte(s(2))

  Checksum = 256 - (Checksum Mod 256)

  Form1.CommPortDriver1.Sendbyte (Checksum)
  If Checksum <> 0 Then
    Form1.CommPortDriver1.Sendbyte (EOF)
  Else
    exit
'End




'==============================================================================
'  ReceiveDatalogStatus()
'
'
'==============================================================================
'Procedure ReceiveDatalogStatus()
'Var
  DataTmp As String
'begin

  '| 2 | 3 |                                   (2 sent byte per la latdir)
  DataTmp = Form1.Buffer(3) + Form1.Buffer(2)
  HexToBin(PChar(DataTmp), @Datalog_Enable, 2)
  Select Case Datalog_Enable 'of
  
Case 1
    'begin
      Form1.Edit3.Brush.Color = cllime
      Form1.Edit3.Text = 'ENABLED'
    'End
Case 0
    'begin
      Form1.Edit3.Brush.Color = clred
      Form1.Edit3.Text = 'DISABLED'
    'End
Case Else
    'begin
      Form1.Edit3.Brush.Color = clWhite
      Form1.Edit3.Text = '--------'
    'End

  End Select

  Form1.DATA_REQUEST = True
'End


