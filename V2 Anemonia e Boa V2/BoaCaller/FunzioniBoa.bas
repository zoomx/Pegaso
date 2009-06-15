Attribute VB_Name = "FunzioniBoa"
'  Option Explicit
'
'  ' Datalog Download...
'  Const FILE_LAST_REQ = 15                        ' 0X0F
'  Const FILE_LAST_RESP = 16                       ' 0X10
'  Const FILE_INFO_REQ = 17                        ' 0X11
'  Const FILE_INFO_RESP = 18                       ' 0X12
'  Const FILE_SYNCHRO_REQ = 19                     ' 0X13
'  Const FILE_SYNCHRO_RESP = 20                    ' 0X14
'  Const FILE_DOWNLOAD_REQ = 21                    ' 0X15
'  Const FILE_DOWNLOAD_RESP = 22                   ' 0X16
'  Const DATALOG_ENDIS_REQ = 23                    ' 0X17
'  Const DATALOG_ENDIS_RESP = 24                   ' 0X18
'  Const DATALOG_SETFREQ_REQ = 25                  ' 0X19
'  Const DATALOG_SETFREQ_RESP = 26                 ' 0X1A
'  Const DATALOG_STATE_REQ = 27                    ' 0X1B
'  Const DATALOG_STATE_RESP = 28                   ' 0X1C
'
'  Const SOF = 1                                  ' start of frame
'  Const EOF = 0                                  ' end of frame
'
'Type TDaylyFileInfo
'    File_ID As Double  'word
'    File_Day As Byte
'    File_Month As Byte
'    File_Year As Byte
'    File_Hour As Byte
'    File_Min As Byte
'End Type
'
'Type TDatalogFields
'    DL_Day As Byte
'    DL_Month As Byte
'    DL_Year As Byte
'    DL_Hour As Byte
'    DL_Min As Byte
'    DL_Sec As Byte
'
'    '....Temperature
'    DL_T1 As Double         'SmallInt T1 (signed int16 moltiplicato x 10)
'    DL_T2 As Double         'SmallInt T2 (signed int16 moltiplicato x 10)
'    DL_T3 As Double         'SmallInt T3 (signed int16 moltiplicato x 10)
'
'    '....Meteo
'    DL_MeteoWindSpeed As Double     'Word (unsigned int16 moltiplicato x 10)
'    DL_MeteoWindDirection As Double 'Word (unsigned int16 moltiplicato x 10)
'
'    '....GPS
'    DL_GPS_Boa_Lat As Double    'integer (unsigned int32 moltiplicato x 10^7)
'    DL_GPS_Boa_Lon As Double    'integer (unsigned int32 moltiplicato x 10^7)
'    DL_GPS_Boa_Lat_Dir As Byte  'Char
'    DL_GPS_Boa_Lon_Dir As Byte  'Char
'    DL_GPSSatUsed As Byte       'byte
'
'    '....POWER
'    DL_MonitorBattery_1_12V As Byte             '  (unsigned int8 molt x 10)
'    DL_MonitorBattery_2_12V As Byte             '  (unsigned int8 molt x 10)
'    DL_MonitorBattery_3_12V As Byte             '  (unsigned int8 molt x 10)
'
'    '....Accelerometers
'    DL_AX As Integer     'SmallInt T1 (signed int16 moltiplicato x 10)
'    DL_AY As Integer     'SmallInt T2 (signed int16 moltiplicato x 10)
'    DL_AZ As Integer     'SmallInt T3 (signed int16 moltiplicato x 10)
'
'    '....H2O SENSORS
'    DL_H2O_1 As Byte             '  (unsigned int8 molt x 10)
'    DL_H2O_2 As Byte             '  (unsigned int8 molt x 10)
'End Type
'
''var
'' Last Closed File INFO
'    LCF_ID As Double    'Word = 0
'    LCF_Year As Byte    'byte = 0
'    LCF_Month As Byte   'byte = 0
'    LCF_Day As Byte     'byte = 0
'    LCF_Hour As Byte    'byte = 0
'    LCF_Minute As Byte  'byte = 0
'
'' File Info Req & download
'    Selected_File_ID As Double          'Word = 0
'    SelectedFileLinesNumber As Double   'Word = 0
'
'' Datalog Enable/Disable
'    Datalog_Enable As Byte  'byte = 3
'
''Datalog Sample time (minute)
'    DatalogSampleTime As Byte   'byte = 0
'
'    Dim FilesNameArray(30) As String      'array(1..30) of string
'
''Datalog Synchronization
'    Dim FileSynchroVector(30) As TDaylyFileInfo     'array(1..30) of TDaylyFileInfo
'
''Download
'    LinesToDownloadNumber As Byte   'byte = 4
'    LineTostartDownload As Double   'Word = 0
'    File_Line_Downloaded As Double  'Word = 0
'
'  'alla freq massima di una scrittura al minuto posso aver scritto su SD 60*24 = 1440 righe.
'    Dim DataDownloadedVector(1440) As TDatalogFields    'array(0..1440) of TDatalogFields
'  'Sapendo che ogni riga occupa 68 byte ne risulta una occupazione di memoria di 10183680 byte in delphi
'
'    prova As Byte 'byte = 0
'    Namefile As String
'    text As String
'    BuoyFile As Variant   'TextFile
'    i As Integer    'Integer
'
'
''==============================================================================
''  SendLastFileRequest()
''
''    Richiedo al Datalogger qual'è l'ultimo file chiuso ed aggiornato
''    Il datalogger mi risponde fornendomi:
''    - ID del file
''    - Day, Month, Hour, Min, Sec  della chiusura del file
''
''==============================================================================
''Procedure SendLastFileRequest()
''begin
''
''  Form1.DATA_REQUEST = False
''  Form1.SendRequest (FILE_LAST_REQ)
''
''End
''******************************************************************************
'
'
'
'
''Procedure ReceiveLastFile()
''Var
'  DataTmp As String
''begin
'
'  '| 2 | 3 | 4 | 5 |         (4 sent byte per l'ID dell'ultimo file chiuso)
'  DataTmp = Form1.Buffer(3) + Form1.Buffer(2) + Form1.Buffer(5) + Form1.Buffer(4)
'  HexToBin(PChar(DataTmp), @LCF_ID, 4)
'
'  '| 14 | 15 |                (2 sent byte per Last Closed File Minute)
'  DataTmp = Form1.Buffer(7) + Form1.Buffer(6)
'  HexToBin(PChar(DataTmp), @LCF_Minute, 2)
'
'  '| 12 | 13 |                (2 sent byte per Last Closed File Hour)
'  DataTmp = Form1.Buffer(9) + Form1.Buffer(8)
'  HexToBin(PChar(DataTmp), @LCF_Hour, 2)
'
'  '| 10 | 11 |                (2 sent byte per Last Closed File Day)
'  DataTmp = Form1.Buffer(11) + Form1.Buffer(10)
'  HexToBin(PChar(DataTmp), @LCF_Year, 2)
'
'  '| 8 | 9 |                  (2 sent byte per Last Closed File Month)
'  DataTmp = Form1.Buffer(13) + Form1.Buffer(12)
'  HexToBin(PChar(DataTmp), @LCF_Month, 2)
'
'  '| 6 | 7 |                  (2 sent byte per Last Closed File Year)
'  DataTmp = Form1.Buffer(15) + Form1.Buffer(14)
'  HexToBin(PChar(DataTmp), @LCF_Day, 2)
''End
''******************************************************************************
'
'
'
''==============================================================================
''  SendFileSynchroRequest()
''
''    Richiedo al Datalogger l'intero file di sincronizzazione
''    Il datalogger mi risponde fornendomi:
''    - ID di ogni file in memoria
''    - Day, Month, Hour, Min, Sec  di ciascun file
''
''==============================================================================
''Procedure SendFileSynchroRequest()
''begin
''
''  Form1.DATA_REQUEST = False
''  sleep (500)
''  Form1.SendRequest (FILE_SYNCHRO_REQ)
''
''End
'
'
'
'
''==============================================================================
''  SendFileInfoRequest()
''
''    Richiedo le informazioni di un file passando come parametro l'ID
''    Ricevo come risultato l'ID del file ed il numero di line
''    di cui è costituito
''
''==============================================================================
''procedure SendFileInfoRequest(File_ID: Word)
''Var
'  Checksum As Byte
'  s As String
'
''begin
'
'  Form1.DATA_REQUEST = False
'  'sleep (500)
'  Form1.CommPortDriver1.Sendbyte (SOF)
'  Form1.CommPortDriver1.Sendbyte (FILE_INFO_REQ)
'  Checksum = SOF + FILE_INFO_REQ
'
'
'  s = inttohex(File_ID, 4)
'  Form1.CommPortDriver1.Sendchar (s(4))
'  Form1.CommPortDriver1.Sendchar (s(3))
'  Form1.CommPortDriver1.Sendchar (s(2))
'  Form1.CommPortDriver1.Sendchar (s(1))
'  Checksum = Checksum + byte(s(1)) + byte(s(2))+ byte(s(3)) + byte(s(4))
'
'  Checksum = 256 - (Checksum Mod 256)
'
'  Form1.CommPortDriver1.Sendbyte (Checksum)
'  If Checksum <> 0 Then
'    Form1.CommPortDriver1.Sendbyte (EOF)
'  Else
'    exit
'
''End
'
'
'
'
'
''==============================================================================
''  SendPacketDownloadRequest()
''
''    Sapendo quante linee contiene il file che voglio scaricare
''    suddivido il download in tanti pacchetti.
''    Ogni volta che ricevo un pacchetto richiedo il successivo, fino
''    alla fine del file.
''
''==============================================================================
''procedure SendPacketDownloadRequest(File_ID: Word RowsNumber: byte RowIndex: Word)
''Var
'  Checksum As Byte
'  s As String
''begin
'
'  Inc (prova)
'  Form1.edit8.text = inttostr(prova)
'
'
'  Form1.CommPortDriver1.Sendbyte (SOF)
'  Form1.CommPortDriver1.Sendbyte (FILE_DOWNLOAD_REQ)
'  Checksum = SOF + FILE_DOWNLOAD_REQ
'
'  s = inttohex(Selected_File_ID, 4)
'  Form1.CommPortDriver1.Sendchar (s(4))
'  Form1.CommPortDriver1.Sendchar (s(3))
'  Form1.CommPortDriver1.Sendchar (s(2))
'  Form1.CommPortDriver1.Sendchar (s(1))
'  Checksum = Checksum + byte(s(1)) + byte(s(2))+ byte(s(3)) + byte(s(4))
'
'  s = inttohex(RowsNumber, 2)
'  Form1.CommPortDriver1.Sendchar (s(2))
'  Form1.CommPortDriver1.Sendchar (s(1))
'  Checksum = Checksum + byte(s(1)) + byte(s(2))
'
'  s = inttohex(RowIndex, 4)
'  Form1.CommPortDriver1.Sendchar (s(4))
'  Form1.CommPortDriver1.Sendchar (s(3))
'  Form1.CommPortDriver1.Sendchar (s(2))
'  Form1.CommPortDriver1.Sendchar (s(1))
'  Checksum = Checksum + byte(s(1)) + byte(s(2))+ byte(s(3)) + byte(s(4))
'
'  Checksum = 256 - (Checksum Mod 256)
'
'  Form1.CommPortDriver1.Sendbyte (Checksum)
'  If Checksum <> 0 Then
'    Form1.CommPortDriver1.Sendbyte (EOF)
'  Else
'    exit
'
''End
'
'
'
'
'
''Procedure ReceiveFileInfo()
''Var
'  DataTmp As String
''begin
'
'  '| 2 | 3 | 4 | 5 |             (4 sent byte File ID)
'  DataTmp = Form1.Buffer(3) + Form1.Buffer(2) + Form1.Buffer(5) + Form1.Buffer(4)
'  HexToBin(PChar(DataTmp), @Selected_File_ID, 4)
'
'  DataTmp = Form1.Buffer(7) + Form1.Buffer(6) + Form1.Buffer(9) + Form1.Buffer(8)
'  HexToBin(PChar(DataTmp), @SelectedFileLinesNumber, 4)
''End
'
'
'
'
'
''Procedure ReceiveFileSynchro()
''Var
'  i As Byte
'  shift As Integer  'Integer
'  DataTmp As String
''begin
'
''  Form1.Memo1.Clear
'
'  For i = 1 To 30 'do
'  'begin
'
'    shift = (i - 1) * 14
'
'    '| 2 | 3 | 4 | 5 |         (4 sent byte File ID)
'    DataTmp = Form1.Buffer(3 + shift) + Form1.Buffer(2 + shift) + Form1.Buffer(5 + shift) + Form1.Buffer(4 + shift)
'    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_ID, 4)
'
'    '| 6 | 7 |                     (2 sent byte per la File Minute)
'    DataTmp = Form1.Buffer(7 + shift) + Form1.Buffer(6 + shift)
'    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Day, 2)
'
'    '| 8 | 9 |                     (2 sent byte per la File Hour)
'    DataTmp = Form1.Buffer(9 + shift) + Form1.Buffer(8 + shift)
'    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Month, 2)
'
'    '| 10 | 11 |                     (2 sent byte per la File Day)
'    DataTmp = Form1.Buffer(11 + shift) + Form1.Buffer(10 + shift)
'    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Year, 2)
'
'    '| 12 | 13 |                     (2 sent byte per la File Month)
'    DataTmp = Form1.Buffer(13 + shift) + Form1.Buffer(12 + shift)
'    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Hour, 2)
'
'    '| 14 | 15 |                   (2 sent byte per la File Year)
'    DataTmp = Form1.Buffer(15 + shift) + Form1.Buffer(14 + shift)
'    HexToBin(PChar(DataTmp), @FileSynchroVector(i).File_Min, 2)
'
'
'   Next 'End
'
'    'Viene riempita la lista dei file disponibili
'    If Form1.TreeViewDatalogDate.Items.Count > 0 Then
'      Form1.TreeViewDatalogDate.Items.Clear
'      For i = 30 To 1   'downto 1 do
'      'begin
'        If (FileSynchroVector(i).File_ID <= 30) Then
'        'begin
'          Form1.TreeViewDatalogDate.Items.AddFirst(Form1.TreeViewDatalogDate.TopItem,
'          format('File %d: del %d.%d.%d alle %d:%d',
'          (FileSynchroVector(i).File_ID, FileSynchroVector(i).File_Day,
'          FileSynchroVector(i).File_Month, FileSynchroVector(i).File_Year,
'          FileSynchroVector(i).File_Hour, FileSynchroVector(i).File_Min)))
'        'End
'        Else
'          Form1.TreeViewDatalogDate.Items.AddFirst(Form1.TreeViewDatalogDate.TopItem, '------------')
'
'      Next 'End
''End
'
'
'
'
'
'
''==============================================================================
''  SendDatalogRequest()
''    cambiare la descrizione!!!!
''
''==============================================================================
''procedure SendDatalogRequest(File_ID: Word LineNumber, LineAddress: byte )
''Var
'  Checksum As Byte
'  s As String
''begin
'
'  Form1.CommPortDriver1.Sendbyte (SOF)
'  Form1.CommPortDriver1.Sendbyte (FILE_DOWNLOAD_REQ)
'  Checksum = SOF + FILE_DOWNLOAD_REQ
'
'  s = inttohex(File_ID, 4)
'  Form1.CommPortDriver1.Sendchar (s(4))
'  Form1.CommPortDriver1.Sendchar (s(3))
'  Form1.CommPortDriver1.Sendchar (s(2))
'  Form1.CommPortDriver1.Sendchar (s(1))
'  Checksum = Checksum + byte(s(1)) + byte(s(2))+ byte(s(3)) + byte(s(4))
'
'  s = inttohex(LineNumber, 2)
'  Form1.CommPortDriver1.Sendchar (s(2))
'  Form1.CommPortDriver1.Sendchar (s(1))
'  Checksum = Checksum + byte(s(1)) + byte(s(2))
'
'  s = inttohex(LineAddress, 2)
'  Form1.CommPortDriver1.Sendchar (s(2))
'  Form1.CommPortDriver1.Sendchar (s(1))
'  Checksum = Checksum + byte(s(1)) + byte(s(2))
'
'  Checksum = 256 - (Checksum Mod 256)
'
'  Form1.CommPortDriver1.Sendbyte (Checksum)
'  If Checksum <> 0 Then
'    Form1.CommPortDriver1.Sendbyte (EOF)
'  Else
'    exit
'
''End
'
'
'
'
''==============================================================================
''  SendDatalogStatusReq()
''
''
''==============================================================================
''Procedure SendDatalogStatusReq()
''Var
'  Checksum As Byte
''begin
'  Form1.DATA_REQUEST = False
'  sleep (500)
'
'  Form1.CommPortDriver1.Sendbyte (SOF)
'  Form1.CommPortDriver1.Sendbyte (DATALOG_STATE_REQ)
'  Checksum = SOF + DATALOG_STATE_REQ
'
'  Checksum = 256 - (Checksum Mod 256)
'
'  Form1.CommPortDriver1.Sendbyte (Checksum)
'
'  If Checksum <> 0 Then
'    Form1.CommPortDriver1.Sendbyte (EOF)
'  Else
'    exit
'
''End
'
'
'
'
''==============================================================================
''  SetDatalogStatus()
''
''
''==============================================================================
''Procedure SetDatalogStatus()
''Var
'  Checksum As Byte
'  s As String
'
''begin
'  Form1.DATA_REQUEST = False
'  sleep (500)
'  Form1.CommPortDriver1.Sendbyte (SOF)
'  Form1.CommPortDriver1.Sendbyte (DATALOG_ENDIS_REQ)
'  Checksum = SOF + DATALOG_ENDIS_REQ
'
'
'  s = inttohex(Datalog_Enable, 2)
'  Form1.CommPortDriver1.Sendchar (s(2))
'  Form1.CommPortDriver1.Sendchar (s(1))
'  Checksum = Checksum + byte(s(1)) + byte(s(2))
'
'  Checksum = 256 - (Checksum Mod 256)
'
'  Form1.CommPortDriver1.Sendbyte (Checksum)
'  If Checksum <> 0 Then
'    Form1.CommPortDriver1.Sendbyte (EOF)
'  Else
'    exit
''End
'
'
'
'
''==============================================================================
''  ReceiveDatalogStatus()
''
''
''==============================================================================
''Procedure ReceiveDatalogStatus()
''Var
'  DataTmp As String
''begin
'
'  '| 2 | 3 |                                   (2 sent byte per la latdir)
'  DataTmp = Form1.Buffer(3) + Form1.Buffer(2)
'  HexToBin(PChar(DataTmp), @Datalog_Enable, 2)
'  Select Case Datalog_Enable 'of
'
'Case 1
'    'begin
'      Form1.Edit3.Brush.Color = cllime
'      Form1.Edit3.Text = 'ENABLED'
'    'End
'Case 0
'    'begin
'      Form1.Edit3.Brush.Color = clred
'      Form1.Edit3.Text = 'DISABLED'
'    'End
'Case Else
'    'begin
'      Form1.Edit3.Brush.Color = clWhite
'      Form1.Edit3.Text = '--------'
'    'End
'
'  End Select
'
'  Form1.DATA_REQUEST = True
''End
'
'
''        { ricevo l'ID, la data e l'ora dell'ultimo file aggiornato e chiuso }
''        FILE_LAST_RESP:
''        begin
''         WaitForDatalogResponse = 0
''         WaitForReadSensorsCounter = 0
''         ReceiveLastFile()
''        End
'
''        { ricevo l'ID ed il numero di linee del file che si vuole scaricare }
''        FILE_INFO_RESP:
''        begin
''          WaitForDatalogResponse = 0
''          WaitForReadSensorsCounter = 0
''          ReceiveFileInfo()
''          BitBtnDownloadFile.Visible = True
''        End
'
'
'
'
'FILE_DOWNLOAD_RESP:
'        'begin
'        Prova2 = prova3 + 1 'inc (prova3)
'
'          'edit9.text = inttostr(prova3)
'          WaitForDatalogResponse = 0
'          WaitForReadSensorsCounter = 0
'          DOWNLOAD_REQUEST = True
'
'          edit5.text = inttostr(File_Line_Downloaded)
'          edit6.text = inttostr(SelectedFileLinesNumber)
'
'          If (LineTostartDownload = 1) Then
'            Datalog_File_Name = "Log_" + FormatDateTime("dd_mm_yy", Date) + "_" + FormatDateTime("hh.mm.ss", Now) + ".txt"
'          End If
'
'         If (File_Line_Downloaded < SelectedFileLinesNumber) Then
'            'begin
'              LineTostartDownload = LineTostartDownload + 4 'Inc(LineTostartDownload, 4)
'              SendPacketDownloadRequest(Selected_File_ID, LinesToDownloadNumber, LineTostartDownload)
'            'End
'         Else
'             'begin
'              FormDataTransfer.Close    'Chiude il form del download
'              DOWNLOAD_REQUEST = False
'         End If
'
'          SetLength(char_tmp, SerialByteIndex)
'
'          j = 1
'          i = 2 'non considero Buffer(0) e Buffer(1)
'
'          while i < SerialByteIndex - 2 do
'          begin
'
'            DataTmp = Buffer(i + 1) + Buffer(i)
'            HexToBin(PChar(DataTmp), @char_tmp(j), 2)
'            If ((j = 76) Or (j = (76 * 2)) Or (j = (76 * 3)) Or (j = (76 * 4))) Then
'            begin
'              File_Line_Downloaded = File_Line_Downloaded + 1 'Inc (File_Line_Downloaded)
'
'              FormDataTransfer.ProgressBarDownload.Position = File_Line_Downloaded
'              If (SelectedFileLinesNumber <> 0) Then
'              dwn = (File_Line_Downloaded / SelectedFileLinesNumber) * 100
'              FormDataTransfer.Caption = format('download of file %d in progress...%d %s', (Selected_File_ID, trunc(dwn), char(37)))
'              FormDataTransfer.LabelDownload.Caption = format(' Downloaded: %d of %d', (File_Line_Downloaded, SelectedFileLinesNumber))
'            End
'
'            i = i + 2 'Inc(i, 2)
'            j = j + 1 'Inc (j)
'
'          End
'
'            'Interpretazione dei dati
'          For i = 0 To 3    ' do
'          'begin
'            shift = 76 * i
'            SetLength(DataVect, 24)
'            Rindex = i + File_Line_Downloaded - 4
'
'            DataTmp = char_tmp(1 + shift) + char_tmp(2 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_Day, 8)
'            DataVect(0) = DataDownloadedVector(Rindex).DL_Day
'
'            DataTmp = char_tmp(3 + shift) + char_tmp(4 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_Month, 8)
'            DataVect(1) = DataDownloadedVector(Rindex).DL_Month
'
'            DataTmp = char_tmp(5 + shift) + char_tmp(6 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_Year, 8)
'            DataVect(2) = DataDownloadedVector(Rindex).DL_Year
'
'            DataTmp = char_tmp(7 + shift) + char_tmp(8 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_Hour, 8)
'            DataVect(3) = DataDownloadedVector(Rindex).DL_Hour
'
'            DataTmp = char_tmp(9 + shift) + char_tmp(10 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_Min, 8)
'            DataVect(4) = DataDownloadedVector(Rindex).DL_Min
'
'            DataTmp = char_tmp(11 + shift) + char_tmp(12 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_Sec, 8)
'            DataVect(5) = DataDownloadedVector(Rindex).DL_Sec
'
'            DataTmp = char_tmp(15 + shift) + char_tmp(16 + shift) + char_tmp(13 + shift) + char_tmp(14 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_T1, 4)
'            DataVect(6) = DataDownloadedVector(Rindex).DL_T1 / 10
'
'            DataTmp = char_tmp(19 + shift) + char_tmp(20 + shift) + char_tmp(17 + shift) + char_tmp(18 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_T2, 4)
'            DataVect(7) = DataDownloadedVector(Rindex).DL_T2 / 10
'
'            DataTmp = char_tmp(23 + shift) + char_tmp(24 + shift) + char_tmp(21 + shift) + char_tmp(22 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_T3, 4)
'            DataVect(8) = DataDownloadedVector(Rindex).DL_T3 / 10
'
'            DataTmp = char_tmp(27 + shift) + char_tmp(28 + shift) + char_tmp(25 + shift) + char_tmp(26 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_MeteoWindSpeed, 4)
'            DataVect(9) = DataDownloadedVector(Rindex).DL_MeteoWindSpeed * 1.852 / 10
'
'            DataTmp = char_tmp(31 + shift) + char_tmp(32 + shift) + char_tmp(29 + shift) + char_tmp(30 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_MeteoWindDirection, 4)
'            DataVect(10) = DataDownloadedVector(Rindex).DL_MeteoWindDirection / 10
'
'            DataTmp = char_tmp(39 + shift) + char_tmp(40 + shift) + char_tmp(37 + shift) + char_tmp(38 + shift) + char_tmp(35 + shift) + char_tmp(36 + shift) + char_tmp(33 + shift) + char_tmp(34 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_GPS_Boa_Lat, 8)
'            DataVect(11) = DataDownloadedVector(Rindex).DL_GPS_Boa_Lat / 10000000
'
'            DataTmp = char_tmp(41 + shift) + char_tmp(42 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_GPS_Boa_Lat_Dir, 8)
'            DataVect(12) = byte(DataDownloadedVector(Rindex).DL_GPS_Boa_Lat_Dir)
'
'            DataTmp = char_tmp(49 + shift) + char_tmp(50 + shift) + char_tmp(47 + shift) + char_tmp(48 + shift) + char_tmp(45 + shift) + char_tmp(46 + shift) + char_tmp(43 + shift) + char_tmp(44 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_GPS_Boa_Lon, 8)
'            DataVect(13) = DataDownloadedVector(Rindex).DL_GPS_Boa_Lon / 10000000
'
'            DataTmp = char_tmp(51 + shift) + char_tmp(52 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_GPS_Boa_Lon_Dir, 8)
'            DataVect(14) = byte(DataDownloadedVector(Rindex).DL_GPS_Boa_Lon_Dir)
'
'            DataTmp = char_tmp(53 + shift) + char_tmp(54 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_GPSSatUsed, 8)
'            DataVect(15) = DataDownloadedVector(Rindex).DL_GPSSatUsed
'
'            DataTmp = char_tmp(55 + shift) + char_tmp(56 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_MonitorBattery_3_12V, 8)
'            DataVect(16) = DataDownloadedVector(Rindex).DL_MonitorBattery_3_12V / 10
'
'            DataTmp = char_tmp(57 + shift) + char_tmp(58 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_MonitorBattery_2_12V, 8)
'            DataVect(17) = DataDownloadedVector(Rindex).DL_MonitorBattery_2_12V / 10
'
'            DataTmp = char_tmp(59 + shift) + char_tmp(60 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_MonitorBattery_1_12V, 8)
'            DataVect(18) = DataDownloadedVector(Rindex).DL_MonitorBattery_1_12V / 10
'
'            DataTmp = char_tmp(63 + shift) + char_tmp(64 + shift) + char_tmp(61 + shift) + char_tmp(62 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_AX, 4)
'            DataVect(19) = DataDownloadedVector(Rindex).DL_AX / 1000
'
'            DataTmp = char_tmp(67 + shift) + char_tmp(68 + shift) + char_tmp(65 + shift) + char_tmp(66 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_AY, 4)
'            DataVect(20) = DataDownloadedVector(Rindex).DL_AY / 1000
'
'            DataTmp = char_tmp(71 + shift) + char_tmp(72 + shift) + char_tmp(69 + shift) + char_tmp(70 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_AZ, 4)
'            DataVect(21) = DataDownloadedVector(Rindex).DL_AZ / 1000
'
'            DataTmp = char_tmp(73 + shift) + char_tmp(74 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_H2O_1, 8)
'            DataVect(22) = DataDownloadedVector(Rindex).DL_H2O_1 / 10
'
'            DataTmp = char_tmp(75 + shift) + char_tmp(76 + shift)
'            HexToBin(PChar(DataTmp), @DataDownloadedVector(Rindex).DL_H2O_2, 8)
'            DataVect(23) = DataDownloadedVector(Rindex).DL_H2O_2 / 10
'
'            PutDataInFile('C:\Programmi\OTe Systems\Buoy GUI\Datalog\', Datalog_File_Name, '|', DataVect)
'
'          If (Rindex < SelectedFileLinesNumber) Then
'                                        'day  month year hour  min sec SPSTe BOATe CPSTe Wspd  Wdir  Lat  dir  Lon dir     T3   T2    T3     AX    AY    AZ    H2O1  H2O2
'            Memo2.Lines.add(format('-%d- |%.2d.%.2d.%.2d|%.2d:%.2d:%.2d|%2.1f\%2.1f\%2.1f\%2.1f\%2.1f\%2.7f\%s\%2.7f\%s\%d\%2.1f\%2.1f\%2.1f\%2.1f\%2.1f\%2.1f\%2.1f\%2.1f',(
'            Rindex,
'            DataDownloadedVector(Rindex).DL_Day,
'            DataDownloadedVector(Rindex).DL_Month,
'            DataDownloadedVector(Rindex).DL_Year,
'            DataDownloadedVector(Rindex).DL_Hour,
'            DataDownloadedVector(Rindex).DL_Min,
'            DataDownloadedVector(Rindex).DL_Sec,
'            DataDownloadedVector(Rindex).DL_T1/10,
'            DataDownloadedVector(Rindex).DL_T2/10,
'            DataDownloadedVector(Rindex).DL_T3/10,
'            DataDownloadedVector(Rindex).DL_MeteoWindSpeed/10,
'            DataDownloadedVector(Rindex).DL_MeteoWindDirection*1.852/10,
'            DataDownloadedVector(Rindex).DL_GPS_Boa_Lat/10000000,
'            DataDownloadedVector(Rindex).DL_GPS_Boa_Lat_Dir,
'            DataDownloadedVector(Rindex).DL_GPS_Boa_Lon/10000000,
'            DataDownloadedVector(Rindex).DL_GPS_Boa_Lon_Dir,
'            DataDownloadedVector(Rindex).DL_GPSSatUsed,
'            DataDownloadedVector(Rindex).DL_MonitorBattery_3_12V/10,
'            DataDownloadedVector(Rindex).DL_MonitorBattery_2_12V/10,
'            DataDownloadedVector(Rindex).DL_MonitorBattery_1_12V/10,
'            DataDownloadedVector(Rindex).DL_AX/1000,
'            DataDownloadedVector(Rindex).DL_AY/1000,
'            DataDownloadedVector(Rindex).DL_AZ/1000,
'            DataDownloadedVector(Rindex).DL_H2O_1/10,
'            DataDownloadedVector(Rindex).DL_H2O_2/10
'            )))
'
''   Riempimento celle di Excel
'' (*           WS.Cells.Item(Rindex+1, 1).Value = DataDownloadedVector(Rindex).DL_Day
''            WS.Cells.Item(Rindex + 1, 2).Value = DataDownloadedVector(Rindex).DL_Month
''            WS.Cells.Item(Rindex + 1, 3).Value = DataDownloadedVector(Rindex).DL_Year
''            WS.Cells.Item(Rindex + 1, 4).Value = DataDownloadedVector(Rindex).DL_Hour
''            WS.Cells.Item(Rindex + 1, 5).Value = DataDownloadedVector(Rindex).DL_Min
''            WS.Cells.Item(Rindex + 1, 6).Value = DataDownloadedVector(Rindex).DL_Sec
''            WS.Cells.Item(Rindex + 1, 7).Value = DataDownloadedVector(Rindex).DL_T1 / 10
''            WS.Cells.Item(Rindex + 1, 8).Value = DataDownloadedVector(Rindex).DL_T2 / 10
''            WS.Cells.Item(Rindex + 1, 9).Value = DataDownloadedVector(Rindex).DL_T3 / 10
''            WS.Cells.Item(Rindex + 1, 10).Value = DataDownloadedVector(Rindex).DL_MeteoWindSpeed / 10
''            WS.Cells.Item(Rindex + 1, 11).Value = DataDownloadedVector(Rindex).DL_MeteoWindDirection * 1.852 / 10
''            WS.Cells.Item(Rindex + 1, 12).Value = DataDownloadedVector(Rindex).DL_GPS_Boa_Lat / 10000000
''            WS.Cells.Item(Rindex + 1, 13).Value = DataDownloadedVector(Rindex).DL_GPS_Boa_Lat_Dir
''            WS.Cells.Item(Rindex + 1, 14).Value = DataDownloadedVector(Rindex).DL_GPS_Boa_Lon / 10000000
''            WS.Cells.Item(Rindex + 1, 15).Value = DataDownloadedVector(Rindex).DL_GPS_Boa_Lon_Dir
''            WS.Cells.Item(Rindex + 1, 16).Value = DataDownloadedVector(Rindex).DL_GPSSatUsed
''            WS.Cells.Item(Rindex + 1, 17).Value = DataDownloadedVector(Rindex).DL_MonitorBattery_3_12V / 10
''            WS.Cells.Item(Rindex + 1, 18).Value = DataDownloadedVector(Rindex).DL_MonitorBattery_2_12V / 10
''            WS.Cells.Item(Rindex + 1, 19).Value = DataDownloadedVector(Rindex).DL_MonitorBattery_1_12V / 10
''            WS.Cells.Item(Rindex + 1, 20).Value = DataDownloadedVector(Rindex).DL_AX / 1000
''            WS.Cells.Item(Rindex + 1, 21).Value = DataDownloadedVector(Rindex).DL_AY / 1000
''            WS.Cells.Item(Rindex + 1, 22).Value = DataDownloadedVector(Rindex).DL_AZ / 1000
''            WS.Cells.Item(Rindex + 1, 23).Value = DataDownloadedVector(Rindex).DL_H2O_1 / 10
''            WS.Cells.Item(Rindex + 1, 24).Value = DataDownloadedVector(Rindex).DL_H2O_2 / 10
'' *)
'            Next i
'
'
'       End
'
'DATALOG_STATE_RESP:
'         ReceiveDatalogStatus()
'
'DATALOG_ENDIS_RESP:
'          ReceiveDatalogStatus()
'
'
''==============================================================================
'
