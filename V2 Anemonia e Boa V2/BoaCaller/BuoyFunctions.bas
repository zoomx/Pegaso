Attribute VB_Name = "BuoyFunctions"
Option Explicit
    Const SOF = 1                                  ' start of frame
    Const EOF = 0                                  ' end of frame

    Const FILE_LAST_REQ As Byte = 15                      ' 0X0F
    Const FILE_LAST_RESP As Byte = 16                      ' 0X10
    Const FILE_INFO_REQ As Byte = 17                        ' 0X11
    Const FILE_INFO_RESP As Byte = 18                      ' 0X12
    Const FILE_DOWNLOAD_REQ As Byte = 21                   ' 0X15
    Const FILE_DOWNLOAD_RESP As Byte = 22                  ' 0X16
    
    Public CommPort As Long
    Public CommSettings As String
    Public ChiamaFlag As Boolean    'Serve per sapere
                                'se il bottone chiama
                                'è stato premuto
    Public Connected As Boolean
    Public SubName As String
    '

Public Function ParseLine(Linea As String) As String
    Dim Stringbuffer As String
    Dim StringBuffer2 As String
    Dim DateDay As String
    Dim DateMonth As String
    Dim DateYear As String
    Dim DateHour As String
    Dim DateMin As String
    Dim DateSec As String
    Dim T1 As Single
    Dim T2 As Single
    Dim T3 As Single
    Dim MeteoWindSpeed As Single
    Dim MeteoWindDirection As Single
    Dim GPS_Lat As Double
    Dim GPS_LatDir As Single
    Dim GPS_Lon As Double
    Dim GPS_LonDir As Single
    Dim GPS_SatUsed As Byte
    Dim MonitorBattery_3_12V As Single
    Dim MonitorBattery_2_12V As Single
    Dim MonitorBattery_1_12V As Single
    Dim AX As Single
    Dim AY As Single
    Dim AZ As Single
    Dim H20_1 As Single
    Dim H2O_2 As Single
    
    'Da cancellare
    Dim SPStemp As Single
    Dim BoaTemp As Single
    Dim CPSTemp As Single
    
    SubName = "ParseLine"
    ParseLine = ""
    
    'Day
    Stringbuffer = Left(Linea, 2)
    'Debug.Print "Stringbuffer="; StringBuffer
    Stringbuffer = HexToDecAscii(Stringbuffer)
    'Debug.Print StringBuffer
    DateDay = ZeroPad(Stringbuffer)
    
    'Month
    Stringbuffer = Mid(Linea, 3, 2)
    'Debug.Print "Stringbuffer="; StringBuffer
    Stringbuffer = HexToDecAscii(Stringbuffer)
    'Debug.Print StringBuffer
    DateMonth = ZeroPad(Stringbuffer)
    
    'Year
    Stringbuffer = Mid(Linea, 5, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    DateYear = "20" + ZeroPad(Stringbuffer)
    
    'Hour
    Stringbuffer = Mid(Linea, 7, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    DateHour = ZeroPad(Stringbuffer)

    'Minute
    Stringbuffer = Mid(Linea, 9, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    DateMin = ZeroPad(Stringbuffer)
    
    'Seconds
    Stringbuffer = Mid(Linea, 11, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    DateSec = ZeroPad(Stringbuffer)

    'SPStemp
    Stringbuffer = Mid(Linea, 13, 4)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    SPStemp = Val(Stringbuffer) / 10
    
    'BOAtemp
    Stringbuffer = Mid(Linea, 17, 4)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    BoaTemp = Val(Stringbuffer) / 10
    
    'CPStemp
    Stringbuffer = Mid(Linea, 21, 4)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    CPSTemp = Val(Stringbuffer) / 10
    
    'MeteoWindSpeed
    Stringbuffer = Mid(Linea, 25, 4)
    StringBuffer2 = Right(Stringbuffer, 2) & Left(Stringbuffer, 2)
    Stringbuffer = HexToDecAscii(StringBuffer2)
    MeteoWindSpeed = Val(Stringbuffer) * 1.852 / 10
    
    'MeteoWindDirection
    Stringbuffer = Mid(Linea, 29, 4)
    StringBuffer2 = Right(Stringbuffer, 2) & Left(Stringbuffer, 2)
    Stringbuffer = HexToDecAscii(StringBuffer2)
    MeteoWindDirection = Val(Stringbuffer) / 10
    
    'GPS_Lat
    Stringbuffer = Mid(Linea, 33, 8)
    'StringBuffer2 = Right(StringBuffer, 2) & Mid
    Stringbuffer = HexToDecAscii(StringBuffer2)
    GPS_Lat = Val(Stringbuffer) / 10000000
    
    'GPS_LatDir
    Stringbuffer = Mid(Linea, 41, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    GPS_LatDir = Val(Stringbuffer)
    
    'GPS_Lon
    Stringbuffer = Mid(Linea, 43, 8)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    GPS_Lon = Val(Stringbuffer) / 10000000
    
    'GPS_LonDir
    Stringbuffer = Mid(Linea, 51, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    GPS_LonDir = Val(Stringbuffer)
    
    'GPS_SatUsed
    Stringbuffer = Mid(Linea, 53, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    GPS_LonDir = Val(Stringbuffer)
   
    'MonitorBattery_3_12V
    Stringbuffer = Mid(Linea, 55, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    MonitorBattery_3_12V = Val(Stringbuffer) / 10
    
    'MonitorBattery_2_12V
    Stringbuffer = Mid(Linea, 57, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    MonitorBattery_2_12V = Val(Stringbuffer) / 10
    
    'MonitorBattery_1_12V
    Stringbuffer = Mid(Linea, 59, 2)
    Stringbuffer = HexToDecAscii(Stringbuffer)
    MonitorBattery_1_12V = Val(Stringbuffer) / 10
   
'    Debug.Print DateDay; " "; DateMonth; " "; DateYear; " "; DateHour; " "; DateMin; " "; DateSec
'    Debug.Print SPStemp, BoaTemp, CPSTemp
'    Debug.Print MeteoWindSpeed, MeteoWindDirection
'    Debug.Print GPS_Lat, GPS_Lon, GPS_LatDir, GPS_LonDir
    
    ParseLine = DateDay & "/" & DateMonth & "/" & DateYear & " " & DateHour & ":" & DateMin & ":" & DateSec & "; "
    ParseLine = ParseLine & Trim(Str(SPStemp)) & "; " & Trim(Str(BoaTemp)) & "; " & Trim(Str(CPSTemp)) & "; "
    ParseLine = ParseLine & Trim(Str(MeteoWindSpeed)) & "; " & Trim(Str(MeteoWindDirection)) & "; "
    ParseLine = ParseLine & Trim(Str(GPS_Lat)) & "; " & Trim(Str(GPS_LatDir)) & "; " & Trim(Str(GPS_Lon)) & "; " & Trim(Str(GPS_LonDir)) & "; "
    ParseLine = ParseLine & Trim(Str(GPS_SatUsed)) & "; "
    ParseLine = ParseLine & Trim(Str(MonitorBattery_3_12V)) & "; " & Trim(Str(MonitorBattery_1_12V)) & "; " & Trim(Str(MonitorBattery_1_12V))
    
    
    'Debug.Print ParseLine
End Function

Public Function ParseLine2(Linea As String) As String
    Dim Stringbuffer As String
    Dim StringBuffer2 As String
    Dim DateDay As String
    Dim DateMonth As String
    Dim DateYear As String
    Dim DateHour As String
    Dim DateMin As String
    Dim DateSec As String
    Dim T1 As Single
    Dim T2 As Single
    Dim T3 As Single
    Dim MeteoWindSpeed As Single
    Dim MeteoWindDirection As Single
    Dim GPS_Lat As Double
    Dim GPS_LatDir As Single
    Dim GPS_Lon As Double
    Dim GPS_LonDir As Single
    Dim GPS_SatUsed As Byte
    Dim MonitorBattery_3_12V As Single
    Dim MonitorBattery_2_12V As Single
    Dim MonitorBattery_1_12V As Single
    Dim AX As Single
    Dim AY As Single
    Dim AZ As Single
    Dim H20_1 As Single
    Dim H2O_2 As Single
    
    'Da cancellare
    Dim SPStemp As Single
    Dim BoaTemp As Single
    Dim CPSTemp As Single
    
    SubName = "ParseLine2"
    ParseLine2 = ""
    
    'Day
    Stringbuffer = Left(Linea, 4)
    'Debug.Print "Stringbuffer="; StringBuffer
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    'Debug.Print StringBuffer
    DateDay = ZeroPad(Stringbuffer)
    
    'Month
    Stringbuffer = Mid(Linea, 5, 4)
    'Debug.Print "Stringbuffer="; StringBuffer
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    'Debug.Print StringBuffer
    DateMonth = ZeroPad(Stringbuffer)
    
    'Year
    Stringbuffer = Mid(Linea, 9, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    DateYear = "20" + ZeroPad(Stringbuffer)
    If Val(DateYear) > (Year(Now) + 1) Then DateYear = Year(Now)        'Problemi con il 31 dicembre. Forse risolto!
    
    'Hour
    Stringbuffer = Mid(Linea, 13, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    DateHour = ZeroPad(Stringbuffer)

    'Minute
    Stringbuffer = Mid(Linea, 17, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    DateMin = ZeroPad(Stringbuffer)
    
    'Seconds
    Stringbuffer = Mid(Linea, 21, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    DateSec = ZeroPad(Stringbuffer)

    'T1
    Stringbuffer = Mid(Linea, 25, 8)
    Stringbuffer = HexToDecAscii3(Stringbuffer)
    T1 = Val(Stringbuffer) / 10
    
    'T2
    Stringbuffer = Mid(Linea, 33, 8)
    Stringbuffer = HexToDecAscii3(Stringbuffer)
    T2 = Val(Stringbuffer) / 10
    
    'T3
    Stringbuffer = Mid(Linea, 41, 8)
    Stringbuffer = HexToDecAscii3(Stringbuffer)
    T3 = Val(Stringbuffer) / 10
    
    'MeteoWindSpeed
    Stringbuffer = Mid(Linea, 49, 8)
    'StringBuffer2 = Right(Stringbuffer, 2) & Left(Stringbuffer, 2)
    Stringbuffer = HexToDecAscii3(StringBuffer2)
    MeteoWindSpeed = Val(Stringbuffer) * 1.852 / 10
    
    'MeteoWindDirection
    Stringbuffer = Mid(Linea, 57, 8)
    'StringBuffer2 = Right(Stringbuffer, 2) & Left(Stringbuffer, 2)
    Stringbuffer = HexToDecAscii3(StringBuffer2)
    MeteoWindDirection = Val(Stringbuffer) / 10
    
    'GPS_Lat
    Stringbuffer = Mid(Linea, 65, 16)
    'StringBuffer2 = Right(StringBuffer, 2) & Mid
    Stringbuffer = HexToDecAscii3(StringBuffer2)
    GPS_Lat = Val(Stringbuffer) / 10000000
    
    'GPS_LatDir
    Stringbuffer = Mid(Linea, 81, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    GPS_LatDir = Val(Stringbuffer)
    
    'GPS_Lon
    Stringbuffer = Mid(Linea, 85, 16)
    Stringbuffer = HexToDecAscii3(Stringbuffer)
    GPS_Lon = Val(Stringbuffer) / 10000000
    
    'GPS_LonDir
    Stringbuffer = Mid(Linea, 101, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    GPS_LonDir = Val(Stringbuffer)
    
    'GPS_SatUsed
    Stringbuffer = Mid(Linea, 105, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    GPS_LonDir = Val(Stringbuffer)
   
    'MonitorBattery_3_12V
    Stringbuffer = Mid(Linea, 109, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    MonitorBattery_3_12V = Val(Stringbuffer) / 10
    
    'MonitorBattery_2_12V
    Stringbuffer = Mid(Linea, 113, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    MonitorBattery_2_12V = Val(Stringbuffer) / 10
    
    'MonitorBattery_1_12V
    Stringbuffer = Mid(Linea, 117, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    MonitorBattery_1_12V = Val(Stringbuffer) / 10
    
    'AX
    Stringbuffer = Mid(Linea, 121, 8)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    AX = Val(Stringbuffer) / 1000
    
    'AY
    Stringbuffer = Mid(Linea, 129, 8)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    AY = Val(Stringbuffer) / 1000

    'AZ
    Stringbuffer = Mid(Linea, 137, 8)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    AZ = Val(Stringbuffer) / 1000
    
    'H20_1
    Stringbuffer = Mid(Linea, 145, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    H20_1 = Val(Stringbuffer) / 10

    'H20_2
    Stringbuffer = Mid(Linea, 149, 4)
    Stringbuffer = HexToDecAscii2(Stringbuffer)
    H20_1 = Val(Stringbuffer) / 10
    
   
'    Debug.Print DateDay; " "; DateMonth; " "; DateYear; " "; DateHour; " "; DateMin; " "; DateSec
'    Debug.Print SPStemp, BoaTemp, CPSTemp
'    Debug.Print MeteoWindSpeed, MeteoWindDirection
'    Debug.Print GPS_Lat, GPS_Lon, GPS_LatDir, GPS_LonDir
    
    ParseLine2 = DateDay & "/" & DateMonth & "/" & DateYear & " " & DateHour & ":" & DateMin & ":" & DateSec & "; "
    ParseLine2 = ParseLine2 & Trim(Str(T1)) & "; " & Trim(Str(T2)) & "; " & Trim(Str(T3)) & "; "
    ParseLine2 = ParseLine2 & Trim(Str(MeteoWindSpeed)) & "; " & Trim(Str(MeteoWindDirection)) & "; "
    ParseLine2 = ParseLine2 & Trim(Str(GPS_Lat)) & "; " & Trim(Str(GPS_LatDir)) & "; " & Trim(Str(GPS_Lon)) & "; " & Trim(Str(GPS_LonDir)) & "; "
    ParseLine2 = ParseLine2 & Trim(Str(GPS_SatUsed)) & "; "
    ParseLine2 = ParseLine2 & Trim(Str(MonitorBattery_3_12V)) & "; " & Trim(Str(MonitorBattery_1_12V)) & "; " & Trim(Str(MonitorBattery_1_12V)) & "; "
    ParseLine2 = ParseLine2 & Trim(Str(AX)) & "; " & Trim(Str(AY)) & "; " & Trim(Str(AZ)) & "; "
    ParseLine2 = ParseLine2 & Trim(Str(H20_1)) & "; " & Trim(Str(H20_1))
    
    'Debug.Print ParseLine2
End Function

Public Function HexToDecAscii(Stringa As String) As String
    'Converte una stringa di Hex in Ascii in una stringa di decimali
    Dim LenghtStringa As Integer
    Dim Intero As Long
    SubName = "HexToDecAscii"
    LenghtStringa = Len(Stringa)
    Stringa = "&H" & Stringa
    Intero = Val(Stringa)
    HexToDecAscii = Trim(Str(Intero))
End Function

Public Function HexToDecAscii2(Stringa As String) As String
    'Converte una stringa di Hex in Ascii in una stringa di decimali
    Dim LenghtStringa As Integer
    Dim Intero As Long
    Dim Char1 As String
    Dim Char2 As String
    
    SubName = "HexToDecAscii2"
    
    Char1 = "&H" + SwapString(Left(Stringa, 2))
    Char2 = "&H" + SwapString(Right(Stringa, 2))
    
    'LenghtStringa = Len(Stringa)
    Stringa = "&H" + Chr(Val(Char1)) + Chr(Val(Char2))
    Intero = Val(Stringa)
    HexToDecAscii2 = Trim(Str(Intero))
End Function

Public Function HexToDecAscii3(Stringa As String) As String
    'Converte una stringa di Hex in Ascii in una stringa di decimali
    Dim LenghtStringa As Integer
    Dim Intero As Long
    Dim Char1 As String
    Dim Char2 As String
    Dim Char3 As String
    Dim Char4 As String
    
    SubName = "HexToDecAscii3"
    
    Char1 = "&H" + SwapString(Left(Stringa, 2))
    Char2 = "&H" + SwapString(Mid(Stringa, 3, 2))
    Char3 = "&H" + SwapString(Mid(Stringa, 5, 2))
    Char4 = "&H" + SwapString(Right(Stringa, 2))
    
    'LenghtStringa = Len(Stringa)
    Stringa = "&H" + Chr(Val(Char1)) + Chr(Val(Char2)) + Chr(Val(Char3)) + Chr(Val(Char4))
    Intero = Val(Stringa)
    HexToDecAscii3 = Trim(Str(Intero))
End Function

Public Function ZeroPad(Stringa As String) As String
    SubName = "ZeroPad"
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    If Len(Stringa) = 0 Then Stringa = "00"
    ZeroPad = Stringa
End Function

Public Function SendLastFileRequest() As Boolean
'    Richiedo al Datalogger qual'è l'ultimo file chiuso ed aggiornato
'    Il datalogger mi risponde fornendomi:
'    - ID del file
'    - Day, Month, Hour, Min, Sec  della chiusura del file

    Dim Packet(3) As Byte
    
    SubName = "SendLastFileRequest"
    
    Packet(0) = SOF   'SOF=1
    Packet(1) = FILE_LAST_REQ  'FILE_LAST_REQ=15
    Packet(2) = 1 + 15 'packet(0)+Packet(1) Checksum
    Packet(2) = 256 - Packet(2)
    Packet(3) = 0
    If fMain.MSComm1.PortOpen = True Then
        fMain.MSComm1.Output = Packet()
        SendLastFileRequest = True
    Else
        SendLastFileRequest = False
    End If
    
End Function

Public Function SendFileInfoRequest(ID As Long) As Boolean
'    Richiedo le informazioni di un file passando come parametro l'ID
'    Ricevo come risultato l'ID del file ed il numero di line
'    di cui è costituito

    Dim Packet(7) As Byte
    Dim Stringa As String
    Dim i As Integer
    
    SubName = "SendFileInfoRequest"
    
    Packet(0) = SOF   'SOF=1
    Packet(1) = FILE_INFO_REQ  'FILE_INFO_REQ=17
    Stringa = Hex(ID)
    If Len(Stringa) = 1 Then Stringa = "000" + Stringa
    If Len(Stringa) = 2 Then Stringa = "00" + Stringa
    If Len(Stringa) = 3 Then Stringa = "0" + Stringa
    'Debug.Print Stringa
    'Stringa = SwapString(Stringa)
    Packet(2) = Asc(Right(Stringa, 1))
    Packet(3) = Asc(Mid(Stringa, 3, 1))
    Packet(4) = Asc(Mid(Stringa, 2, 1))
    Packet(5) = Asc(Left(Stringa, 1))
    Packet(6) = 1 + 17 + Packet(2) + Packet(3) + Packet(4) + Packet(5) 'Checksum e l'overflow?
    Packet(6) = 256 - (Packet(6) Mod 256)
    Packet(7) = 0
'    For i = 0 To 7
'        Debug.Print Chr(Packet(i));
'    Next i
    
    If fMain.MSComm1.PortOpen = True Then
        fMain.MSComm1.Output = Packet()
        SendFileInfoRequest = True
    Else
        SendFileInfoRequest = False
    End If

End Function

Public Function SendPacketDownloadRequest(ID As Long, RowsNumber As Byte, RowIndex As Long) As Boolean
    Dim Checksum As Integer
    Dim Packet(13) As Byte  '01 FILE_DOWNLOAD_REQ ID(4) RowsNumber(2) RowIndex(4) CheckSum 00
    Dim i As Integer
    Dim Stringa As String
    
    SubName = "SendPacketDownloadRequest"
    SendPacketDownloadRequest = False
    
    Packet(0) = SOF   'SOF=1
    Packet(1) = FILE_DOWNLOAD_REQ  'FILE_DOWNLOAD_REQ=21
    Stringa = Hex(ID)
    If Len(Stringa) = 1 Then Stringa = "000" + Stringa
    If Len(Stringa) = 2 Then Stringa = "00" + Stringa
    If Len(Stringa) = 3 Then Stringa = "0" + Stringa
    'Debug.Print Stringa
    'Stringa = SwapString(Stringa)
    Packet(2) = Asc(Right(Stringa, 1))
    Packet(3) = Asc(Mid(Stringa, 3, 1))
    Packet(4) = Asc(Mid(Stringa, 2, 1))
    Packet(5) = Asc(Left(Stringa, 1))
    Stringa = Hex(RowsNumber)
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    'Debug.Print Stringa
    'Stringa = SwapString(Stringa)
    Packet(6) = Asc(Right(Stringa, 1))
    Packet(7) = Asc(Left(Stringa, 1))
    Stringa = Hex(RowIndex)
    If Len(Stringa) = 1 Then Stringa = "000" + Stringa
    If Len(Stringa) = 2 Then Stringa = "00" + Stringa
    If Len(Stringa) = 3 Then Stringa = "0" + Stringa
    'Debug.Print Stringa
    'Stringa = SwapString(Stringa)
    Packet(8) = Asc(Right(Stringa, 1))
    Packet(9) = Asc(Mid(Stringa, 3, 1))
    Packet(10) = Asc(Mid(Stringa, 2, 1))
    Packet(11) = Asc(Left(Stringa, 1))
    Checksum = 0
    For i = 0 To 11
        Checksum = Checksum + Packet(i)
    Next
    Packet(12) = 256 - (Checksum Mod 256)
    Packet(13) = 0

'    For i = 0 To 13
'        Debug.Print Chr(Packet(i));
'    Next i
'    Debug.Print
    
    If fMain.MSComm1.PortOpen = True Then
        fMain.MSComm1.Output = Packet()
        SendPacketDownloadRequest = True
    Else
        SendPacketDownloadRequest = False
    End If


End Function


Public Function InputComTimeOut(TimeOut As Integer) As String
'Attende un input il cui terminatore e' vbLF
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As String
    Dim Dummy As String
    
    SubName = "InputComTimeOut"
    
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

        Dim Counter As Long
        
        SubName = "InputComTimeOutTerm"
        
        Counter = 0

        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
'            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until Dummy = Chr(Terminator) Or (Timer > TimeStop)
                DoEvents
                'Dummy = StrConv(fMain.MSComm1.Input, vbFromUnicode)
                Dummy = fMain.MSComm1.Input
                If Dummy <> "" Then
                    Linea = Linea + Dummy
                    Counter = Counter + 1
                End If
            Loop
        Else
            Linea = "TimeOut"
        End If
        Debug.Print "Counter->"; Counter
        Debug.Print "Lenght->"; Len(Linea)
        InputComTimeOutTerm = Linea

End Function
Public Function InputComTimeOutTerm2(TimeOut As Integer, Terminator As Byte) As String
'Attende un input il cui terminatore e' Terminator
'Input mode = Binary
'Con TIMEOUT

        Dim TimeStop As Long
        Dim Linea As String
        Dim Dummy As Variant
        Dim nDummy As Long
        Dim Counter As Long
        Dim i As Long
        
        SubName = "InputComTimeOutTerm2"
        Counter = 0

        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputMode = comInputModeBinary
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            'Dummy = 0
'            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until Timer > TimeStop
                DoEvents
                If fMain.MSComm1.InBufferCount >= 1 Then
                    nDummy = fMain.MSComm1.InBufferCount
                    Dummy = fMain.MSComm1.Input
                    Debug.Print "nDummy->"; nDummy; " ubound->"; UBound(Dummy); " dummy"; Dummy(0)
                    'For i = 0 To nDummy - 1
                        
                        'Linea = Linea + StrConv(Trim(Str((Dummy(0)))), vbFromUnicode)
                        Linea = Linea + Trim(Str(Dummy(0)))
                        Counter = Counter + 1
                        If Asc(Dummy(0)) = 0 Then GoTo uscita      'Orrendo ma forse efficace
                    'Next i
                End If
            Loop
        Else
            Linea = "TimeOut"
        End If
uscita:
        Debug.Print "Counter->"; Counter
        Debug.Print "Lenght->"; Len(Linea)
        fMain.MSComm1.InputMode = comInputModeText
        InputComTimeOutTerm2 = Linea

End Function
Public Function InputComTimeOutTerm3(TimeOut As Integer, Terminator As Byte) As String
'Attende un input il cui terminatore e' Terminator
'Input mode = Binary
'Con TIMEOUT

        Dim TimeStop As Long
        Dim Linea As String
        Dim Dummy As Variant
        Dim nDummy As Long
        Dim Counter As Long
        Dim i As Long
        
        Dim FileName As String
        Dim FileN As Long

        
        SubName = "InputComTimeOutTerm3"
        

        Counter = 0

        TimeStop = Timer + TimeOut
        'fMain.MSComm1.InputMode = comInputModeBinary
        fMain.MSComm1.InputLen = 0
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 300) Or (Timer > TimeStop)
        nDummy = fMain.MSComm1.InBufferCount
        Dummy = fMain.MSComm1.Input
        'Debug.Print "nDummy->"; nDummy; " ubound->"; UBound(Dummy); " dummy"; Dummy(0)

uscita:
        FileName = App.Path + "\pacchetto.txt"
        FileN = FreeFile
        Open FileName For Output As FileN
        Print #FileN, Dummy
        Close FileN
        FileName = App.Path + "\pacchetto2.txt"
        FileN = FreeFile
        Open FileName For Output As FileN
        For i = 0 To UBound(Dummy)
            Print #FileN, Dummy(i)
        Next
        Close FileN

        Debug.Print "Counter->"; Counter
        Debug.Print "Lenght->"; Len(Linea)
        fMain.MSComm1.InputMode = comInputModeText
        InputComTimeOutTerm3 = Linea

End Function

Public Function InputComTimeOutTerm4(TimeOut As Integer, Terminator As Byte) As String
'Attende un input il cui terminatore e' Terminator
'Input mode = Binary
'Con TIMEOUT

        Dim TimeStop As Long
        Dim Linea As String
'        Dim Dummy As Variant
        Dim nDummy As Long
        Dim Counter As Long
        Dim i As Long
        
        Dim Blocco() As Byte        'Blocco dati temporaneo in bytes
        Dim Buffer As Byte       'buffer temporaneo per i dati
        Dim BloccoDati() As Byte    'Blocco dati
        Dim iBloccoDati As Long     'Indice all'interno di BloccoDati()
        Dim DFPNT As Long           'Numero di bytes da scaricare
        Dim Bytes As Long           'Numero bytes scaricati
        Dim LungCounter As Long
        Dim Barra As Double
        Dim IncBarra As Double 'Incremento barra contatore per ogni riga
        Dim TimeOuts As Long        'Contatore dei Time Out
        Dim iDumm As Long
        Dim Dummy As String
        Dim Float As Single
        Dim LastByte As Byte
        Dim EndPacket As Boolean
        Dim Intero As Integer
        Dim dati As Long
        
        Dim FileName As String
        Dim FileN As Long


        SubName = "InputComTimeOutTerm4"

        fMain.MSComm1.InBufferCount = 0
'        fMain.MSComm1.InputLen = 1
        fMain.MSComm1.InputLen = 0
        fMain.MSComm1.RThreshold = 0
        fMain.MSComm1.InputMode = comInputModeBinary

        BloccoDati = ""
        TimeOuts = 0
        Bytes = 0
        Intero = 0
        dati = 0
        iBloccoDati = 0

        Counter = 0


        ReDim BloccoDati(700)

        EndPacket = False

        Do
            DoEvents
            TimeStop = Timer + 10
            Do
                DoEvents
            Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
            If fMain.MSComm1.InBufferCount = 0 Then
                TimeOuts = TimeOuts + 1
            Else
                dati = dati + fMain.MSComm1.InBufferCount
                Debug.Print "dati->"; dati
                Blocco = fMain.MSComm1.Input
                For i = LBound(Blocco) To UBound(Blocco)
                    BloccoDati(iBloccoDati) = Blocco(i)
                    iBloccoDati = iBloccoDati + 1
                    
                Next i
                Debug.Print "iBloccoDati->"; iBloccoDati
                TimeOuts = 0
                If BloccoDati(iBloccoDati - 1) = 0 Then
                    LastByte = 0
                    EndPacket = True
                    Debug.Print "Rilevato EOF"
                End If
            End If
            If TimeOuts > 13 Or iBloccoDati >= 700 Then Exit Do
            
            For i = 0 To iBloccoDati - 1
                Dummy = Dummy + Chr(BloccoDati(i))
            Next
            
            DoEvents
            
        Loop Until EndPacket = True

uscita:
        FileName = App.Path + "\pacchetto.txt"
        FileN = FreeFile
        Open FileName For Output As FileN
        Print #FileN, Dummy
        Close FileN
        FileName = App.Path + "\pacchetto2.txt"
        FileN = FreeFile
        Open FileName For Binary As FileN
            Put #FileN, , BloccoDati()
        Close FileN

'        Debug.Print "Counter->"; Counter
'        Debug.Print "Lenght->"; Len(Linea)
        fMain.MSComm1.InputMode = comInputModeText
        InputComTimeOutTerm4 = Dummy

End Function

Public Function StripZero(Stringa As String) As String
    SubName = "StripZero"
    StripZero = ""
    If Len(Stringa) >= 2 Then
        Stringa = Left(Stringa, Len(Stringa) - 1)
        StripZero = Stringa
    End If
End Function

Public Sub StampaAscii(Stringa As String)
    'Stampa il valore dei caratteri ASCII di una stringa
    'nella finestra di Debug
    Dim lStringa As Double
    Dim i As Integer
    
    SubName = "StampaAscii"
    
    lStringa = Len(Stringa)
    If lStringa = 0 Then Exit Sub
    Debug.Print "Risposta"; Stringa; " ";
    For i = 1 To lStringa
        If Mid(Stringa, i, 1) = Chr(0) Then
            Debug.Print "00";
        Else
            Debug.Print Asc(Mid(Stringa, i, 1));
        End If
    Next
    Debug.Print
End Sub

Public Function SwapString(Stringa As String) As String
    Dim lStringa As Long
    Dim Dummy As String
    Dim i As Long
    
    SubName = "SwapString"
    
    lStringa = Len(Stringa)
    'Capovolge la stringa
    Dummy = ""
    For i = lStringa To 1 Step -1
        Dummy = Dummy + Mid(Stringa, i, 1)
    Next
    SwapString = Dummy
End Function

Public Function String2ascii2(Stringa As String) As String
'Trasforma una stringa contenente caratteri ASCII e non
'ASCII in stringa di codici di caratteri ASCII
'Viene gestito anche il chr$(0)
    Dim lStringa As Integer
    Dim tStringa As String
    Dim i As Integer
    
    SubName = "String2ascii2"
    
    lStringa = Len(Stringa)
    For i = 1 To lStringa
        If Mid(Stringa, i, 1) = Chr$(0) Then
            tStringa = tStringa + " " + "00"
        Else
            tStringa = tStringa + Str(Asc(Mid(Stringa, i, 1)))
        End If
    Next
    String2ascii2 = tStringa
End Function

Public Function CallNumber(Number As String) As Boolean
    Dim i As Integer
    Dim Risposta As String
    Dim Tempo0 As Long
    Dim DiffTempo As Long
    Dim Contatore As Long
    Dim Msg As String
    Dim Messaggio As String
    
    SubName = "CallNumber"
    CallNumber = False

    'Setta la COM
    If fMain.MSComm1.PortOpen = True Then fMain.MSComm1.PortOpen = False
    fMain.MSComm1.CommPort = CommPort
    fMain.MSComm1.Settings = CommSettings
    fMain.MSComm1.Handshaking = comRTS
    fMain.MSComm1.RTSEnable = True
    fMain.lMonitor.Caption = "Setting the COM port"

    'Apre la COM
    fMain.MSComm1.PortOpen = True
    fMain.lMonitor.Caption = fMain.lMonitor.Caption + vbCrLf + "Opening COM port"
    
    'Pulisce il buffer
    fMain.MSComm1.InBufferCount = 0
    
'    'Fa la telefonata
    fMain.MSComm1.Output = "ATE0&D2" + vbCr   'Modem will interpret a DTR drop as a hang up command; auto answer doesn't work with this
    Risposta = UCase(InputComTimeOut(2))
    Debug.Print "AT&D2 ->"; Risposta
    Risposta = UCase(InputComTimeOut(2))
    Debug.Print "AT&D2 ->"; Risposta

    fMain.lMonitor.Caption = fMain.lMonitor.Caption + vbCrLf + Risposta + vbCrLf
    fMain.MSComm1.Output = "ATDT" + Number + vbCr
    fMain.lMonitor.Caption = fMain.lMonitor.Caption + vbCrLf + "Calling remote Modem" + vbCrLf
    Connected = False
    fMain.lMonitor.Caption = fMain.lMonitor.Caption & "Waiting for remote Modem" + vbCrLf
    
    Tempo0 = Timer
retry:
    Risposta = UCase(InputComTimeOut(10))
    Debug.Print "3 "; Risposta;
'    If fDebug Then Print #fdn, "3 "; Risposta
    
'    If ChiamaFlag = False Then
'        Debug.Print "ANNULLATO!"
'        Exit Sub
'    End If
    
    If Left(Risposta, 1) < " " Then GoTo retry
    If Risposta = "TIMEOUT" Then
        'controllo timeout
        If Timer < Contatore Then       '??????Contatore non è inizializzato!!! serve per gestire il passaggio di mezzanotte attulmente non implementato
            DiffTempo = Timer + 86400 - Tempo0
        Else
            DiffTempo = Timer - Tempo0
        End If
        Debug.Print DiffTempo
        GoTo retry
    End If
    'Risposta = Left(Risposta, Len(Risposta) - 2)
    Messaggio = Risposta
    Risposta = Left(Risposta, 4)

    Select Case Risposta
        Case "CONNECT"
            'MsgBox "Connect" + Risposta
            fMain.lMonitor.Caption = fMain.lMonitor.Caption + "Got answer from remote modem" + vbCrLf
            Connected = True
        Case "CONN"
            'MsgBox "Conn" + Risposta
            fMain.lMonitor.Caption = fMain.lMonitor.Caption + "Got answer from remote modem" + vbCrLf
            Connected = True
        Case "BUSY"
            MsgBox "Busy Line!!!" + Messaggio
            'GoTo Fallimento
        Case "NO CARRIER"
            MsgBox "Il modem remoto non ha risposto correttamente, nessuna portante" + Messaggio
            'GoTo Fallimento
        Case "NO C"
            MsgBox "No Carrier!" + Messaggio
            'GoTo Fallimento
        Case "NO DIALTONE"
            MsgBox "No Dial Tone!!" + Messaggio
            'GoTo Fallimento
        Case "NO D"
            MsgBox "No Dial Tone!!" + Messaggio
            'GoTo Fallimento
        Case "DELA"
            MsgBox "Modem in Delayed Mode!!!" + Messaggio
            'GoTo Fallimento
        Case "ATDT"
            GoTo retry
        Case Else
            MsgBox "Wrong answer from local Modem " + Messaggio
            GoTo retry
            'GoTo Fallimento
    End Select
    CallNumber = Connected

End Function

Function IntLittleE(strTwoChars As String) As Integer
    SubName = "IntLittleE"
    IntLittleE = Asc(Mid$(strTwoChars, 1, 1)) + Asc(Mid$(strTwoChars, 2, 1)) * 2 ^ 8
End Function
 
Function IntBigE(strTwoChars As String) As Integer
    SubName = "IntBigE"
    IntBigE = Asc(Mid$(strTwoChars, 1, 1)) * 2 ^ 8 + Asc(Mid$(strTwoChars, 2, 1))
End Function

Public Function TranscodeLine(Stringa As String) As String
    Dim i As Integer
    Dim Intero As Integer
    Dim FileN As Long
    Dim FileName As String
    Dim Carattere As String
    Dim Char1 As String
    Dim Char2 As String
    
    'FileName = App.Path + "\" + "File3.txt"
    'FileN = FreeFile
    'Open FileName For Output As #FileN

    TranscodeLine = ""
    For i = 1 To Len(Stringa) Step 2
        'TranscodeLine = TranscodeLine + Chr(Val(Mid(Stringa, i, 1)))
        'TranscodeLine = TranscodeLine + Chr(Val(Mid(Stringa, i, 2)))
        Char1 = Mid(Stringa, i, 1)
        Char2 = Mid(Stringa, i + 1, 1)
        Carattere = Char2 + Char1
        'Print #FileN, Carattere
        Carattere = HexToDecAscii(Carattere)
        'Print #FileN, Val(Carattere)
        Intero = Val(Carattere)
        'Print #FileN, Chr(Val(Mid(Stringa, i, 2)))
        'Print #FileN, Chr(Intero)
        TranscodeLine = TranscodeLine + Chr(Intero)
    Next i
    'Close #FileN
    'End
End Function
