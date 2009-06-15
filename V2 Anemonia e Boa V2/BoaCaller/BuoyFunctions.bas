Attribute VB_Name = "BuoyFunctions"
Option Explicit

Public Function ParseLine(Linea As String) As String
    Dim StringBuffer As String
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
    
    
    ParseLine = ""
    
    'Day
    StringBuffer = Left(Linea, 2)
    'Debug.Print "Stringbuffer="; StringBuffer
    StringBuffer = HexToDecAscii(StringBuffer)
    'Debug.Print StringBuffer
    DateDay = ZeroPad(StringBuffer)
    
    'Month
    StringBuffer = Mid(Linea, 3, 2)
    'Debug.Print "Stringbuffer="; StringBuffer
    StringBuffer = HexToDecAscii(StringBuffer)
    'Debug.Print StringBuffer
    DateMonth = ZeroPad(StringBuffer)
    
    'Year
    StringBuffer = Mid(Linea, 5, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    DateYear = "20" + ZeroPad(StringBuffer)
    
    'Hour
    StringBuffer = Mid(Linea, 7, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    DateHour = ZeroPad(StringBuffer)

    'Minute
    StringBuffer = Mid(Linea, 9, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    DateMin = ZeroPad(StringBuffer)
    
    'Seconds
    StringBuffer = Mid(Linea, 11, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    DateSec = ZeroPad(StringBuffer)

    'SPStemp
    StringBuffer = Mid(Linea, 13, 4)
    StringBuffer = HexToDecAscii(StringBuffer)
    SPStemp = Val(StringBuffer) / 10
    
    'BOAtemp
    StringBuffer = Mid(Linea, 17, 4)
    StringBuffer = HexToDecAscii(StringBuffer)
    BoaTemp = Val(StringBuffer) / 10
    
    'CPStemp
    StringBuffer = Mid(Linea, 21, 4)
    StringBuffer = HexToDecAscii(StringBuffer)
    CPSTemp = Val(StringBuffer) / 10
    
    'MeteoWindSpeed
    StringBuffer = Mid(Linea, 25, 4)
    StringBuffer = HexToDecAscii(StringBuffer)
    MeteoWindSpeed = Val(StringBuffer) * 1.852 / 10
    
    'MeteoWindDirection
    StringBuffer = Mid(Linea, 29, 4)
    StringBuffer = HexToDecAscii(StringBuffer)
    MeteoWindDirection = Val(StringBuffer) / 10
    
    'GPS_Lat
    StringBuffer = Mid(Linea, 33, 8)
    StringBuffer = HexToDecAscii(StringBuffer)
    GPS_Lat = Val(StringBuffer) / 10000000
    
    'GPS_LatDir
    StringBuffer = Mid(Linea, 41, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    GPS_LatDir = Val(StringBuffer)
    
    'GPS_Lon
    StringBuffer = Mid(Linea, 43, 8)
    StringBuffer = HexToDecAscii(StringBuffer)
    GPS_Lon = Val(StringBuffer) / 10000000
    
    'GPS_LonDir
    StringBuffer = Mid(Linea, 51, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    GPS_LonDir = Val(StringBuffer)
    
    'GPS_SatUsed
    StringBuffer = Mid(Linea, 53, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    GPS_LonDir = Val(StringBuffer)
   
    'MonitorBattery_3_12V
    StringBuffer = Mid(Linea, 55, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    MonitorBattery_3_12V = Val(StringBuffer) / 10
    
    'MonitorBattery_2_12V
    StringBuffer = Mid(Linea, 57, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    MonitorBattery_2_12V = Val(StringBuffer) / 10
    
    'MonitorBattery_1_12V
    StringBuffer = Mid(Linea, 59, 2)
    StringBuffer = HexToDecAscii(StringBuffer)
    MonitorBattery_1_12V = Val(StringBuffer) / 10
   
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

Public Function HexToDecAscii(Stringa As String) As String
    'Converte una stringa di Hex in Ascii in una stringa di decimali
    Dim LenghtStringa As Integer
    Dim Intero As Long
    LenghtStringa = Len(Stringa)
    Stringa = "&H" & Stringa
    Intero = Val(Stringa)
    HexToDecAscii = Trim(Str(Intero))
End Function

Public Function ZeroPad(Stringa As String) As String
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    If Len(Stringa) = 0 Then Stringa = "00"
    ZeroPad = Stringa
End Function
