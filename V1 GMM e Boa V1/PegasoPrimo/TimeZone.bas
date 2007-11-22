Attribute VB_Name = "TimeZone"
Option Explicit
Private Declare Function GetTimeZoneInformation _
   Lib "Kernel32" (lpTimeZoneInformation As _
   TIME_ZONE_INFORMATION) As Long

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2


Public Sub GetTimeZone(TimeZone As String, Bias As Single, tz2 As Long)
'Modificata, adesso restituisce anche l'informazione sull'ora estiva
   Dim nRet As Long
   Dim tz As TIME_ZONE_INFORMATION

   nRet = GetTimeZoneInformation(tz)
   tz2 = nRet
   If nRet <> TIME_ZONE_ID_INVALID Then
      Select Case nRet
         Case TIME_ZONE_ID_UNKNOWN
            'Me.Print "Time Zone Unknown!"
         Case TIME_ZONE_ID_STANDARD
            'Me.Print "Standard Time..."
         Case TIME_ZONE_ID_DAYLIGHT
            'Me.Print "Daylight Savings Time..."
      End Select
   
'      Me.Print "UTC Bias: "; tz.Bias / 60; " hrs."
'      Me.Print " ST Zone: "; TrimNull(CStr(tz.StandardName))
'      Me.Print " ST Date: "; tzDate(tz.StandardDate)
'      Me.Print " ST Bias: "; tz.StandardBias; " mins."
'      Me.Print " DT Zone: "; TrimNull(CStr(tz.DaylightName))
'      Me.Print " DT Date: "; tzDate(tz.DaylightDate)
'      Me.Print " DT Bias: "; tz.DaylightBias; " mins."
      TimeZone = TrimNull(CStr(tz.StandardName))
      Bias = tz.Bias / 60
   End If

    
End Sub
Private Function tzDate(st As SYSTEMTIME) As Date
   Dim i As Long
   Dim n As Long
   Dim d1 As Long
   Dim d2 As Long
   
   ' This member supports two date formats. Absolute format
   ' specifies an exact date and time when standard time
   ' begins. In this form, the wYear, wMonth, wDay, wHour,
   ' wMinute, wSecond, and wMilliseconds members of the
   ' SYSTEMTIME structure are used to specify an exact date.
   If st.wYear Then
      tzDate = _
         DateSerial(st.wYear, st.wMonth, st.wDay) + _
         TimeSerial(st.wHour, st.wMinute, st.wSecond)
   
   ' Day-in-month format is specified by setting the wYear
   ' member to zero, setting the wDayOfWeek member to an
   ' appropriate weekday, and using a wDay value in the
   ' range 1 through 5 to select the correct day in the
   ' month. Using this notation, the first Sunday in April
   ' can be specified, as can the last Thursday in October
   ' (5 is equal to "the last").
   Else
      ' Get first day of month
      d1 = DateSerial(Year(Now), st.wMonth, 1)
      ' Get last day of month
      d2 = DateSerial(Year(d1), st.wMonth + 1, 0)
      
      ' Match weekday with appropriate week...
      If st.wDay = 5 Then
         ' Work backwards
         For i = d2 To d1 Step -1
            If Weekday(i) = (st.wDayOfWeek + 1) Then
               Exit For
            End If
         Next i
      Else
         ' Start at 1st and work forward
         For i = d1 To d2
            If Weekday(i) = (st.wDayOfWeek + 1) Then
               n = n + 1  'incr week value
               If n = st.wDay Then
                  Exit For
               End If
            End If
         Next i
      End If
      
      ' Got the serial date!  Just format it and
      ' add in the appropriate time.
      tzDate = i + _
         TimeSerial(st.wHour, st.wMinute, st.wSecond)
   End If
End Function

Private Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim(StrIn)
   End Select
End Function




