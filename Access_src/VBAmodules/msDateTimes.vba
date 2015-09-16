Option Compare Database
Option Explicit

' Generate and handle date/time with millisecond accuracy in VBA.
'
' (c) 2006. Cactus Data ApS, CPH.
'
' Note:
' Dates in VBA follow a pseudo Gregorian calender from year 100 to
' the introduction of the Gregorian calendar in October 1582:
'
' http://www.ghgrb.ch/genealogicalIntroduction/kalender_greg_start_chron.html
'
'
' General note:
'   The numeric value of date
'     100-1-1 23:59:59.999
'   is lower than that of date
'     100-1-1 00:00:00.000
'
'
' SQL methods.
' To extract the millisecond of a date value and this rounded down to integer second
' using SQL only while sorting date values prior to 1899-12-30 correctly:
'
'   SELECT
'     [DateTimeMs],
'     Fix([DateTimeMs]*24*60*60)/(24*60*60) AS RoundSecSQL,
'     (([DateTimeMs]-Fix([DateTimeMs]))*24*60*60*1000)*Sgn([DateTimeMs]) Mod 1000 AS MsecSQL
'   FROM
'     tblTimeMsec
'   ORDER BY
'     Fix([DateTimeMs]),
'     Abs([DateTimeMs]);

  
Private Const cstrIntervalHour        As String = "h"
Private Const cstrIntervalMinute      As String = "n"
Private Const cstrIntervalSecond      As String = "s"
Private Const cstrIntervalMsec        As String = "ms"
Private Const cstrIntervalDsec        As String = "sms"

Private Const clngHoursPerDay         As Long = 24&
Private Const clngMinutesPerDay       As Long = 24& * 60&
Private Const clngSecondsPerDay       As Long = 24& * 60& * 60&
Private Const clngMillisecondsPerDay  As Long = 24& * 60& * 60& * 1000&

Private Const cdatDateMinAccess       As Date = #1/1/100#
Private Const cdatDateMinMySql        As Date = #1/1/1000#
Private Const cdatDateMinSqlServer    As Date = #1/1/1753#

Private Const TIME_ZONE_ID_UNKNOWN    As Long = &H0
Private Const TIME_ZONE_ID_STANDARD   As Long = &H1
Private Const TIME_ZONE_ID_DAYLIGHT   As Long = &H2
Private Const TIME_ZONE_ID_INVALID    As Long = &HFFFFFFFF

Private Type SYSTEMTIME
  wYear                               As Integer
  wMonth                              As Integer
  wDayOfWeek                          As Integer
  wDay                                As Integer
  wHour                               As Integer
  wMinute                             As Integer
  wSecond                             As Integer
  wMilliseconds                       As Integer
End Type
  
Private Type TIME_ZONE_INFORMATION
  Bias                                As Long
  StandardName(0 To (32 * 2 - 1))     As Byte   ' Unicode.
  StandardDate                        As SYSTEMTIME
  StandardBias                        As Long
  DaylightName(0 To (32 * 2 - 1))     As Byte   ' Unicode.
  DaylightDate                        As SYSTEMTIME
  DaylightBias                        As Long
End Type

' TIME_ZONE_INFORMATION-type variables hold information about the system's selected time zone.
' The two arrays in the structure are actually strings, each element holding the ASCII codes for
' each character (the end of the string is marked by a NULL character, ASCII code 0).
' For more information about how to convert the arrays into usable data, see the example for
' GetTimeZoneInformation.
'
' Bias
' The difference in minutes between UTC (a.k.a. GMT) time and local time.
' It satisfies the formula UTC time = local time + Bias.
'
' StandardName(0 To 31)
' Holds the name of the time zone for standard time.
'
' StandardDate
' The relative date for when daylight savings time ends.
'
' StandardBias
' A number to add to Bias to form the true bias during standard time.
'
' DaylightTime(0 To 31)
' Holds the name of the time zone for daylight savings time.
'
' DaylightDate
' The relative date for when daylight savings time begins.
'
' DaylightBias
' A number to add to Bias to form the true bias during daylight savings time.

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" ( _
  ByRef lpSystemTime As SYSTEMTIME)

Private Declare PtrSafe Function GetTimeZoneInformation Lib "Kernel32.dll" ( _
  ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) _
  As Long

Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" ( _
  ByVal uPeriod As Long) _
  As Long

Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () _
  As Long
'

Public Function DatePartMsec( _
  ByVal strInterval As String, _
  ByVal datDate As Date, _
  Optional ByVal bytFirstDayOfWeek As Byte = vbSunday, _
  Optional ByVal bytFirstWeekOfYear As Byte = vbFirstJan1) _
  As Double

' Extracts milliseconds or decimal seconds - as well as all
' the standard date/time parts - from datDate.
'
' strInterval "ms" will return millisecond of datDate.
' strInterval "sms" will return decimal seconds of datDate.
' Any other strInterval is handled by DatePart as usual.

  Dim dblPart               As Double
  
  On Error GoTo Err_DatePartMsec
  
  Select Case LCase(strInterval)
    Case cstrIntervalMsec
      dblPart = Millisecond(datDate)
    Case cstrIntervalDsec
      dblPart = SecondMsec(datDate)
    Case Else
      dblPart = DatePart(strInterval, datDate, bytFirstDayOfWeek, bytFirstWeekOfYear)
  End Select
  
  DatePartMsec = dblPart

Exit_DatePartMsec:
  Exit Function
  
Err_DatePartMsec:
  Call MsgErr(Err)
  Resume Exit_DatePartMsec

End Function

Public Function DateAddMsec( _
  ByVal strInterval As String, _
  ByVal dblNumber As Double, _
  ByVal datDate As Date) _
  As Date

' Adds milliseconds or decimal seconds - as well as
' all the standard date/time intervals - to datDate.
'
' strInterval "ms" will add milliseconds to datDate.
' Input range is any value between
'   -312,413,759,999,999 and 312,413,759,999,999
' that will result in a valid date value between
'   100-1-1 00:00:00.000 and 9999-12-31 23:59:59.999
'
' strInterval "sms" will add decimal seconds to datDate.
' Input range is any value between
'   x and y
' that will result in a valid date value between
'   100-1-1 00:00:00.000 and 9999-12-31 23:59:59.999
'
' Any other strInterval is handled by DateAdd as usual.

  Dim datNext               As Date
  Dim dblMilliseconds       As Double
  
  On Error GoTo Err_DateAddMsec
  
  Select Case LCase(strInterval)
    Case cstrIntervalMsec
      datNext = MsecSerial(dblNumber, datDate)
    Case cstrIntervalDsec
      datNext = MsecSerial(dblNumber * 1000, datDate)
    Case Else
      datNext = DateAddFull(strInterval, dblNumber, datDate)
  End Select
  
  DateAddMsec = datNext

Exit_DateAddMsec:
  Exit Function
  
Err_DateAddMsec:
  Call MsgErr(Err)
  Resume Exit_DateAddMsec

End Function

Public Function DateDiffMsec( _
  ByVal strInterval As String, _
  ByVal datDate1 As Date, _
  ByVal datDate2 As Date, _
  Optional ByVal bytFirstDayOfWeek As Byte = vbSunday, _
  Optional ByVal bytFirstWeekOfYear As Byte = vbFirstJan1) _
  As Double

' Will calculate difference in milliseconds or decimal seconds from
' datDate1 to datDate2 as well as all the standard date/time diffs.
'
' datDate1 and datDate2 can be any valid date value between
'   100-1-1 00:00:00.000 and 9999-12-31 23:59:59.999
'
' strInterval "ms" will return the difference in milliseconds.
' strInterval "sms" will return the difference in decimal seconds.
' Any other strInterval is handled by DateDiff as usual.
  
  Dim dblDiff               As Double
  
  On Error GoTo Err_DateDiffMsec
  
  Select Case LCase(strInterval)
    Case cstrIntervalMsec
      dblDiff = MsecDiff(datDate1, datDate2)
    Case cstrIntervalDsec
      dblDiff = MsecDiff(datDate1, datDate2) / 1000
    Case Else
      dblDiff = DateDiff(strInterval, datDate1, datDate2, bytFirstDayOfWeek, bytFirstWeekOfYear)
  End Select
  
  DateDiffMsec = dblDiff
  
Exit_DateDiffMsec:
  Exit Function
  
Err_DateDiffMsec:
  Call MsgErr(Err)
  Resume Exit_DateDiffMsec
  
End Function

Public Function MsecDiff( _
  ByVal datDate1 As Date, _
  ByVal datDate2 As Date) _
  As Double
  
' Returns the difference in milliseconds between datDate1 and datDate2.
' Accepts any valid Date value including milliseconds from
'   100-01-01 00:00:00.000 to 9999-12-31 23:59:59.999 or reverse
' which will return from
'   -312,413,759,999,999 to 312,413,759,999,999

  Dim dblMsecDiff As Double
  
  ' Convert native date values to linear date values.
  Call ConvDateToLinear(datDate1)
  Call ConvDateToLinear(datDate2)
  ' Convert to milliseconds and find the difference.
  dblMsecDiff = CDec(datDate2 * clngMillisecondsPerDay) - CDec(datDate1 * clngMillisecondsPerDay)
  
  MsecDiff = dblMsecDiff
  
End Function

Public Function DateSort( _
  ByVal datDate As Date) _
  As Double

' Returns a continous value for datDate including milliseconds
' that can be sorted on correctly even for negative date values.
' Return values to sort on span from
'   0 for
'     100-01-01 00:00:00.000
' to
'   312,413,759,999,999 for
'    9999-12-31 23:59:59.999

  Dim dblSort As Double
  
  dblSort = MsecDiff(cdatDateMinAccess, datDate)
  
  DateSort = dblSort
    
End Function

Public Function Msec( _
  Optional ByVal intTimePart As Integer) _
  As Date

' This is the core function.
' It generates the current time with millisecond resolution.
'
' Returns current (local) date/time including millisecond.
' Parameter intTimePart determines level of returned value:
'   0: Millisecond value only.
'   1: Time value only including milliseconds.
'   2: Full Date/time value including milliseconds.
'   None or any other value: Millisecond value only.

  Const cintMsecOnly            As Integer = 0
  Const cintMsecTime            As Integer = 1
  Const cintMsecDate            As Integer = 2
  
  Static typTime      As SYSTEMTIME
  Static lngMsecInit  As Long

  Dim datMsec         As Date
  Dim datDate         As Date
  Dim intMilliseconds As Integer
  Dim lngTimeZoneBias As Long
  Dim lngMsec         As Long
  Dim lngMsecCurrent  As Long
  Dim lngMsecOffset   As Long
  
  ' Set resolution of timer to 1 ms.
  timeBeginPeriod 1
  lngMsecCurrent = timeGetTime()
  
  If lngMsecInit = 0 Or lngMsecCurrent < lngMsecInit Then
    ' Initialize.
    ' Get bias for local time zone respecting
    ' current setting for daylight savings.
    lngTimeZoneBias = GetLocalTimeZoneBias(False)
    ' Get current UTC system time.
    Call GetSystemTime(typTime)
    intMilliseconds = typTime.wMilliseconds
    ' Repeat until GetSystemTime retrieves next count of milliseconds.
    ' Then retrieve and store count of milliseconds from launch.
    Do
      Call GetSystemTime(typTime)
    Loop Until typTime.wMilliseconds <> intMilliseconds
    lngMsecInit = timeGetTime()
    ' Adjust UTC to local system time by correcting for time zone bias.
    typTime.wMinute = typTime.wMinute - lngTimeZoneBias
    ' Note: typTime may now contain an invalid (zero or negative) minute count.
    ' However, the minute count is acceptable by TimeSerial().
  Else
    ' Retrieve offset from initial time to current time.
    lngMsecOffset = lngMsecCurrent - lngMsecInit
  End If
  
  With typTime
    ' Now, current system time is initial system time corrected for
    ' time zone bias.
    lngMsec = (.wMilliseconds + lngMsecOffset)
    Select Case intTimePart
      Case cintMsecTime, cintMsecDate
        ' Calculate the time to add as a date/time value with millisecond resolution.
        datMsec = lngMsec / 1000 / clngSecondsPerDay
        ' Add to this the current system time.
        datDate = datMsec + TimeSerial(.wHour, .wMinute, .wSecond)
        If intTimePart = cintMsecDate Then
          ' Add to this the current system date.
          datDate = datDate + DateSerial(.wYear, .wMonth, .wDay)
        End If
      Case Else
        ' Calculate millisecond part as a date/time value with millisecond resolution.
        datMsec = (lngMsec Mod 1000) / 1000 / clngSecondsPerDay
        ' Return millisecond part only.
        datDate = datMsec
    End Select
  End With

  Msec = datDate
  
End Function

Public Function TimeMsec() As Date

  Const cintMsecTime            As Integer = 1

  Dim datTime As Date
  
  datTime = Msec(cintMsecTime)
  
  TimeMsec = datTime
  
End Function

Public Function NowMsec() As Date
  
  Const cintMsecDate            As Integer = 2

  Dim datTime As Date
  
  datTime = Msec(cintMsecDate)
  
  NowMsec = datTime

End Function

Public Function TimerMsec() As Double

' Returns count of seconds from Midnight with millisecond resolution.
' Mimics Timer which returns the value with a resolution of 1/64 second.

  Const cintMsecTime            As Integer = 1

  Dim dblSeconds  As Double
  
  dblSeconds = Msec(cintMsecTime) * clngSecondsPerDay

  TimerMsec = dblSeconds

End Function

Public Function Millisecond( _
  ByVal datDate As Date) _
  As Integer

' Returns the millisecond part from datDate.

  Dim dblTime                   As Double
  Dim intMillisecond            As Integer
  
  dblTime = CDbl(datDate)
  ' Remove date part from date/time value and extract count of milliseconds.
  ' Note the use of CDec() to prevent bit errors for very large date values.
  intMillisecond = Abs(dblTime - CDec(Fix(dblTime))) * clngMillisecondsPerDay Mod 1000
  
  Millisecond = intMillisecond

End Function

Public Function SecondMsec( _
  ByVal datDate As Date) _
  As Double

' Returns the second and the millisecond from datDate as a decimal value.

  Dim intSecond           As Integer
  Dim intMillisecond      As Integer
  Dim dblSecondMsec       As Double

  ' Get milliseconds of datDate.
  intMillisecond = Millisecond(datDate)
  ' Round off datDate to the second.
  Call RoundSecondOff(datDate)
  ' Get the rounded count of seconds.
  intSecond = Second(datDate)
  ' Calculate seconds and milliseconds as decimal seconds.
  dblSecondMsec = intSecond + intMillisecond / 1000
  
  SecondMsec = dblSecondMsec
  
End Function

Public Function MsecSerial( _
  ByVal dblMillisecond As Double, _
  Optional ByVal datBase As Date) _
  As Date
  
' Returns the date/time value of dblMillisecond rounded to integer milliseconds.
' Typical usage:
'   datMsec = MsecSerial(milliseconds)
'   datDateTimeMsec = MsecSerial(milliseconds, datDateTime)
'
' Values of dblMillisecond beyond +/-999 are still converted to valid date values.
' Accepts, with no datBase, any input in the interval:
'   -56,802,297,600,000 to 255,611,462,399,999
' Possible return value is any Date value from:
'   100-1-1 00:00:00.000 to 9999-12-31 23:59:59.999
'
' If a datBase is specified, dblMillisecond will be added to this, and
' the acceptable input range is shifted accordingly.
' Min. datBase:
'    100-01-01 00:00:00.000. Will accept dblMillisecond of
'   0 to 312,413,759,999,999
' Max. datBase:
'   9999-12-31 23:59:59.999. Will accept dblMillisecond of
'   -312.413.759.999.999 to 0
' Resulting return dates must be within the limits of datatype Date:
'   100-1-1 00:00:00.000 to 9999-12-31 23:59:59.999

  Dim datDate       As Date
  Dim dblDate       As Double
  
  On Error GoTo Err_MsecSerial
  
  ' Convert (invalid) numeric negative date values less than one day.
  Call ValidateDate(datBase)
  
  If dblMillisecond = 0 Then
    ' Nothing to add. Return base date.
    datDate = datBase
  Else
    ' Convert datBase to linear date value.
    Call ConvDateToLinear(datBase)
    ' Convert datbase to milliseconds and adjust with dblMilliseconds.
    dblDate = CDbl(datBase * clngMillisecondsPerDay) + CDec(Fix(dblMillisecond))
    ' Convert milliseconds to linear date value.
    datDate = CVDate(dblDate / clngMillisecondsPerDay)
    ' Convert datBase to native date value.
    Call ConvDateToNative(datDate)
  End If
      
  MsecSerial = datDate
  
Exit_MsecSerial:
  Exit Function
  
Err_MsecSerial:
  Call MsgErr(Err)
  Resume Exit_MsecSerial
  
End Function

Public Function TimeSerialMsec( _
  ByVal intHour As Integer, _
  ByVal intMinute As Integer, _
  ByVal dblSecond As Double, _
  Optional ByVal intMillisecond As Integer) _
  As Date

' Returns the date/time value of the combined parameters for
' hour, minute, second and millisecond.
' Accepts decimal input for seconds.
' The fraction of second is rounded to integer milliseconds.
'
' If input values for second or millisecond beyond those of datatype
' Integer are expected, use MsecSerial() or DateAddMsec().

  Dim datTime         As Date
  Dim intSecond       As Integer
  Dim lngMillisecond  As Long
  
  On Error GoTo Err_TimeSerialMsec
  
  ' Raise error if integer part of second exceeds Integer datatype.
  intSecond = Int(dblSecond)
  ' Convert fraction of second to milliseconds and add rounded intMillisecond.
  lngMillisecond = Int((CDec(dblSecond) - intSecond) * 1000 + CDec(0.5)) + intMillisecond
  datTime = MsecSerial(lngMillisecond, TimeSerialFull(intHour, intMinute, intSecond))

  TimeSerialMsec = datTime

Exit_TimeSerialMsec:
  Exit Function
  
Err_TimeSerialMsec:
  Call MsgErr(Err)
  Resume Exit_TimeSerialMsec

End Function

Public Function MsecValueMsec( _
  ByVal strTime As String) _
  As Date
  
' Wrapper for ExtractMsec.
' Passes strTime by value and acts as sister function for:
'   DateValueMsec and
'   TimeValueMsec

  Dim datMsec As Date

  datMsec = ExtractMsec((strTime))
  
  MsecValueMsec = datMsec
  
End Function

Public Function ExtractMsec( _
  ByRef strTime As String) _
  As Date

' Returns millisecond date/time value from the last digits of a strTime.
'
' Note:
'   Returns ByRef strTime without millisecond part.
'   To pass strTime ByVal, call the function like this:
'     datMsec = ExtractMsec((strTime))
'   or use MsecValueMsec(strTime).
'
' Examples:
'   "01:13"
'     Returns 0 milliseconds
'   "09:25.17"
'     Returns 0 milliseconds
'   "11:45:27"
'     Returns 0 milliseconds
'   "08:33:12 AM 60"
'     Returns 60 milliseconds
'   "18:23:22.322"
'     Returns 322 milliseconds
'   "18:23:22.322 ms"
'     Returns 322 milliseconds
'   "18:23:22-322"
'     Returns 322 milliseconds
'   "08:33:42.87391"
'     Returns 873 milliseconds
'   "08:33:42.87.391"
'     Returns 391 milliseconds

  Dim datMsec As Date
  Dim strMsec As String
  
  strTime = Trim(strTime)
  If IsDate(strTime) Then
    datMsec = MsecSerial(0)
    ' strTime represents a valid time, thus it contains no milliseconds.
  Else
    ' Clean strTime from trailing non-numeric chars.
    Do While Not IsNumeric(Right(strTime, 1)) And Len(strTime) > 0
      strTime = Left(strTime, Len(strTime) - 1)
    Loop
    ' Extract and convert millisecond part of strTime.
    Do While IsNumeric(Right(strTime, 1))
      strMsec = Right(strTime, 1) & strMsec
      strTime = Left(strTime, Len(strTime) - 1)
    Loop
    ' Clean strTime from trailing non-numeric chars.
    Do While Not IsNumeric(Right(strTime, 1)) And Len(strTime) > 0
      strTime = Left(strTime, Len(strTime) - 1)
    Loop
    ' Ignore minus sign and more than three digits.
    datMsec = MsecSerial(Val(Left(strMsec, 3)))
  End If
  
  ExtractMsec = datMsec

  End Function

Public Function CDateMsec( _
  ByVal varDate As Variant) _
  As Date

' Converts a date/time expression including millisecond
' to a date/time value.
' Note: Will raise error 13 if the cleaned varDate
' does not represent a valid date expression.
'
' For converting pure numerics, use CVDateMsec().
  
  Dim datDate As Date
  Dim datMsec As Date
  Dim strDate As String
  
  On Error GoTo Err_CDateMsec

  ' First try IsDate(varDate). If success, simply use CDate(varDate).
  If IsDate(varDate) Then
    ' Convert varDate to a date value.
    datDate = CDate(varDate)
  Else
    ' Try to convert varDate to a String and strip a millisecond part.
    strDate = CStr(varDate)
    datMsec = ExtractMsec(strDate)
    ' ExtractMsec returned a cleaned strDate.
    ' Convert strDate to Date value and add the milliseconds.
'    datDate = DateNative(DateLinear(CDate(strDate)) + DateLinear(datMsec))
    datDate = DateLinear(CDate(strDate)) + DateLinear(datMsec)
  
  End If
  
  CDateMsec = datDate

Exit_CDateMsec:
  Exit Function
  
Err_CDateMsec:
  Call MsgErr(Err)
  Resume Exit_CDateMsec
  
End Function

Public Function CVDateMsec( _
  ByVal varDate As Variant) _
  As Variant

' Converts a date/time expression including millisecond or
' a number to a date/time value.
' Null is accepted and is returned as Null.
' Note: Will raise error 13 if the cleaned varDate is not Null and
' does not represent a valid date value.

  Dim varTest As Variant
  
  If Not IsNull(varDate) Then
    If IsNumeric(varDate) Then
      varTest = CVDate(varDate)
    Else
      varTest = CDateMsec(varDate)
    End If
  End If
  
  CVDateMsec = varTest

End Function

Public Function IsDateMsec( _
  ByVal varDate As Variant) _
  As Boolean
  
  Dim strDate As String
  Dim booDate As Boolean
  
' Checks an expression if it represents a date/time value
' with or without a millisecond part.

  On Error GoTo Err_IsDateMsec
  
  ' First try IsDate(varDate). If success, we are done.
  booDate = IsDate(varDate)
  If booDate = False And Not IsNull(varDate) Then
    ' Try to convert varDate to a String and strip a millisecond part.
    strDate = CStr(varDate)
    Call ExtractMsec(strDate)
    ' ExtractMsec returned a cleaned strDate.
    ' Validate this.
    booDate = IsDate(strDate)
  End If

  IsDateMsec = booDate
  
Exit_IsDateMsec:
  Exit Function

Err_IsDateMsec:
  Err.Clear
  Resume Exit_IsDateMsec

End Function

Public Function DateValueMsec( _
  ByVal strDate As String) _
  As Date
  
' Cleans strDate for a time part.
' Returns the date value of strDate.
'
' Note:
'   Will raise error 13 if the cleaned strDate
'   does not represent a valid date value.

  Dim datDate As Date
  
  On Error GoTo Err_DateValueMsec
  
  ' Strip a millisecond part from strDate.
  Call ExtractMsec(strDate)
  ' Convert strDate to date value with no time part.
  datDate = DateValue(strDate)

  DateValueMsec = datDate

Exit_DateValueMsec:
  Exit Function
  
Err_DateValueMsec:
  Call MsgErr(Err)
  Resume Exit_DateValueMsec
  
End Function

Public Function TimeValueMsec( _
  ByVal strTime As String) _
  As Date
  
' Cleans strTime for a millisecond part.
' Returns the time value of strTime excluding millisecond.
'
' Note:
'   Will raise error 13 if the cleaned strTime
'   does not represent a valid time value.
'
' To retrieve time part including millisecond use CDateMsec().

  Dim datTime As Date
  
  On Error GoTo Err_TimeValueMsec
  
  ' Clean strTime for a millisecond part.
  Call ExtractMsec(strTime)
  ' Convert cleaned strTime to time value.
  datTime = TimeValue(strTime)
  
  TimeValueMsec = datTime

Exit_TimeValueMsec:
  Exit Function
  
Err_TimeValueMsec:
  Call MsgErr(Err)
  Resume Exit_TimeValueMsec
  
End Function

Public Function GetLocalTimeZoneBias( _
  Optional ByVal booIgnoreDaylightSetting As Boolean) _
  As Long
 
  Dim tzi           As TIME_ZONE_INFORMATION
  Dim lngTimeZoneID As Long
  Dim lngBias       As Long
  
  lngTimeZoneID = GetTimeZoneInformation(tzi)
  
  Select Case lngTimeZoneID
    Case TIME_ZONE_ID_STANDARD, TIME_ZONE_ID_DAYLIGHT
      lngBias = tzi.Bias
      If lngTimeZoneID = TIME_ZONE_ID_DAYLIGHT Then
        If booIgnoreDaylightSetting = False Then
          lngBias = lngBias + tzi.DaylightBias
        End If
      End If
  End Select
  
  GetLocalTimeZoneBias = lngBias
   
End Function

Public Function DateTimeRound( _
  ByVal datTime As Date) _
  As Date
  
' Returns datTime rounded off to the second by
' removing a millisecond portion.
  
  Call RoundSecondOff(datTime)
  
  DateTimeRound = datTime
  
  End Function

Public Function DateTimeRoundMsec( _
  ByVal datDate As Date, _
  Optional booRoundSqlServer As Boolean, _
  Optional booRoundSecondUp As Boolean) _
  As Date
  
' Returns datDate rounded to the nearest millisecond approximately by 4/5.
' The dividing point for up/down rounding may vary between 0.3 and 0.7ms
' due to the limited resolution of data type Double.
'
' If booRoundSqlServer is True, milliseconds are rounded by 3.333ms to match
' the rounding of the datetime data type of SQL Server - to 0, 3 or 7 as the
' least significant digit:
'
' Msec SqlServer
'   0    0
'   1    0
'   2    3
'   3    3
'   4    3
'   5    7
'   6    7
'   7    7
'   8    7
'   9   10
'  10   10
'  11   10
'  12   13
'  13   13
'  14   13
'  15   17
'  16   17
'  17   17
'  18   17
'  19   20
' ...
' 990  990
' 991  990
' 992  993
' 993  993
' 994  993
' 995  997
' 996  997
' 997  997
' 998  997
' 999 1000
'
' If booRoundSqlServer is True and if booRoundSecondUp is True, 999ms
' will be rounded up to 1000ms - the next second - which may not be
' what you wish. If booRoundSecondUp is False, 999ms will be rounded
' down to 997ms:
'
' 994  993
' 995  997
' 996  997
' 997  997
' 998  997
' 999  997
'
' If booRoundSqlServer is False, booRoundSecondUp is ignored.
  
  Dim intMsec As Integer
  Dim datMsec As Date
  Dim datTemp As Date
  
  ' Retrieve count of milliseconds.
  intMsec = Millisecond(datDate)
  If booRoundSqlServer = True Then
    ' Perform special rounding to match data type datetime of SQL Server.
    intMsec = (intMsec \ 10) * 10 + Choose(intMsec Mod 10 + 1, 0, 0, 3, 3, 3, 7, 7, 7, 7, 10)
    If booRoundSecondUp = False Then
      If intMsec = 1000 Then
        intMsec = 997
      End If
    End If
  End If
  ' Round datdate to the second.
  Call RoundSecondOff(datDate)
  ' Get milliseconds as date value.
  datMsec = MsecSerial(intMsec)
  ' Add milliseconds to rounded date.
  datTemp = DateNative(DateLinear(datDate) + DateLinear(datMsec))
  
  DateTimeRoundMsec = datTemp
  
End Function

Private Sub RoundSecondOff( _
  ByRef datDate As Date)
  
' Rounds off datDate to the second by
' removing a millisecond portion.

  Dim lngDate             As Long
  Dim lngTime             As Long
  Dim dblTime             As Double
  
  ' Get date part.
  lngDate = Fix(datDate)
  ' Get time part.
  dblTime = datDate - lngDate
  ' Round time part to the second.
  lngTime = Fix(dblTime * clngSecondsPerDay)
  ' Return date part and rounded time part.
  datDate = CVDate(lngDate + lngTime / clngSecondsPerDay)
  
End Sub
  
Public Function MsgErr( _
  ByVal errObj As ErrObject) _
  As Integer
  
' Calls Microsoft Access styled MsgBox to display runtime error.

  Const cstrErrHeader As String = "Runtime Error"
  
  Dim strErrNumber  As String
  Dim strMsgPrompt  As String
  Dim lngMsgStyle   As Long
  Dim intMsgResult  As Integer
  
  With errObj
    strErrNumber = " " & Chr(34) & .Number & Chr(34) & ":"
    strMsgPrompt = cstrErrHeader & strErrNumber & String(2, vbCrLf) & .Description
    lngMsgStyle = vbExclamation + vbOKOnly
  End With
  intMsgResult = MsgBox(strMsgPrompt, lngMsgStyle)

  MsgErr = intMsgResult

End Function

Public Function StrDateFull( _
  ByVal datDate As Date) _
  As String
  
' Returns datDate formatted as to the current settings for
' "Short Date" and "Long Time" with trailing milliseconds
' formatted as fixed length numeric string.
'
' Example:
'   10-12-2007 11:48:23.010

  ' Separators.
  Const cstrSeparatorDateTime As String = " "
  Const cstrSeparatorTimeMsec As String = "."
  
  Dim strDate   As String
  Dim strMsec   As String
  
  ' Build millisecond string.
  strMsec = StrDateMsec(datDate)
  ' Round off to seconds.
  Call RoundSecondOff(datDate)
  strDate = Format(datDate, "Short Date") & cstrSeparatorDateTime & _
    Format(datDate, "Long Time") & cstrSeparatorTimeMsec & strMsec
  
  StrDateFull = strDate
  
End Function

Public Function StrDateMsec( _
  ByVal datDate As Date, _
  Optional ByVal strFormat As String) _
  As String
  
' Returns milliseconds of datDate formatted using strFormat.
' Default format is fixed length numeric string.

  ' Default format for milliseconds.
  Const cstrFormat  As String = "000"
  
  Dim strMsec   As String
  
  If Len(strFormat) = 0 Then
    ' Apply default format.
    strFormat = cstrFormat
  End If
  strMsec = Format(Millisecond(datDate), strFormat)
  
  StrDateMsec = strMsec
  
End Function

Public Function FormatMsec( _
  ByVal varExpression As Variant, _
  ByVal strFormat As String, _
  Optional ByVal strFormatMsec As String, _
  Optional ByVal bytFirstDayOfWeek As Byte = vbSunday, _
  Optional ByVal bytFirstWeekOfYear As Byte = vbFirstJan1) _
  As Variant
  
' Returns, if possible, a value as a string formatted with milliseconds
' or any other value using Format().
' Accepts a separate subformat for milliseconds.
'
' A special placeholder for milliseconds specifies where milliseconds
' should be inserted in the formatted string.
'
' Note:
'   If varExpression contains milliseconds but no millisecond part
'   is requested, milliseconds will be 4/5 rounded to the second.
'
' Typical usage.
'   Return ISO-8601 formatted date/time string:
'     FormatMsec(NowMsec, "yyyy\-mm\-dd hh\:nn\:ssi")
'       2007-03-27 13:48:32.017
'   Return US style formatted date/time string:
'     FormatMsec(NowMsec, "m\/d\/yyyy h\:nn\:ss AM/PM i", "0ms")
'       3/27/2007 1:48:32 PM 17ms
'   Rounding date/time with milliseconds to seconds:
'     FormatMsec(NowMsec, "yyyy\-mm\-dd hh\:nn\:ss")
'       2007-03-27 13:48:32
  
  ' Placeholder in format string for millisecond part.
  Const cstrPlaceholderMsec As String = "i"
  ' Escape character.
  Const cstrCharEscape      As String = "\"
  ' Default format for milliseconds.
  Const cstrFormatMsec      As String = cstrCharEscape & ".000"
  
  Dim strText     As String
  Dim strChar     As String
  Dim strForm1    As String
  Dim strForm2    As String
  Dim intPos      As Integer
  Dim intMid      As Integer
  Dim intLen      As Integer
  Dim datDate     As Date
    
  intLen = Len(strFormat)
  ' Loop through strFormat to locate a possible placeholder for milliseconds
  ' skipping any character with a preceding escape character.
  Do While intPos = 0 And intMid < intLen
    intMid = intMid + 1
    strChar = Mid(strFormat, intMid, 1)
    If StrComp(strChar, cstrCharEscape, vbTextCompare) = 0 Then
      ' Skip (escape) next character.
      intMid = intMid + 1
    ElseIf StrComp(strChar, cstrPlaceholderMsec, vbTextCompare) = 0 Then
      ' Placeholder for millisecond located.
      intPos = intMid
    End If
  Loop
  
  If intPos > 0 And IsDateMsec(varExpression) Then
    ' Retrieve and format millisecond part of varExpression.
    ' Convert varExpression to a Date value.
    datDate = CDateMsec(varExpression)
    If Len(strFormatMsec) = 0 Then
      strFormatMsec = cstrFormatMsec
    End If
    strText = StrDateMsec(datDate, strFormatMsec)
    ' Round down seconds of varExpression.
    varExpression = DateTimeRound(datDate)
    ' Split format string.
    strForm1 = Mid(strFormat, 1, intPos - 1)
    strForm2 = Mid(strFormat, intPos + 1)
    ' Build and concatenate formatted string.
    If Len(strForm1) > 0 Then
      strText = Format(varExpression, strForm1, bytFirstDayOfWeek, bytFirstWeekOfYear) & strText
    End If
    If Len(strForm2) > 0 Then
      strText = strText & Format(varExpression, strForm2, bytFirstDayOfWeek, bytFirstWeekOfYear)
    End If
  Else
    ' No request for milliseconds; or no date value or no format string passed.
    ' Use standard formatting. This will round milliseconds to seconds for a date expression.
    strText = Format(varExpression, strFormat, bytFirstDayOfWeek, bytFirstWeekOfYear)
  End If
  
  FormatMsec = strText
  
End Function

Public Function StrDateIso8601Msec( _
  ByVal varExpression As Variant) _
  As String
  
' Returns, if possible, a value as a string formatted with milliseconds
' according to ISO-8601.
'
' Typical usage.
'   StrDateIso8601Msec(NowMsec)
'     2007-03-27 13:48:32.017

  Const cstrFormatIso8601 As String = "yyyy\-mm\-dd hh\:nn\:ssi"
  
  Dim strDate             As String
  
  If IsDateMsec(varExpression) Then
    strDate = FormatMsec(varExpression, cstrFormatIso8601)
  End If
  
  StrDateIso8601Msec = strDate
  
End Function

Public Function DateSumDates( _
  ByVal datDate1 As Date, _
  ByVal datDate2 As Date, _
  Optional ByVal datDate3 As Date) _
  As Date

' Adds/subtracts partial date values to create a compound date value.
'
' Typical usage:
'   datCompound = DateSumDates(datDate, datTime, datMsec)

  Dim decDate       As Variant
  Dim datDate       As Date
  
  Call ConvDateToLinear(datDate1)
  Call ConvDateToLinear(datDate2)
  Call ConvDateToLinear(datDate3)
  decDate = CDec(datDate1) + CDec(datDate2) + CDec(datDate3)
  datDate = DateNative(CVDate(decDate))
  
  DateSumDates = datDate
  
End Function
  
Public Function TimeSerialFull( _
  ByVal intHour As Integer, _
  ByVal intMinute As Integer, _
  ByVal intSecond As Integer) _
  As Date
  
  Dim datTime As Date
  Dim dblDate As Double
  Dim dblTime As Double
  
' Returns correct numeric negative date values,
' which TimeSerial() does not.
' This applies to Access 2003 and below.
' Not known if needed for Access 12/2007 as well.
'
' 2006-04-23. Cactus Data ApS, CPH.
'
' Example sequence:
'   Numeric      Date     Time
'   2            19000101 0000
'   1.75         18991231 1800
'   1.5          18991231 1200
'   1.25         18991231 0600
'   1            18991231 0000
'   0.75         18991230 1800
'   0.5          18991230 1200
'   0.25         18991230 0600
'   0            18991230 0000
'  -1.75         18991229 1800
'  -1.5          18991229 1200
'  -1.25         18991229 0600
'  -1            18991229 0000
'  -2.75         18991228 1800
'  -2.5          18991228 1200
'  -2.25         18991228 0600
'  -2            18991228 0000

  datTime = TimeSerial(intHour, intMinute, intSecond)
  If datTime < 0 Then
    ' Get date (integer) part of datTime shifted one day
    ' if a time part is present as Int() rounds down.
    dblDate = Int(datTime)
    ' Retrieve and reverse time (decimal) part.
    dblTime = dblDate - datTime
    ' Assemble and convert date and time part.
    datTime = CVDate(dblDate + dblTime)
  End If
    
  TimeSerialFull = datTime

End Function

Public Function DateAddFull( _
  ByVal strInterval As String, _
  ByVal lngNumber As Long, _
  ByVal datDate As Date) _
  As Date

' This is for Access 97 only in which DateAdd() is buggy.
' Does not return invalid date values between -1 and 0.
' With Access 2000 and newer, DateAdd() can be used as is.

  ' Version major of Access 97.
  Const cbytVersionMajorMax     As Byte = 8
  
  ' Store current version of Access.
  Static bytVersionMajor        As Byte
  
  Dim datNext                   As Date
  Dim lngFactor                 As Long
  Dim dblMilliseconds           As Double
  
  If bytVersionMajor = 0 Then
    ' Read and store current version of Access.
    bytVersionMajor = Val(SysCmd(acSysCmdAccessVer))
  End If
  
  If bytVersionMajor > cbytVersionMajorMax Then
    ' Use DateAdd() as is.
    datNext = DateAdd(strInterval, lngNumber, datDate)
  Else
    Select Case strInterval
      Case cstrIntervalHour, cstrIntervalMinute, cstrIntervalSecond
        Select Case strInterval
          Case cstrIntervalHour
            lngFactor = clngHoursPerDay
          Case cstrIntervalMinute
            lngFactor = clngMinutesPerDay
          Case cstrIntervalSecond
            lngFactor = clngSecondsPerDay
        End Select
        dblMilliseconds = clngMillisecondsPerDay * lngNumber / lngFactor
        datNext = MsecSerial(dblMilliseconds, datDate)
      Case Else
        datNext = DateAdd(strInterval, lngNumber, datDate)
    End Select
  End If
  
  DateAddFull = datNext
  
End Function

Public Function DateValid( _
  ByVal datDate As Date) _
  As Date
  
' Returns invalid numeric negative date values less than one day
' as their positive equivalents.
  
  Call ValidateDate(datDate)
  
  DateValid = datDate
  
End Function

Public Function DateLinear( _
  ByVal datDateNative As Date) _
  As Date

' Converts a native date value and returns the linear date value.
' Useful only for date values prior to 1899-12-30 as these have
' a negative numeric value.
  
  Dim datDate As Date
  
  datDate = datDateNative
  Call ConvDateToLinear(datDate)
  
  DateLinear = datDate

End Function

Public Function DateNative( _
  ByVal datDateLinear As Date) _
  As Date
  
' Converts a linear date value and returns the native date value.
' Useful only for date values prior to 1899-12-30 as these have
' a negative numeric value.
  
  Dim datDate As Date
  
  datDate = datDateLinear
  Call ConvDateToNative(datDate)
  
  DateNative = datDate
  
End Function

Private Sub ConvDateToLinear( _
  ByRef datDateNative As Date)

' Converts a native date value to a linear date value.
' Example:
'
'   Date     Time  Linear        Native
'   19000101 0000  2             2
'
'   18991231 1800  1,75          1,75
'   18991231 1200  1,5           1,5
'   18991231 0600  1,25          1,25
'   18991231 0000  1             1
'
'   18991230 1800  0,75          0,75
'   18991230 1200  0,5           0,5
'   18991230 0600  0,25          0,25
'   18991230 0000  0             0
'
'   18991229 1800 -0,25         -1,75
'   18991229 1200 -0,5          -1,5
'   18991229 0600 -0,75         -1,25
'   18991229 0000 -1            -1
'
'   18991228 1800 -1,25         -2,75
'   18991228 1200 -1,5          -2,5
'   18991228 0600 -1,75         -2,25
'   18991228 0000 -2            -2
  
  Dim dblDate As Double
  Dim dblTime As Double
  
  If datDateNative < 0 Then
    ' Get date (integer) part of datDateNative shifted one day
    ' if a time part is present as -Int() rounds up.
    dblDate = -Int(-datDateNative)
    ' Retrieve and reverse time (decimal) part.
    dblTime = dblDate - datDateNative
    ' Assemble and convert date and time part to linear date value.
    datDateNative = CVDate(dblDate + dblTime)
  Else
    ' Positive date values are linear by design.
  End If

End Sub

Private Sub ConvDateToNative( _
  ByRef datDateLinear As Date)
  
' Converts a linear date value to a native date value.
' Example:
'
'   Date     Time  Linear        Native
'   19000101 0000  2             2
'
'   18991231 1800  1,75          1,75
'   18991231 1200  1,5           1,5
'   18991231 0600  1,25          1,25
'   18991231 0000  1             1
'
'   18991230 1800  0,75          0,75
'   18991230 1200  0,5           0,5
'   18991230 0600  0,25          0,25
'   18991230 0000  0             0
'
'   18991229 1800 -0,25         -1,75
'   18991229 1200 -0,5          -1,5
'   18991229 0600 -0,75         -1,25
'   18991229 0000 -1            -1
'
'   18991228 1800 -1,25         -2,75
'   18991228 1200 -1,5          -2,5
'   18991228 0600 -1,75         -2,25
'   18991228 0000 -2            -2
  
  Dim dblDate As Double
  Dim dblTime As Double
  
  If datDateLinear < 0 Then
    ' Get date (integer) part of datDateLinear shifted one day
    ' if a time part is present as Int() rounds down.
    dblDate = Int(CDbl(datDateLinear))
    ' Retrieve and reverse time (decimal) part.
    dblTime = dblDate - datDateLinear
    ' Assemble and convert date and time part to native date value.
    datDateLinear = CVDate(dblDate + dblTime)
  Else
    ' Positive date values are linear by design.
  End If
  
End Sub

Private Sub ValidateDate( _
  ByRef datDate As Date)

' Converts numeric negative date values less than one day
' to their positive equivalents.
  
  If datDate < 0 Then
    If datDate > -1 Then
      ' Convert invalid date value to valid date value.
      datDate = -datDate
    End If
  End If

End Sub

Public Function DateMsecSet( _
  ByVal intMillisecond As Integer, _
  ByVal datDate As Date) _
  As Date
  
' Rounds off datDate to the second and optionally adds specified
' millisecond part up to and including 999 milliseconds.
'
' Typical usage:
'   Sequentialize a series of identical date values.
'
'   For each element in <collection of identical date values>
'     datDate = <read date value>
'     datDate = DateMsecSet(i, datDate)
'     <write date value> = datDate
'     i = i + 1
'   Next

  Const cintMillisecondMax  As Integer = 999
  
  ' Round off a millisecond part.
  Call RoundSecondOff(datDate)
  ' Add count of milliseconds up to 999.
  If intMillisecond > 0 Then
    If intMillisecond > cintMillisecondMax Then
      intMillisecond = cintMillisecondMax
    End If
    datDate = DateAddMsec(cstrIntervalMsec, intMillisecond, datDate)
  End If
  
  DateMsecSet = datDate
  
End Function

Public Function SplitDateMsec( _
  ByVal datDateMsec As Date, _
  Optional ByRef datDate As Date, _
  Optional ByRef datMsec As Date) _
  As Variant

' Splits datDateMsec into its components of
' date/time and millisecond as date values.
' These are returned ByRef as well as an array.

  Dim adatDate(0 To 1)  As Date
  
  datDate = datDateMsec
  ' Remove millisecond part.
  Call RoundSecondOff(datDate)
  ' Get milliseconds.
  datMsec = MsecSerial(DateDiffMsec("ms", datDate, datDateMsec))
  adatDate(0) = datDate
  adatDate(1) = datMsec
  
  SplitDateMsec = adatDate
  
End Function

Public Function JoinDateMsec( _
  ByVal avarDateMsec As Variant) _
  As Date

' Joins array of date/time and millisecond as date values
' to a compound date value.

  Dim datDate As Date
  
  datDate = DateSumDates(avarDateMsec(0), avarDateMsec(1))
  
  JoinDateMsec = datDate
  
End Function

