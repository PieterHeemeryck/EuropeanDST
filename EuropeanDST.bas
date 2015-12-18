Attribute VB_Name = "EuropeanDST"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EuropeanDST
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' Modified by Pieter Heemeryck to be used for Central European time which states that Daylight
' Saving Time (DST) starts the last sunday of March (2 AM -> 3 AM) and ends the last sunday of October
' (3 AM -> 2 AM).
'
' This module contains functions for working with local times in Daylight Saving Time
' and Standard Time (STD). The rules and computations here are valid for Central European time only !!!
' Users of other locales are welcome to change the code to meet the local DST/Standard
' rules.
'
' For more information about Summer Time in Europe, see
' https://en.wikipedia.org/wiki/Summer_Time_in_Europe
'
' Transition from DST To Standard time moves the current time of day backwards (earlier)
' one hour at 03:00:00 AM on the transition day. Therefore, the interval 02:00:00 to
' 02:59:59 occurs twice on the transition day, once during DST and then again when the
' time moves earlier at the start of STD. It is impossible to definitively determine
' whether the time, for example, 02:30:00 has is DST or STD.
'
' Transition from Standard Time to DST moves the current time of day forward (later), at
' 02:00:00 AM changing to 03:00:00 AM on the transition day. Therefore, the interval
' 02:00:00 to 02:59:59 does not exist on the transition day. For example, the time
' 02:30:00 on the transition day does not exist and is an invalid time.
'
' This module contains the following functions:
'           FirstDayOfMonth
'               This returns the date of the first day of the specified month and year.
'
'           FirstDayOfWeekInMonth
'               This returns the date of the first DayOfWeek in the specified month and year.
'
'           IsDateWithinDST
'               This returns True or False indicating whether the specified date and time is
'               within Daylight Savings Time. If the date is the same as the transition date
'               between DST and STD, and time is within 01:00:00 and 01:59:59, DST is assumed.
'               If the date is the same as the transition date between STD and DST, and the
'               time with within 02:00:00 and 02:59:59, DST is assumed.
'
'           IsDSTNow
'               This returns True or False indicating whether it is currently DST.
'
'           LastDayOfMonth
'               This returns the date of the last day of the specified month and year.
'
'           LastDayOfWeekInMonth
'               This returns the date of the last DayOfWeek in the specified month and year.
'
'           NthDayOfWeekInMonth
'               This returns the date of the Nth DayOfWeek in the specified month and year.
'
'           TransitionDateDstToStandard
'               This returns the date and time of the transition from DST to STD.
'
'           TransitonDateStandardToDST
'               This returns the date and time of the transion from STD to DST.
'
' The functions in this module call upon one another, so it is strongly recommended that you
' import the entire module into your project, rather than copy/pasting individual procedures.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Application Constants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''
' These constants define the day of week,
' month, and the Nth Day Of Week in which
''''''''''''''''''''''''''''''''''''''''
Private Const C_RULE_SWITCH_YEAR = 2007 ' Not used, only needed for the USA version of DST
Private Const C_TRANSITION_DAY = vbSunday
Private Const C_TRANSITION_HOUR = 2
Private Const OCTOBER = 10
Private Const NOVEMBER = 11
Private Const MARCH = 3
Private Const APRIL = 4



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows Constants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
Private Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF
Private Const TIME_ZONE_ID_STANDARD As Long = 1
Private Const TIME_ZONE_ID_UNKNOWN As Long = 0

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows Types
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

' NOTE: The Windows API Text Viewer utility incorrect sizes the arrays as (32). They should
' be explicitly dimension as (0 to 31).
Public Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows APIs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function GetTimeZoneInformation Lib "kernel32" ( _
    lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Procedures
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function DaylightTimeZoneName() As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DaylightTimeZoneName
' This returns a string with the name of the daylight time zone
' of the current locale, as defined by Windows. Note that this
' is value does NOT reflect whether DST is presently in effect.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim TZI As TIME_ZONE_INFORMATION
    Dim Res As Long
    Dim S As String
    Res = GetTimeZoneInformation(TZI)
    S = IntArrayToString(Arr:=TZI.DaylightName)
    DaylightTimeZoneName = S

End Function


Public Function FirstDayOfMonth(MM As Integer, YYYY As Integer) As Date
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FirstDayOfMonth
' This returns the first day of the specified month MM and year YYYY.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    FirstDayOfMonth = DateSerial(YYYY, MM, 1)
End Function



Public Function FirstDayOfWeekInMonth(MM As Integer, YYYY As Integer, _
    DayOfWeek As VbDayOfWeek) As Date
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FirstDayOfWeekInMonth
' This returns the First DayOfWeek in the specified
' Month MM of Year YYYY. Returns -1 if an error occurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FirstOfMonth As Date
Dim DD As Long
Dim FirstOfMonthDay As VbDayOfWeek
''''''''''''''''''''''''''''''''
' Ensure MM is between 1 and 12
''''''''''''''''''''''''''''''''
Select Case MM
    Case 1 To 12
        ' ok
    Case Else
        FirstDayOfWeekInMonth = -1
        Exit Function
End Select
'''''''''''''''''''''''''''''''''
' Ensure year is 4 digits between
' 1900 and 9999.
'''''''''''''''''''''''''''''''''
Select Case YYYY
    Case 1900 To 9999
        'ok
    Case Else
        FirstDayOfWeekInMonth = -1
        Exit Function
End Select
        
'''''''''''''''''''''''''''''''''''''''''''''''''
' Get the first day of the month.
'''''''''''''''''''''''''''''''''''''''''''''''''
FirstOfMonth = FirstDayOfMonth(MM, YYYY)
'''''''''''''''''''''''''''''''''''''''''''''''''
' Get the weekday (Sunday = 1, Saturday = 7)
' of the first day of the month.
'''''''''''''''''''''''''''''''''''''''''''''''''
FirstOfMonthDay = Weekday(FirstOfMonth, vbSunday)
'''''''''''''''''''''''''''''''''''''''''''''''''
' compute the Day number (1 to 7) of the first
' DayOfWeek of the month.
'''''''''''''''''''''''''''''''''''''''''''''''''
DD = ((DayOfWeek - FirstOfMonthDay + 7) Mod 7) + 1

'''''''''''''''''''''''''''''''''''''''''''''''''
' Return the result as a date.
'''''''''''''''''''''''''''''''''''''''''''''''''
FirstDayOfWeekInMonth = DateSerial(YYYY, MM, DD)

End Function


Public Function IsDateWithinDST(TheDate As Date) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsDateWithinDST
' This function returns True or False indicating whether TheDate is within Daylight Savings
' Time. If TheDate is the transition date from STD to DST, and the time is between 02:00:00 AM
' and 02:59:59 AM (this time never actually exists, as the time moves from 02:00:00 to 03:00:00)
' it is assumed to be in DST. If TheDate is the transition date from DST to STD and the time
' is between 02:00:00 and 02:59:59 (this time occurs twice, once in DST and again when the
' time is moved earlier back to 02:00:00), it is assumed to be in DST Time.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim DstToStd As Date
Dim StdToDst As Date
Dim StartDSTTime As Date
Dim StartSTDTIme As Date
Dim InTimeOnly As Double
Dim CompTimeOnly As Double

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the transition dates DST to/from STD.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
DstToStd = TransitionDateDstToStandard(Year(TheDate))
StdToDst = TransitonDateStandardToDST(Year(TheDate))

''''''''''''''''''''''''''''''''''''''''''''''''''''
' If TheDate is greater than Int(StdToDst) and less
' Int(StdToDst), then the date is not on a transition
' date and is within DST.
''''''''''''''''''''''''''''''''''''''''''''''''''''
If (TheDate > StdToDst) And (TheDate < DstToStd) Then
    IsDateWithinDST = True
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''
' If TheDate is less than Int(StdToDst) Or greater
' the Int(DstToStd) then the date is not on a transition
' date and with NOT within DST.
'''''''''''''''''''''''''''''''''''''''''''''''''''
If (TheDate < StdToDst) And (TheDate < DstToStd) Then
    IsDateWithinDST = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''
' If int(TheDate) is equal to Int(StdToDst) then
' we have to compare the time of the input date
' to the time of the StdToDst time.
''''''''''''''''''''''''''''''''''''''''''''''''''
If Int(TheDate) = Int(StdToDst) Then
    '''''''''''''''''''''''''''''''''
    ' TheDate is the transition from
    ' StdToDst. Test the time.
    ' InTimeOnly contains ONLY the
    ' fraction (time) portion of
    ' TheDate, with no date component.
    ' 0 <= InTimeOnly < 1
    '''''''''''''''''''''''''''''''''
    InTimeOnly = TrueFraction(CDbl(TheDate))
    If (InTimeOnly >= TimeSerial(C_TRANSITION_HOUR, 0, 0)) And _
        (InTimeOnly <= TimeSerial(C_TRANSITION_HOUR, 59, 59)) Then
        ''''''''''''''''''''''''''''''''''''''''
        ' We're in the interval between 02:00:00
        ' and 02:59:59 on the transition between
        ' Standard to Daylight time. This time
        ' doesn't really exist. Assume DST.
        ''''''''''''''''''''''''''''''''''''''''
        IsDateWithinDST = True
        Exit Function
    
    Else
        ''''''''''''''''''''''''''''''''''''''''
        ' We're not in between 02:00:00 and
        ' 02:59:59. Test the time and return
        ' a result.
        ''''''''''''''''''''''''''''''''''''''''
        If InTimeOnly < TimeSerial(C_TRANSITION_HOUR, 0, 0) Then
            IsDateWithinDST = False
        Else
            IsDateWithinDST = True
        End If
    End If
Else
    If Int(TheDate) = Int(DstToStd) Then
        ''''''''''''''''''''''''''''''''''''''''''
        ' TheDate is the transition date between
        ' DST To STD. Test the time.
        ' InTimeOnly contains ONLY the fractional
        ' (time) component of TheDate, with no
        ' date component.
        ' 0 <= InTimeOnly < 1
        '''''''''''''''''''''''''''''''''''''''''
        InTimeOnly = TrueFraction(CDbl(TheDate))
        If (InTimeOnly >= TimeSerial(C_TRANSITION_HOUR, 0, 0)) And _
            (InTimeOnly <= TimeSerial(C_TRANSITION_HOUR, 59, 59)) Then
            '''''''''''''''''''''''''''''''''''''''''''''''''
            ' We're within the duplicated hours (02:00:00 to
            ' 02:59:59), assume DST
            '''''''''''''''''''''''''''''''''''''''''''''''''
            IsDateWithinDST = True
            Exit Function
        Else
            '''''''''''''''''''''''''''''''''''''''''''''''''
            ' We're not within the duplicate hour. Test the
            ' time. If the time is less than the transition
            ' time, we're in DST. Otherwise, we're in STD.
            '''''''''''''''''''''''''''''''''''''''''''''''''
            If InTimeOnly < TimeSerial(C_TRANSITION_HOUR, 0, 0) Then
                IsDateWithinDST = True
                Exit Function
            Else
                ''''''''''''''''''''''''
                ' we're not in DST
                ''''''''''''''''''''''''
                IsDateWithinDST = False
                Exit Function
            End If
        End If
    End If
End If

End Function


Public Function IsDSTNow() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsDSTNow
' This return TRUE or FALSE indicating whether the current date and time
' is within DST. If this cannot be determined, it defaults to False.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim TZI As TIME_ZONE_INFORMATION
    Dim Res As Long
    Res = GetTimeZoneInformation(TZI)
    If Res = TIME_ZONE_ID_DAYLIGHT Then
        IsDSTNow = True
    Else
        IsDSTNow = False
    End If
End Function


Public Function LastDayOfMonth(MM As Integer, YYYY As Integer) As Date
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LastDayOfMonth
' This returns the last day of the specified month MM and year YYYY.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    LastDayOfMonth = DateSerial(YYYY, MM + 1, 0)
End Function


Public Function LastDayOfWeekInMonth(MM As Integer, YYYY As Integer, DayOfWeek As VbDayOfWeek) As Date
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LastDayOfWeekInMonth
' This funciton returns the LAST DayOfWeek in the specified month MM and
' year YYYY.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim DT As Date
Dim SaveDT As Date
Dim TestDT As Date
Dim N As Long
DT = FirstDayOfWeekInMonth(MM, YYYY, DayOfWeek)
SaveDT = DT
TestDT = DT
For N = 3 To 6
    TestDT = DT + (N * 7)
    If Month(TestDT) = MM Then
        SaveDT = TestDT
    Else
        Exit For
    End If
Next N
        
LastDayOfWeekInMonth = SaveDT

End Function

Public Function NthDayOfWeekInMonth(MM As Integer, YYYY As Integer, _
    Nth As Integer, DayOfWeek As VbDayOfWeek) As Date
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NthDayOfWeekInMonth
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FirstDayOfWeekOfMonth As Date
Dim NthDayOfWeek As Date

'''''''''''''''''''''''''''''''''''''''''''''''''
' Get the first DayOfWeek in the give month MM
' in year YYYY.
'''''''''''''''''''''''''''''''''''''''''''''''''
FirstDayOfWeekOfMonth = FirstDayOfWeekInMonth(MM, YYYY, DayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''
' Add the number of weeks - 1 to the
' FirstDayOfWeekOfMonth.
'''''''''''''''''''''''''''''''''''''''''''''''''
NthDayOfWeek = FirstDayOfWeekOfMonth + ((Nth - 1) * 7)
'''''''''''''''''''''''''''''''''''''''''''''''''
' Return the result
'''''''''''''''''''''''''''''''''''''''''''''''''
NthDayOfWeekInMonth = NthDayOfWeek

End Function


Public Function StandardTimeZoneName() As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' StandardTimeZoneName
' This returns a string with the name of the standard time zone
' of the current locale, as defined by Windows. Note that this
' is value does NOT reflect whether DST is presently in effect.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim TZI As TIME_ZONE_INFORMATION
    Dim Res As Long
    Dim S As String
    Res = GetTimeZoneInformation(TZI)
    S = IntArrayToString(Arr:=TZI.StandardName)
    StandardTimeZoneName = S

End Function


Public Function TransitionDateDstToStandard(YYYY As Integer) As Date
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TransitionDateDstToStandard
' This returns the date and time of the transition from DST to STD.
' Note that since the time is moved backwards (earlier), the hour
' 02:00:00 AM to 02:59:59 AM occurs twice on the transition day.
' There is no definitive way to determine whether a time like 02:30:00AM
' on the transition day is DST or STD.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

TransitionDateDstToStandard = LastDayOfWeekInMonth(OCTOBER, YYYY, C_TRANSITION_DAY) + _
    TimeSerial(C_TRANSITION_HOUR + 1, 0, 0)

End Function


Public Function TransitonDateStandardToDST(YYYY As Integer) As Date
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TransitonDateStandardToDST
' This returns the date and time of the transition from STD to DST.
' Note that since the time is moved forward (later), the time interval
' 02:00:00AM to 02:59:59AM on the transition day does not exist. It is
' an invalid time.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

TransitonDateStandardToDST = LastDayOfWeekInMonth(MARCH, YYYY, C_TRANSITION_DAY) + _
    TimeSerial(C_TRANSITION_HOUR, 0, 0)


End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function TrimToNull(Text As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''
' TrimToNull
' This function returns the portion of Text that
' is to the left of the first vbNullChar character.
' If vbNullChar is not found, the function reutrns
' the complete Text string.
'''''''''''''''''''''''''''''''''''''''''''''''''
Dim Pos As Integer
Pos = InStr(1, Text, vbNullChar, vbBinaryCompare)
If Pos Then
    TrimToNull = Left(Text, Pos - 1)
Else
    TrimToNull = Text
End If

End Function

Private Function TrueFraction(D As Double) As Double
''''''''''''''''''''''''''''''''''''''''''''''''''''
' TrueFraction
' This function returns the fractional portion of
' portion of the number D. NO ROUNDING IS DONE. Only
' the fraction portion is returned as a double,
' regardless of its value or sign. The sign of the
' result is the same as the sign of the input value.
''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim S As String
    Dim Pos As Integer
    S = CStr(D)
    Pos = InStr(1, S, Application.International(xlDecimalSeparator))
    If Pos Then
        S = "0" & Mid(S, Pos)
    End If
    TrueFraction = CDbl(S) * IIf(D <= 0, -1, 1)
End Function


Private Function TrueInteger(D As Double) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''
' TrueInteger
' This funciton returns the integer portion of the
' Double value D. NO ROUNDING IS DONE. The integer
' portion of the number, regardless of the value
' of the fraction portion is returned as a long.
' This differs from CInt function which will round
' values up to the next integer (positive values)
' or down to the next integer (negative values), and
' differs from the Int function which always rounds
' down. Examples follow:
'       Double      Int     CInt    TrueInteger
'       ------      ----    ----    -----------
'       8.9         8       9           8
'       8.1         8       8           8
'      -8.9        -9      -9          -8
'      -8.1        -9      -8          -8
''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim S As String
    Dim Pos As Integer
    S = CStr(D)
    Pos = InStr(1, S, Application.International(xlDecimalSeparator))
    If Pos Then
        S = Left(S, Pos - 1)
    End If
    TrueInteger = CLng(S)
End Function


Private Function IntArrayToString(Arr() As Integer) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IntArrayToString
' Converts an array of integers into an regular string.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim S As String
Dim N As Long
For N = LBound(Arr) To UBound(Arr)
    S = S & Chr(Arr(N))
Next N
IntArrayToString = S


End Function


