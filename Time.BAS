Attribute VB_Name = "mdlTime"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

'mdlTime
'As internet time values are long values
'we'll have to convert them to user-friendly
'values. In addition, internet time uses GMT.
'We'll have to convert the time to the local
'time zone.

'allow only declared variables to be used
Option Explicit

'a time type used to store hours, minutes and seconds
Public Type TimeF
    H As Integer
    M As Integer
    S As Integer
End Type

'the systemTime type used to store an "absolute position in time"...
Private Type SYSTEMTIME
        'date... year, month, day of week, day...
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer 'which day in week, each index indicated a day
        wDay As Integer
        'time... hour, minute, second and millisecond
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

'this type is used to store the local time zone information
Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName As String * 64
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName As String * 64
        DaylightDate As SYSTEMTIME
        DaylightBias As Long 'The bias is the difference, in minutes, between UTC time and local time.
End Type

'API used to get the time zone of the user to evaluate the time correctly
Private Declare Function GetTimeZoneInformation& Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION)

'This function converts a GMT time value to the local time
Public Function GMTtoLT(myTime As TimeF) As TimeF
    Dim myTZ As TIME_ZONE_INFORMATION 'the timezone of the user
    
    'get the user's time zone and store it in myTZ.
    'if the timezone is in "Daylight" range...
    If GetTimeZoneInformation(myTZ) = 2 Then
        '...also calculate the Daylight difference
        myTime.H = myTime.H - (myTZ.DaylightBias \ 60)
    End If
    
    'calculate the difference between the local time and UTC.
    myTime.H = myTime.H - (myTZ.Bias \ 60)
    
    'if we got under zero
    If myTime.H < 0 Then
        'add 24 to the hours
        myTime.H = 24 + myTime.H
    End If
    
    'return the local time
    GMTtoLT = myTime
End Function
Public Function GetTimeZoneBIAS() As Byte
    Dim myTZ As TIME_ZONE_INFORMATION 'the timezone of the user
    Dim intReturn As Integer
    
    'get the user's time zone and store it in myTZ.
    'if the timezone is in "Daylight" range...
    If GetTimeZoneInformation(myTZ) = 2 Then
        '...also calculate the Daylight difference
        intReturn = (myTZ.DaylightBias \ 60)
    End If
    
    'calculate the difference between the local time and UTC.
    intReturn = intReturn + 11 - (myTZ.Bias \ 60)
    
    'if we got under zero
    If intReturn < 0 Then
        'add 24 to the hours
        intReturn = 24 + intReturn
    End If

    'return the local timezone byte
    GetTimeZoneBIAS = intReturn
End Function
Public Function IsDaylight() As Boolean
    'function used to determine if the local timezone is "Daylight"
    Dim myTZ As TIME_ZONE_INFORMATION 'the local timezone of the user
    Dim myDate As Date 'the date of the user, which we get using VB's Date function
    'value calculated with the formula Months * 31 + Days...
    Dim TD As Integer '...using VB's date
    Dim DL As Integer '...using Daylight date
    Dim ST As Integer '...using Standart(non-daylight) date
    'get VB's date
    myDate = DateTime.Date
    'calculate the value...
    TD = month(myDate) * 31 + Day(myDate)
    'get the timezone of the user
    GetTimeZoneInformation myTZ
    'calculate the same value for Daylight date and Standart date.
    DL = myTZ.DaylightDate.wMonth * 31 + myTZ.DaylightDate.wDay
    ST = myTZ.StandardDate.wMonth * 31 + myTZ.StandardDate.wDay
    
    'if VB's value is greater or equal than the Daylight value and it's smaller than the Standart value
    'it's daylight; else it's not.
    IsDaylight = (TD >= DL) And (TD < ST)
End Function
Public Function LocalTime(ByVal lValue As Long) As String
    ' Now for the LocalTime function. Take t
    '     he long value representing the number
    ' of seconds since January 1, 1970 and c
    '     reate a useable time structure from it.
    ' Return a formatted string MM/DD/YYYY H
    '     H:MM:SS
    Const lSecPerDay As Long = 86400 'seconds in a day = 60*60*24
    Const lSecPerHour As Integer = 3600 'seconds in an hour = 60 * 60
    Const lSecPerYearNormal As Long = 31536000 'seconds in a year, if it's normal
    Const lSecPerYearLeap As Long = 31622400 'seconds in a year, if it's leap
    'the current year, month, day, hour, minute and second
    Dim year As Long
    Dim month As Long
    Dim Day As Long
    Dim Hour As Long
    Dim Minute As Long
    Dim second As Long
    Dim Temp As Long 'a temporary long variable
    Dim timeResult As TimeF 'the result time
    
    year = 70 ' starting point, 1970
    'Calculate the year
    'we start looping to calculate it
    Do While (lValue > 0)
        'if there is a year left in lValue...
        If (lValue - IIf(isLeapYear(year), lSecPerYearLeap, lSecPerYearNormal)) > 0 Then
            'remove these seconds
            lValue = lValue - IIf(isLeapYear(year), lSecPerYearLeap, lSecPerYearNormal)
            'and add a year
            year = year + 1
        'if there no more years
        Else
            'we are done with years
            Exit Do
        End If
    Loop
    'as we start from 70 which is 1970 we'll have to add 1900
    year = year + 1900
    'Calculate the month
    'we start from month 1
    month = 1

    'loop again
    Do While (lValue > 0)
        'get the seconds in this month(it's different for each month and also
        '                              different if the year is leap and the month is February)
        Temp = SecsInMonth(year, month)

        'if there's a month left
        If (lValue - Temp) > 0 Then
            'remove it from lValue
            lValue = lValue - Temp
            'and add it to months
            month = month + 1
        'there are no more months in lValue
        Else
            'we are done with months
            Exit Do
        End If
    Loop
    ' calculate day
    ' get how many hours are in lValue
    Day = 1 + lValue \ lSecPerDay
    ' and remove them from lValue
    lValue = lValue - (Day - 1) * lSecPerDay
    
    ' calculate Hour using the same method
    Hour = lValue \ lSecPerHour
    
    lValue = lValue - Hour * lSecPerHour

    ' calculate the mintutes...
    Minute = lValue \ 60
    lValue = lValue - Minute * 60
    
    'what's left are seconds
    second = lValue
    
    'convert what we've done to TimeF type...
    timeResult.H = Hour
    timeResult.M = Minute
    timeResult.S = second
    'convert the time from GMT to local time
    timeResult = GMTtoLT(timeResult)
    
    'return timeResult formatted correctly.
    LocalTime = month & "/" & Day & "/" & year & " at " & timeResult.H & ":" & FixLeadingZero(timeResult.M) & ":" & FixLeadingZero(timeResult.S)
End Function
Public Function isLeapYear(ByVal NumYear As Long) As Boolean
    'if the year can be divided by 4 it's leap
    isLeapYear = NumYear / 4 = NumYear \ 4
End Function
Public Function SecsInMonth(ByVal YearNum As Long, ByVal MonthNum As Long) As Long
    'calculate how many seconds are in month MonthNum
    If MonthNum = 1 Or MonthNum = 3 Or MonthNum = 5 Or MonthNum = 7 Or MonthNum = 8 Or MonthNum = 10 Or MonthNum = 12 Then
        'it's a month with 31 days
        SecsInMonth = 2678400
    ElseIf MonthNum = 4 Or MonthNum = 6 Or MonthNum = 9 Or MonthNum = 11 Then
        'it's a month with 30 days
        SecsInMonth = 2592000
    ElseIf MonthNum = 2 Then
        'it's February, depends on the year...
        'if it's leap it has 29 days, else it has 28 days...
        SecsInMonth = IIf(isLeapYear(YearNum), 2505600, 2419200)
    End If
End Function

Public Function ToTimeStamp() As Long
'this returns the seconds from 1.1.1970 to now
ToTimeStamp = DateDiff("s", CDate(#1/1/1970#), Now)
End Function
