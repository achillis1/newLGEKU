Attribute VB_Name = "TimeZone"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function SystemTimeToFileTime Lib _
      "kernel32" (lpSystemTime As SYSTEMTIME, _
      lpFileTime As FILETIME) As Long
    
    Public Declare PtrSafe Function LocalFileTimeToFileTime Lib _
      "kernel32" (lpLocalFileTime As FILETIME, _
      lpFileTime As FILETIME) As Long
    
    Public Declare PtrSafe Function FileTimeToSystemTime Lib _
      "kernel32" (lpFileTime As FILETIME, lpSystemTime _
      As SYSTEMTIME) As Long

#Else
    Public Declare Function SystemTimeToFileTime Lib _
      "kernel32" (lpSystemTime As SYSTEMTIME, _
      lpFileTime As FILETIME) As Long
    
    Public Declare Function LocalFileTimeToFileTime Lib _
      "kernel32" (lpLocalFileTime As FILETIME, _
      lpFileTime As FILETIME) As Long
    
    Public Declare Function FileTimeToSystemTime Lib _
      "kernel32" (lpFileTime As FILETIME, lpSystemTime _
      As SYSTEMTIME) As Long
#End If

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


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

Public Function LocalTimeToUTC(dteTime As Date) As Date
    Dim dteLocalFileTime As FILETIME
    Dim dteFileTime As FILETIME
    Dim dteLocalSystemTime As SYSTEMTIME
    Dim dteSystemTime As SYSTEMTIME

    dteLocalSystemTime.wYear = CInt(Year(dteTime))
    dteLocalSystemTime.wMonth = CInt(Month(dteTime))
    dteLocalSystemTime.wDay = CInt(Day(dteTime))
    dteLocalSystemTime.wHour = CInt(Hour(dteTime))
    dteLocalSystemTime.wMinute = CInt(Minute(dteTime))
    dteLocalSystemTime.wSecond = CInt(Second(dteTime))

    Call SystemTimeToFileTime(dteLocalSystemTime, _
      dteLocalFileTime)
    Call LocalFileTimeToFileTime(dteLocalFileTime, _
      dteFileTime)
    Call FileTimeToSystemTime(dteFileTime, dteSystemTime)

    LocalTimeToUTC = CDate(dteSystemTime.wMonth & "/" & _
      dteSystemTime.wDay & "/" & _
      dteSystemTime.wYear & " " & _
      dteSystemTime.wHour & ":" & _
      dteSystemTime.wMinute & ":" & _
      dteSystemTime.wSecond)
End Function

Public Function LocalTimeToET(LocalTime As Date) As Date
    Dim StartDST As Date
    Dim EndDST As Date
    Dim UTCTime As Date
    UTCTime = LocalTimeToUTC(LocalTime)

    StartDST = DateAdd("h", 2, NthWeekday(2, 1, 3, Year(LocalTime)))
    EndDST = DateAdd("h", 2, NthWeekday(1, 1, 11, Year(LocalTime)))

    If LocalTime >= StartDST And LocalTime <= EndDST Then
        LocalTimeToET = DateAdd("h", -4, UTCTime)
    Else
        LocalTimeToET = DateAdd("h", -5, UTCTime)
    End If
End Function

Public Function NthWeekday(Position, DayIndex As Long, TargetMonth As Long, Optional TargetYear As Long)
    Dim FirstDate As Date

    If DayIndex < 1 Or DayIndex > 7 Then
        NthWeekday = CVErr(xlErrValue)
        Exit Function
    End If
    
    If TargetYear = 0 Then TargetYear = Year(Now)
    
    Select Case Position
        Case 1, 2, 3, 4, 5, "L", "l"
            FirstDate = DateSerial(TargetYear, TargetMonth, 1)
            If Weekday(FirstDate, vbSunday) < DayIndex Then
                FirstDate = FirstDate + (DayIndex - Weekday(FirstDate, vbSunday))
            ElseIf Weekday(FirstDate, vbSunday) > DayIndex Then
                FirstDate = FirstDate + (DayIndex + 7 - Weekday(FirstDate, vbSunday))
            End If
            
            If IsNumeric(Position) Then
                NthWeekday = FirstDate + (Position - 1) * 7
                If Month(NthWeekday) <> Month(FirstDate) Then NthWeekday = CVErr(xlErrValue)
            Else
                NthWeekday = FirstDate
                Do Until Month(NthWeekday) <> Month(NthWeekday + 7)
                    NthWeekday = NthWeekday + 7
                Loop
            End If
        
        Case Else
            NthWeekday = CVErr(xlErrValue)
    End Select
    
End Function


