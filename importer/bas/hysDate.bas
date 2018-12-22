Attribute VB_Name = "hysDate"
Option Explicit

Function GetWeek(d As Date) As String
    Select Case Weekday(d)
        Case vbSunday
            GetWeek = "日"
        Case vbMonday
            GetWeek = "月"
        Case vbTuesday
            GetWeek = "火"
        Case vbWednesday
            GetWeek = "水"
        Case vbThursday
            GetWeek = "木"
        Case vbFriday
            GetWeek = "金"
        Case vbSaturday
            GetWeek = "土"
    End Select
End Function

Function GetSerial(dt As Date) As Integer
    'シリアルデイト関数
    GetSerial = DateSerial(Year(dt), Month(dt), Day(dt))
End Function

Function GetYYYYMMDD(dt As Date) As String
    'YYYYMMDDを文字列化
    GetYYYYMMDD = Format(GetSerial(dt), "yyyymmdd")
End Function

Function GetYYYY(dt As Date) As String
    'YYYYを文字列化
    GetYYYY = Format(GetSerial(dt), "yyyy")
End Function

Function GetYY(dt As Date) As String
    'YYを文字列化
    GetYY = Format(GetSerial(dt), "yy")
End Function

Function GetMM(dt As Date) As String
    'MMを文字列化
    GetMM = Format(GetSerial(dt), "mm")
End Function

Function GetDD(dt As Date) As String
    'DDを文字列化
    GetDD = Format(GetSerial(dt), "dd")
End Function

Function GetLastDayOfPreviousMonth(dt As Date) As String
    '前月末日を文字列化
    GetLastDayOfPreviousMonth = Format((DateSerial(Year(dt), Month(dt), 1) - 1), "dd")
End Function
    
Function GetYYOfPreviousMonth(dt As Date) As String
    '前月YYを文字列化
    GetYYOfPreviousMonth = Format(DateSerial(Year(dt), Month(dt) - 1, Day(dt)), "yy")
End Function

Function GetPreviousMonth(dt As Date) As String
    '前月MMを文字列化
    GetPreviousMonth = Format(DateSerial(Year(dt), Month(dt) - 1, Day(dt)), "mm")
End Function

Function GetYYMMLastDayOfPreviousMonth(dt As Date) As String
    '前月末日YYMMDDを文字列化
    GetYYMMLastDayOfPreviousMonth = getFormatPreYY(dt) & getFormatPreMM(dt) & getPreLastDay(dt)
End Function
