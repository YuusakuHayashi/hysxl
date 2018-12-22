Attribute VB_Name = "hysDate"
Option Explicit

'日付を引数に取る汎用的関数を提供する

Function hyYoubi(d As Date) As String
    Select Case Weekday(d)
        Case vbSunday
            hyYoubi = "日"
        Case vbMonday
            hyYoubi = "月"
        Case vbTuesday
            hyYoubi = "火"
        Case vbWednesday
            hyYoubi = "水"
        Case vbThursday
            hyYoubi = "木"
        Case vbFriday
            hyYoubi = "金"
        Case vbSaturday
            hyYoubi = "土"
    End Select
End Function

Function hySerial(dt As Date)
    'シリアルデイト関数
    hySerial = DateSerial(Year(dt), Month(dt), Day(dt))
End Function

Function hyFormatYYYYMMDD(dt As Date)
    'YYYYMMDDを文字列化
    hyFormatYYYYMMDD = Format(hySerial(dt), "yyyymmdd")
End Function

Function hyFormatYYYY(dt As Date)
    'YYYYを文字列化
    hyFormatYYYY = Format(hySerial(dt), "yyyy")
End Function

Function hyFormatYY(dt As Date)
    'YYを文字列化
    hyFormatYY = Format(hySerial(dt), "yy")
End Function

Function hyFormatMM(dt As Date)
    'MMを文字列化
    hyFormatMM = Format(hySerial(dt), "mm")
End Function

Function hyFormatDD(dt As Date)
    'DDを文字列化
    hyFormatDD = Format(hySerial(dt), "dd")
End Function

Function hyFormatPreLastDay(dt As Date)
    '前月末日を文字列化
    hyFormatPreLastDay = Format((DateSerial(Year(dt), Month(dt), 1) - 1), "dd")
End Function
    
Function hyFormatPreYY(dt As Date)
    '前月YYを文字列化
    hyFormatPreYY = Format(DateSerial(Year(dt), Month(dt) - 1, Day(dt)), "yy")
End Function

Function hyFormatPreMM(dt As Date)
    '前月MMを文字列化
    hyFormatPreMM = Format(DateSerial(Year(dt), Month(dt) - 1, Day(dt)), "mm")
End Function

Function hyFormatPreYYMM26(dt As Date) As String
    '前月26日を文字列化
    hyFormatPreYYMM26 = getFormatPreYY(dt) & getFormatPreMM(dt) & "26"
End Function

Function hyFormatPreYYMMLastDay(dt As Date) As String
    '前月末日YYMMDDを文字列化
    hyFormatPreYYMMLastDay = getFormatPreYY(dt) & getFormatPreMM(dt) & getPreLastDay(dt)
End Function

Function hyFormatTanaYYMM(dt As Date) As String
    '棚卸年月を文字列化
    hyFormatTanaYYMM = getFormatPreYY(dt) & getFormatPreMM(dt)
End Function

Function hyFormatRunDay(dt As Date) As String
    '処理日を文字列化
    hyFormatRunDay = hyFormatYYYY(dt) & hyFormatMM(dt) & hyFormatDD(dt)
End Function

Function hyFormatStartDay(dt As Date) As String
    '起動日を文字列化
    hyFormatStartDay = hyFormatYY(dt) & hyFormatMM(dt) & hyFormatDD(dt)
End Function

Function hyFormatCutOffDay(dt As Date, COD As String) As String
    '締切日を文字列化
    Select Case COD
        Case CInt(COD) >= 30
            hyFormatCutOffDay = getFormatPreYY(dt) & getFormatPreMM(dt) & COD
        Case Else
            hyFormatCutOffDay = hyFormatYY(dt) & getFormatPreMM(dt) & COD
    End Select
End Function
