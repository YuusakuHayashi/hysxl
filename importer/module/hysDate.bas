Attribute VB_Name = "hysDate"
Option Explicit

Function GetWeek(d As Date) As String
    Select Case Weekday(d)
        Case vbSunday
            GetWeek = "��"
        Case vbMonday
            GetWeek = "��"
        Case vbTuesday
            GetWeek = "��"
        Case vbWednesday
            GetWeek = "��"
        Case vbThursday
            GetWeek = "��"
        Case vbFriday
            GetWeek = "��"
        Case vbSaturday
            GetWeek = "�y"
    End Select
End Function

Function GetSerial(dt As Date) As Integer
    '�V���A���f�C�g�֐�
    GetSerial = DateSerial(Year(dt), Month(dt), Day(dt))
End Function

Function GetYYYYMMDD(dt As Date) As String
    'YYYYMMDD�𕶎���
    GetYYYYMMDD = Format(GetSerial(dt), "yyyymmdd")
End Function

Function GetYYYY(dt As Date) As String
    'YYYY�𕶎���
    GetYYYY = Format(GetSerial(dt), "yyyy")
End Function

Function GetYY(dt As Date) As String
    'YY�𕶎���
    GetYY = Format(GetSerial(dt), "yy")
End Function

Function GetMM(dt As Date) As String
    'MM�𕶎���
    GetMM = Format(GetSerial(dt), "mm")
End Function

Function GetDD(dt As Date) As String
    'DD�𕶎���
    GetDD = Format(GetSerial(dt), "dd")
End Function

Function GetLastDayOfPreviousMonth(dt As Date) As String
    '�O�������𕶎���
    GetLastDayOfPreviousMonth = Format((DateSerial(Year(dt), Month(dt), 1) - 1), "dd")
End Function
    
Function GetYYOfPreviousMonth(dt As Date) As String
    '�O��YY�𕶎���
    GetYYOfPreviousMonth = Format(DateSerial(Year(dt), Month(dt) - 1, Day(dt)), "yy")
End Function

Function GetPreviousMonth(dt As Date) As String
    '�O��MM�𕶎���
    GetPreviousMonth = Format(DateSerial(Year(dt), Month(dt) - 1, Day(dt)), "mm")
End Function

Function GetYYMMLastDayOfPreviousMonth(dt As Date) As String
    '�O������YYMMDD�𕶎���
    GetYYMMLastDayOfPreviousMonth = getFormatPreYY(dt) & getFormatPreMM(dt) & getPreLastDay(dt)
End Function
