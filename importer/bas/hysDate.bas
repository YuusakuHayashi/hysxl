Attribute VB_Name = "hysDate"
Option Explicit

'���t�������Ɏ��ėp�I�֐���񋟂���

Function hyYoubi(d As Date) As String
    Select Case Weekday(d)
        Case vbSunday
            hyYoubi = "��"
        Case vbMonday
            hyYoubi = "��"
        Case vbTuesday
            hyYoubi = "��"
        Case vbWednesday
            hyYoubi = "��"
        Case vbThursday
            hyYoubi = "��"
        Case vbFriday
            hyYoubi = "��"
        Case vbSaturday
            hyYoubi = "�y"
    End Select
End Function

Function hySerial(dt As Date)
    '�V���A���f�C�g�֐�
    hySerial = DateSerial(Year(dt), Month(dt), Day(dt))
End Function

Function hyFormatYYYYMMDD(dt As Date)
    'YYYYMMDD�𕶎���
    hyFormatYYYYMMDD = Format(hySerial(dt), "yyyymmdd")
End Function

Function hyFormatYYYY(dt As Date)
    'YYYY�𕶎���
    hyFormatYYYY = Format(hySerial(dt), "yyyy")
End Function

Function hyFormatYY(dt As Date)
    'YY�𕶎���
    hyFormatYY = Format(hySerial(dt), "yy")
End Function

Function hyFormatMM(dt As Date)
    'MM�𕶎���
    hyFormatMM = Format(hySerial(dt), "mm")
End Function

Function hyFormatDD(dt As Date)
    'DD�𕶎���
    hyFormatDD = Format(hySerial(dt), "dd")
End Function

Function hyFormatPreLastDay(dt As Date)
    '�O�������𕶎���
    hyFormatPreLastDay = Format((DateSerial(Year(dt), Month(dt), 1) - 1), "dd")
End Function
    
Function hyFormatPreYY(dt As Date)
    '�O��YY�𕶎���
    hyFormatPreYY = Format(DateSerial(Year(dt), Month(dt) - 1, Day(dt)), "yy")
End Function

Function hyFormatPreMM(dt As Date)
    '�O��MM�𕶎���
    hyFormatPreMM = Format(DateSerial(Year(dt), Month(dt) - 1, Day(dt)), "mm")
End Function

Function hyFormatPreYYMM26(dt As Date) As String
    '�O��26���𕶎���
    hyFormatPreYYMM26 = getFormatPreYY(dt) & getFormatPreMM(dt) & "26"
End Function

Function hyFormatPreYYMMLastDay(dt As Date) As String
    '�O������YYMMDD�𕶎���
    hyFormatPreYYMMLastDay = getFormatPreYY(dt) & getFormatPreMM(dt) & getPreLastDay(dt)
End Function

Function hyFormatTanaYYMM(dt As Date) As String
    '�I���N���𕶎���
    hyFormatTanaYYMM = getFormatPreYY(dt) & getFormatPreMM(dt)
End Function

Function hyFormatRunDay(dt As Date) As String
    '�������𕶎���
    hyFormatRunDay = hyFormatYYYY(dt) & hyFormatMM(dt) & hyFormatDD(dt)
End Function

Function hyFormatStartDay(dt As Date) As String
    '�N�����𕶎���
    hyFormatStartDay = hyFormatYY(dt) & hyFormatMM(dt) & hyFormatDD(dt)
End Function

Function hyFormatCutOffDay(dt As Date, COD As String) As String
    '���ؓ��𕶎���
    Select Case COD
        Case CInt(COD) >= 30
            hyFormatCutOffDay = getFormatPreYY(dt) & getFormatPreMM(dt) & COD
        Case Else
            hyFormatCutOffDay = hyFormatYY(dt) & getFormatPreMM(dt) & COD
    End Select
End Function
