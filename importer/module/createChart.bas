Attribute VB_Name = "createChart"
Option Explicit

Sub Main(tgt, dn, dt, dc, dv)
'{{{
initializing:
    Application.EnableEvents = False
    
        Dim creationrange As Range
        Set creationrange = Range(Cells(tgt.Row, tgt.Column), Cells(tgt.Row, tgt.Column + dc - 1))
        
        creationrange.Value = ""
        
        With creationrange.Offset(1)
            .UnMerge
            .Value = ""
            With .Borders
                .Weight = xlThin
            End With
            .EntireRow.AutoFit
        End With
        
FILLER_TREATMENT:
        Dim nolabel As Boolean
        Dim dn_digit() As String
        Dim i As Integer
        ReDim dn_digit(Len(dn))
        
        If dn = "filler" _
        Or dn = "FILLER" _
        Or dv = "filler" _
        Or dv = "FILLER" Then
            creationrange.Font.Color = RGB(255, 0, 0)
            nolabel = True
            GoTo CREATE_DIGIT
        Else
            creationrange.Font.Color = RGB(0, 0, 0)
        End If
SOKEOK_TREATMENT:
        If dn = "sok" _
        Or dn = "SOK" _
        Or dn = "eok" _
        Or dn = "EOK" Then
            creationrange.Font.Color = RGB(255, 0, 0)
            nolabel = True
            GoTo CREATE_DIGIT
        Else
            creationrange.Font.Color = RGB(0, 0, 0)
        End If
KANJI_TREATEMENT:
        For i = 1 To Len(dn)
            If (Asc(Mid(dn, i, 1)) >= 65 And Asc(Mid(dn, i, 1)) <= 90) Or _
            (Asc(Mid(dn, i, 1)) >= 97 And Asc(Mid(dn, i, 1)) <= 122) Then
                GoTo CLOSING
            Else
                dn_digit(i) = Mid(dn, i, 1)
            End If
        Next
CREATE_DIGIT:
        If dt = "9" Then
            creationrange.Value = "9"
        ElseIf dt = "x" Or dt = "X" Then
            creationrange.Value = "X"
        ElseIf dt = "" Then
            Set creationrange = Range(tgt, tgt.Offset(, 1))
            For i = 1 To UBound(dn_digit)
                With creationrange
                    .Merge
                    .Value = dn_digit(i)
                    .HorizontalAlignment = xlCenter
                End With
                Set creationrange = Range(creationrange.Offset(, 1), creationrange.Offset(, 2))
            Next
            nolabel = True
        End If
        
        If nolabel = True Then
            GoTo CLOSING
        End If
CREATE_LABEL:

        With creationrange.Offset(1)
            .Merge
            .Value = dn
            .HorizontalAlignment = xlCenter
            With .Borders
                '.LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
CLOSING:
    Application.EnableEvents = True
'}}}
End Sub

Function checkString(tgt) As Variant()
    Dim tmp, tmp21, tmp22 As Variant
    Dim dataname, datatype, datavalue As String
    Dim datacount As Integer
    Dim ary(4) As Variant
        
    tmp = Split(tgt.Value, " ")
    On Error Resume Next
    tmp21 = Split(tmp(2), "(")
    tmp22 = Split(tmp21(1), ")")
    
    dataname = tmp(0)
    datavalue = tmp(4)
    datatype = tmp21(0)
    datacount = CInt(tmp22(0))
    
    If datacount < 1 Then
        datacount = 1
    End If
    
    ary(0) = dataname
    ary(1) = datatype
    ary(2) = datacount
    ary(3) = datavalue
    
    checkString = ary
    
End Function
 
Function checkRange(tgt) As Boolean
    If Application.Intersect(tgt, Range("C4:FJ65")) Is Nothing Then
        checkRange = True
    End If
End Function

Function checkNull(tgt) As Boolean
    If tgt.Value = "" Then
        checkNull = True
    End If
End Function

Function checkSpec() As Boolean
    If Cells(1, 1).Value = "x" Or Cells(1, 1).Value = "X" Then
        checkSpec = True
    End If
End Function

Sub enableevents_debug()
    Application.EnableEvents = False
    Application.EnableEvents = True
End Sub



