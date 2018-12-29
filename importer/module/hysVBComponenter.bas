Attribute VB_Name = "hysVBComponenter"
Option Explicit

Function GetComponentsList(ComponentType As Integer)
    Dim i As Integer
    Dim j As Integer: j = 0
    Dim cnt As Integer: cnt = 0

    Dim bas_cnt As Integer: bas_cnt = 0
    Dim cls_cnt As Integer: cls_cnt = 0
    Dim frm_cnt As Integer: frm_cnt = 0
    Dim axd_cnt As Integer: axd_cnt = 0
    Dim obj_cnt As Integer: obj_cnt = 0
    Dim tmp1()
    Dim tmp2() As String
    
    ReDim tmp1(ThisWorkbook.VBProject.vbcomponents.Count, 4)
    
    With ThisWorkbook.VBProject
        For i = 1 To .vbcomponents.Count
            Select Case .vbcomponents(i).Type
                Case 1
                    tmp1(bas_cnt, 0) = .vbcomponents(i).Name
                    bas_cnt = bas_cnt + 1
                Case 2
                    tmp1(cls_cnt, 1) = .vbcomponents(i).Name
                    cls_cnt = cls_cnt + 1
                Case 3
                    tmp1(frm_cnt, 2) = .vbcomponents(i).Name
                    frm_cnt = frm_cnt + 1
                Case 11
                    tmp1(axd_cnt, 3) = .vbcomponents(i).Name
                    axd_cnt = axd_cnt + 1
                Case 100
                    tmp1(obj_cnt, 4) = .vbcomponents(i).Name
                    obj_cnt = obj_cnt + 1
                Case Else
            End Select
        Next
        Select Case ComponentType
            Case 1
                ReDim tmp2(bas_cnt)
                For i = 0 To bas_cnt
                    tmp2(i) = tmp1(i, 0)
                Next
            Case 2
                ReDim tmp2(cls_cnt)
                For i = 0 To cls_cnt
                    tmp2(i) = tmp1(i, 1)
                Next
            Case 3
                ReDim tmp2(frm_cnt)
                For i = 0 To frm_cnt
                    tmp2(i) = tmp1(i, 2)
                Next
            Case 11
                ReDim tmp2(axd_cnt)
                For i = 0 To axd_cnt
                    tmp2(i) = tmp1(i, 3)
                Next
            Case 100
                ReDim tmp2(obj_cnt)
                For i = 0 To obj_cnt
                    tmp2(i) = tmp1(i, 4)
                Next
'            Case 999
'                GetComponentsList = tmp1
            Case Else
        End Select
    End With
    
    If UBound(tmp2) > 0 Then
        ReDim Preserve tmp2(UBound(tmp2) - 1)
        GetComponentsList = tmp2
    Else
        GetComponentsList = Split(VBA.vbNullString)
    End If
End Function
