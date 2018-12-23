Attribute VB_Name = "hysDebugger"
Option Explicit

Sub Logger(msg As String, Optional cnt = "")
    Debug.Print "Logger>>> " & Now & " " & msg & cnt
End Sub

Sub MemoryLogger(arr)
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i) & " => " & VarPtr(arr(i))
    Next
End Sub

Sub PrintListItem(arr)
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Debug.Print i & " " & arr(i)
    Next
End Sub
