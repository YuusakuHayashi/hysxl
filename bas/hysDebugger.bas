Attribute VB_Name = "hysDebugger"
Option Explicit

Sub Logger(msg As String, Optional cnt = "")
    Debug.Print "Logger>>> " & Now & " " & msg & cnt
End Sub
