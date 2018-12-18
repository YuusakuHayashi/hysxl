Attribute VB_Name = "mod_executeShell"
Option Explicit

Sub execApp(PATH As String)
    'Dim obj
    'Set obj = CreateObject('WScript.Shell')

    Dim obj As New IWshRuntimeLibrary.WshShell
    Set obj = New IWshRuntimeLibrary.WshShell
    Dim r As Long

    r = obj.Run(PATH, 1, False)


End Sub

Sub getCurDir()
    
End Sub


Sub mydebug()
    Application.EnableEvents = True
    Range("A4").Columns.AutoFit
End Sub
