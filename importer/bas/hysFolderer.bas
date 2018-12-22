Attribute VB_Name = "hysFolderer"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function convertFilesToStrs(fs) As String()
    
    If fs.Count = 0 Then
        convertFilesToStrs = Split(VBA.vbNullString)
        Exit Function
    End If

    Dim l() As String
    Dim i As Integer
    Dim f As Object
    
    ReDim l(fs.Count - 1)

    For Each f In fs
        l(i) = f.Name
        i = i + 1
    Next
    
    convertFilesToStrs = l()
    
End Function

Sub Migrater(from_folder, to_folder, Optional ex_list = "")
    Dim fso As Object
    Dim fs As Object
    Dim i As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If IsArray(ex_list) Then
        For i = LBound(ex_list) To UBound(ex_list)
            If fso.FileExists(from_folder & "\" & ex_list(i)) Then
                fso.CopyFile from_folder & "\" & ex_list(i), to_folder & "\" & ex_list(i)
                Sleep 200
                fso.DeleteFile from_folder & "\" & ex_list(i), True
                Sleep 200
            End If
        Next
    Else
        If fso.FolderExists(from_folder) Then
            fso.CopyFolder from_folder, to_folder
            Sleep 200
            fso.DeleteFolder from_folder
            Sleep 200
            fso.CreateFolder from_folder
            Sleep 200
        End If
    End If
    
    Set fso = Nothing
    Set fs = Nothing
End Sub
