Attribute VB_Name = "main"
Option Explicit

Dim row As Integer              '-- セル行
Dim col As Integer              '-- セル列
Dim root As String              '-- ルートディレクトリ
Const col_desc As Integer = 3   '-- 説明列(get_index())
Const col_kind As Integer = 2   '-- ファイル/フォルダ種類列(get_index())

Sub test()
    row = 2
    col = 1
    Call get_index("")
End Sub

Sub get_index(ByVal r As String)
    'r  --- ルートディレクトリ格納用
    root = r
    
    Dim fso As Object   '-- file system obj
    Dim fo As Object    '-- folder obj
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fo = fso.GetFolder(r)
    With fo
        Call get_file_index(.Files)
        Call get_index_re(.SubFolders)
    End With
    
End Sub

Sub get_file_index(ByRef fc As Object)

    Dim o As Object     '-- コレクションからイテレーションの格納用obj
    Dim f As Object     '-- objを参照用file obj
    Dim bef As String   '-- 既に入力したあった値を格納
    
    For Each o In fc
        Set f = o
        With f
            If Cells(row, col_desc) <> "" Then
                bef = Cells(row, col)
                If bef <> .Name Then
                    Cells(row, col_desc).Font.ColorIndex = 15
                End If
            End If
            Cells(row, col) = .Name
            Cells(row, col_kind) = "File"
            row = row + 1
        End With
    Next
    
End Sub

Sub get_index_re(ByRef foc As Object)

    Dim o As Object     '-- コレクションからイテレーションの格納用obj
    Dim f As Object     '-- objを参照用file obj
    Dim bef As String   '-- 既に入力したあった値を格納
    Dim p As String     '-- ルートパス文字列(省略用)
    
    For Each o In foc
        Set fo = o
        With fo
            p = Replace(.path, root, "")
            If Cells(row, col_desc) <> "" Then
                bef = Cells(row, col)
                If bef <> p Then
                    Cells(row, col_desc).Font.ColorIndex = 15
                End If
            End If
            Cells(row, col) = p
            Cells(row, col_kind) = "Folder"
            row = row + 1
            Call get_file_index(.Files)
            Call get_index_re(.SubFolders)
        End With
    Next
End Sub

