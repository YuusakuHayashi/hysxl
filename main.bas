Attribute VB_Name = "main"
Option Explicit

Dim row As Integer              '-- �Z���s
Dim col As Integer              '-- �Z����
Dim root As String              '-- ���[�g�f�B���N�g��
Const col_desc As Integer = 3   '-- ������(get_index())
Const col_kind As Integer = 2   '-- �t�@�C��/�t�H���_��ޗ�(get_index())

Sub test()
    row = 2
    col = 1
    Call get_index("")
End Sub

Sub get_index(ByVal r As String)
    'r  --- ���[�g�f�B���N�g���i�[�p
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

    Dim o As Object     '-- �R���N�V��������C�e���[�V�����̊i�[�pobj
    Dim f As Object     '-- obj���Q�Ɨpfile obj
    Dim bef As String   '-- ���ɓ��͂����������l���i�[
    
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

    Dim o As Object     '-- �R���N�V��������C�e���[�V�����̊i�[�pobj
    Dim f As Object     '-- obj���Q�Ɨpfile obj
    Dim bef As String   '-- ���ɓ��͂����������l���i�[
    Dim p As String     '-- ���[�g�p�X������(�ȗ��p)
    
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

