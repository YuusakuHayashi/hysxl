Attribute VB_Name = "main"

Dim row As Integer              '-- �Z���s
Dim col As Integer              '-- �Z����
Dim root As String              '-- ���[�g�f�B���N�g���i�[
Const col_desc As Integer = 3   '-- ������(get_index())
Const col_kind As Integer = 2   '-- �t�@�C��/�t�H���_��ޗ�(get_index())

Sub test(d)
    row = 2
    col = 1
    Call get_index(d)
End Sub

Sub get_index(ByVal r As String)
	'r 					 -- ���[�U�w��̃��[�g�f�B���N�g��	
    Dim fso As Object   '-- file system obj
    Dim fo As Object    '-- folder obj

    root = r

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fo = fso.GetFolder(r)
    With fo
        Call get_file_index(.Files)	
        Call get_index_re(.SubFolders)
    End With
End Sub

Sub get_file_index(ByRef fc As Object, Optional Byval idx as Integer = 1)
	'fc					 --	file collection
	'idx 				 --	�K�w��\������counter
    Dim o As Object     '-- �R���N�V��������C�e���[�V�����̊i�[�pobj
    Dim f As Object     '-- obj���Q�Ɨpfile obj
    Dim bef As String   '-- ���ɓ��͂����������l���i�[

    For Each o In fc
        Set f = o
        With f
            'If Cells(row, col_desc) <> "" Then
            '    bef = Cells(row, col)
            '    If bef <> .Name Then
            '        Cells(row, col_desc).Font.ColorIndex = 15
            '    End If
            'End If
            Cells(row, idx) = .Name
            'Cells(row, col_kind) = "File"
            row = row + 1
        End With
    Next
    
End Sub

Sub get_index_re(ByRef foc As Object, Optional Byval idx as Integer = 1)
	'foc 				---	folder collection
	'idx 				---	�K�w��\������counter
    Dim o As Object     '-- �R���N�V��������C�e���[�V�����̊i�[�pobj
    Dim fo As Object    '-- obj���Q�Ɨpfolder obj
    Dim bef As String   '-- ���ɓ��͂����������l���i�[
    Dim p As String     '-- ���[�g�p�X������(�ȗ��p)

    For Each o In foc
        Set fo = o
        With fo
            p = Replace(.path, root, "")
            'If Cells(row, col_desc) <> "" Then
            '    bef = Cells(row, col)
            '    If bef <> p Then
            '        Cells(row, col_desc).Font.ColorIndex = 15
            '    End If
            'End If
            Cells(row, idx) = p
            'Cells(row, col_kind) = "Folder"
            row = row + 1
            Call get_file_index(.Files, idx + 1)
            Call get_index_re(.SubFolders, idx + 1)
        End With
    Next
End Sub

