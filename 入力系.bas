Attribute VB_Name = "���͌n"
Option Explicit
Sub �V�K�o�^���[�h()
    With Sheets("���̓t�H�[��")
        .Unprotect
        .Range("_�]�L��s") = "�V�K"
        Call ���̓t�H�[���N���A
        .Protect
    End With
End Sub
Sub ���̓t�H�[���N���A()
    Dim �I�s As Long, �s As Long
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim ���ڃ��X�g(2 To �I�s, 1 To 1)
        For �s = 2 To �I�s
            ���ڃ��X�g(�s, 1) = .Cells(�s, 1)
        Next
    End With
    With Sheets("���̓t�H�[��")
        For �s = 2 To �I�s
            .Range(���ڃ��X�g(�s, 1)).MergeArea.ClearContents
        Next
    End With
End Sub
Sub �o�^�X�V()
    Dim �I�s As Long, �s As Long, �ŉE�� As Long, �� As Long, �L�^�s As Long
    Dim �� As String
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim �ݒ�(1 To �I�s - 1, 1 To 2)
        For �s = 2 To �I�s
            �ݒ�(�s - 1, 1) = .Cells(�s, 1)
            �ݒ�(�s - 1, 2) = .Cells(�s, 2)
            If �ŉE�� < �ݒ�(�s - 1, 2) Then �ŉE�� = �ݒ�(�s - 1, 2)
        Next
    End With
    With Sheets("���̓t�H�[��")
        ReDim �z��(1 To 1, 1 To �ŉE��)
        For �s = 1 To UBound(�ݒ�, 1)
            �z��(1, �ݒ�(�s, 2)) = .Range(�ݒ�(�s, 1))
        Next
        If .Range("_�]�L��s") <> "�V�K" Then �L�^�s = .Range("_�]�L��s")
    End With
    With Sheets("�Ǘ��䒠")
        If �L�^�s = 0 Then
            For �� = 1 To �ŉE��
                If �L�^�s < .Cells(Rows.Count, ��).End(xlUp).Row + 1 Then
                    �L�^�s = .Cells(Rows.Count, ��).End(xlUp).Row + 1
                End If
            Next
            �� = "�V�K�o�^���Ă�낵���ł����H"
            Else: �� = "�䒠���X�V���Ă�낵���ł����H"
        End If
        If MsgBox(��, vbYesNo) = vbYes Then
            .Cells(�L�^�s, 1).Resize(1, �ŉE��) = �z��
            Call �V�K�o�^���[�h
            Application.EnableEvents = False
            .Activate
            Application.EnableEvents = True
            .Cells(�L�^�s, 1).Activate
        End If
    End With
End Sub
Sub �䒠�߂�(�L�^�s As Long)
    Dim �I�s As Long, �s As Long, �ŉE�� As Long, �� As Long
    Application.ScreenUpdating = False
    With Sheets("�䒠�]�L�ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim �ݒ�(1 To �I�s - 1, 1 To 2)
        For �s = 2 To �I�s
            �ݒ�(�s - 1, 1) = .Cells(�s, 1)
            �ݒ�(�s - 1, 2) = .Cells(�s, 2)
            If �ŉE�� < �ݒ�(�s - 1, 2) Then �ŉE�� = �ݒ�(�s - 1, 2)
        Next
    End With
    With Sheets("�Ǘ��䒠")
        ReDim �z��(1 To 1, 1 To �ŉE��)
        For �� = 1 To �ŉE��
            �z��(1, ��) = .Cells(�L�^�s, ��)
        Next
    End With
    With Sheets("���̓t�H�[��")
        Call ���̓t�H�[���N���A
        .Unprotect
        For �s = 1 To UBound(�ݒ�, 1)
            .Range(�ݒ�(�s, 1)) = �z��(1, �ݒ�(�s, 2))
        Next
        .Protect
        .Activate
    End With
    Application.ScreenUpdating = True
End Sub
