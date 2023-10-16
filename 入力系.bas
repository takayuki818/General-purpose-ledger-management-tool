Attribute VB_Name = "入力系"
Option Explicit
Sub 新規登録モード()
    With Sheets("入力フォーム")
        .Unprotect
        .Range("_転記先行") = "新規"
        Call 入力フォームクリア
        .Protect
    End With
End Sub
Sub 入力フォームクリア()
    Dim 終行 As Long, 行 As Long
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 項目リスト(2 To 終行, 1 To 1)
        For 行 = 2 To 終行
            項目リスト(行, 1) = .Cells(行, 1)
        Next
    End With
    With Sheets("入力フォーム")
        For 行 = 2 To 終行
            .Range(項目リスト(行, 1)).MergeArea.ClearContents
        Next
    End With
End Sub
Sub 登録更新()
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long, 記録行 As Long
    Dim 文 As String
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 設定(1 To 終行 - 1, 1 To 2)
        For 行 = 2 To 終行
            設定(行 - 1, 1) = .Cells(行, 1)
            設定(行 - 1, 2) = .Cells(行, 2)
            If 最右列 < 設定(行 - 1, 2) Then 最右列 = 設定(行 - 1, 2)
        Next
    End With
    With Sheets("入力フォーム")
        ReDim 配列(1 To 1, 1 To 最右列)
        For 行 = 1 To UBound(設定, 1)
            配列(1, 設定(行, 2)) = .Range(設定(行, 1))
        Next
        If .Range("_転記先行") <> "新規" Then 記録行 = .Range("_転記先行")
    End With
    With Sheets("管理台帳")
        If 記録行 = 0 Then
            For 列 = 1 To 最右列
                If 記録行 < .Cells(Rows.Count, 列).End(xlUp).Row + 1 Then
                    記録行 = .Cells(Rows.Count, 列).End(xlUp).Row + 1
                End If
            Next
            文 = "新規登録してよろしいですか？"
            Else: 文 = "台帳を更新してよろしいですか？"
        End If
        If MsgBox(文, vbYesNo) = vbYes Then
            .Cells(記録行, 1).Resize(1, 最右列) = 配列
            Call 新規登録モード
            Application.EnableEvents = False
            .Activate
            Application.EnableEvents = True
            .Cells(記録行, 1).Activate
        End If
    End With
End Sub
Sub 台帳戻し(記録行 As Long)
    Dim 終行 As Long, 行 As Long, 最右列 As Long, 列 As Long
    Application.ScreenUpdating = False
    With Sheets("台帳転記設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim 設定(1 To 終行 - 1, 1 To 2)
        For 行 = 2 To 終行
            設定(行 - 1, 1) = .Cells(行, 1)
            設定(行 - 1, 2) = .Cells(行, 2)
            If 最右列 < 設定(行 - 1, 2) Then 最右列 = 設定(行 - 1, 2)
        Next
    End With
    With Sheets("管理台帳")
        ReDim 配列(1 To 1, 1 To 最右列)
        For 列 = 1 To 最右列
            配列(1, 列) = .Cells(記録行, 列)
        Next
    End With
    With Sheets("入力フォーム")
        Call 入力フォームクリア
        .Unprotect
        For 行 = 1 To UBound(設定, 1)
            .Range(設定(行, 1)) = 配列(1, 設定(行, 2))
        Next
        .Protect
        .Activate
    End With
    Application.ScreenUpdating = True
End Sub
