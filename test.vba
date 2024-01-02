   Sub test()

   End Sub
   '1 to 57
    For i = 1 To 57
        'WT(A)
        If Left(formula(i), 5) = "WT(A)" Then
            'MsgBox (formula(i))
            findr = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(i), 7), lookat:=xlPart).Row
            'MsgBox (find)
            If Right(formula(i), 1) = "X" Then
                findc = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(i), 7), lookat:=xlPart).Column
                'MsgBox (Worksheets("各儀器").Cells(findr, findc + 5))
                Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, findc + 5)
                Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            End If
            If Right(formula(i), 1) = "Y" Then
                findc = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(i), 7), lookat:=xlPart).Column
                'MsgBox (Worksheets("各儀器").Cells(findr, findc + 7))
                 Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, findc + 7)
                 Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            End If
        End If
        'WT
        If Left(formula(i), 2) = "WT" Then
            'MsgBox (formula(i))
            findr = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(i), 4), lookat:=xlPart).Row
            'MsgBox (find)
            If Right(formula(i), 1) = "X" Then
                findc = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(i), 4), lookat:=xlPart).Column
                'MsgBox (Worksheets("各儀器").Cells(findr, findc + 5))
                Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, findc + 5)
                Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            End If
            If Right(formula(i), 1) = "Y" Then
                findc = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(i), 4), lookat:=xlPart).Column
                'MsgBox (Worksheets("各儀器").Cells(findr, findc + 7) & "," & findr & "," & findc + 7)
                 Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, findc + 7)
                 Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            End If
        Else
            '一般
            'MsgBox (formula(i))
            findr = Worksheets("各儀器").Range("B4:B91").Find(formula(i), lookat:=xlWhole).Row
            findc = Worksheets("各儀器").Range("B4:B91").Find(formula(i), lookat:=xlWhole).Column
            Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, findc + 5)
            Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            'MsgBox (Worksheets("各儀器").Cells(findr, findc + 5)) '(findr & "," & findc + 5)
            'MsgBox (findr & "," & findc + 5)
        End If
    Next i
    '57 to 101
    For i = 1 To 44
        'WT(A)
        If Left(formula(57 + i), 5) = "WT(A)" Then
            'MsgBox (formula(i))
            findr = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(57 + i), 7), lookat:=xlPart).Row
            'MsgBox (find)
            If Right(formula(57 + i), 1) = "X" Then
                findc = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(57 + i), 7), lookat:=xlPart).Column
                'MsgBox (Worksheets("各儀器").Cells(findr, findc + 5))
                Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, findc + 5)
                Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            End If
            If Right(formula(57 + i), 1) = "Y" Then
                findc = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(57 + i), 7), lookat:=xlPart).Column
                'MsgBox (Worksheets("各儀器").Cells(findr, findc + 7))
                 Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, findc + 7)
                 Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            End If
        End If
        'WT
        If Left(formula(57 + i), 2) = "WT" Then
            'MsgBox (formula(i))
            findr = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(57 + i), 4), lookat:=xlPart).Row
            'MsgBox (find)
            If Right(formula(57 + i), 1) = "X" Then
                findc = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(57 + i), 4), lookat:=xlPart).Column
                'MsgBox (Worksheets("各儀器").Cells(findr, findc + 5))
                Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, findc + 5)
                Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            End If
            If Right(formula(57 + i), 1) = "Y" Then
                findc = Worksheets("各儀器").Range("B4:B91").Find(Left(formula(57 + i), 4), lookat:=xlPart).Column
                'MsgBox (Worksheets("各儀器").Cells(findr, findc + 7))
                 Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, findc + 7)
                 Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            End If
        Else
            '一般
            'MsgBox (formula(i))
            findr = Worksheets("各儀器").Range("B4:B91").Find(formula(57 + i), lookat:=xlWhole).Row
            findc = Worksheets("各儀器").Range("B4:B91").Find(formula(57 + i), lookat:=xlWhole).Column
            Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, findc + 5)
            Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            'MsgBox (Worksheets("各儀器").Cells(findr, findc + 5)) '(findr & "," & findc + 5)
            'MsgBox (findr & "," & findc + 5)
        End If
    Next i
End Sub


'選擇檔案位置
Private Sub CommandButton1_Click()
With Application.FileDialog(msoFileDialogOpen)
   .InitialFileName = "W:\2017"
   .AllowMultiSelect = True
   .Show
   For i = 1 To .SelectedItems.Count
     Cells(i, 1) = .SelectedItems(i)
     MsgBox .SelectedItems(i)
   Next
End With
End Sub


