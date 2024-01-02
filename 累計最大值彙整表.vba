Sub update()
    Dim formula(101) '儲存儀器名稱
    Dim findr '尋找儀器表格位置row
    Dim findc '尋找儀器表格位置column
    '先更新最大值再抓取本次最大值
    '**變更最大值
    '左邊
    For i = 1 To 57
        If Abs(Worksheets("報告書").Cells(1 + i, "F")) > Abs(Worksheets("報告書").Cells(1 + i, "E")) Then
            Text = Cells(1 + i, "E")
            Worksheets("報告書").Cells(1 + i, "E") = Worksheets("報告書").Cells(1 + i, "F")
            Worksheets("報告書").Cells(1 + i, "E").Interior.Color = RGB(255, 0, 0)
            MsgBox (Worksheets("報告書").Cells(1 + i, "D") & "已替換" & vbCrLf & Text & "=>" & Cells(1 + i, "E"))
        End If
    Next i
    '右邊
    For i = 1 To 44
        If Abs(Worksheets("報告書").Cells(1 + i, "O")) > Abs(Worksheets("報告書").Cells(1 + i, "N")) Then
            Text = Cells(1 + i, "N")
            Worksheets("報告書").Cells(1 + i, "N") = Worksheets("報告書").Cells(1 + i, "O")
            Worksheets("報告書").Cells(1 + i, "N").Interior.Color = RGB(255, 0, 0)
            MsgBox (Worksheets("報告書").Cells(1 + i, "D") & "已替換" & vbCrLf & Text & "=>" & Cells(1 + i, "O"))
        End If
    Next i
    '**抓取本次最大值
    '抓取儀器名稱
    '左
    For i = 1 To 57
        formula(i) = Worksheets("報告書").Cells(1 + i, "D")
    Next i

    '右
    For i = 1 To 44
        formula(57 + i) = Worksheets("報告書").Cells(1 + i, "M")
    Next i
    '搜索更新
    '左
    For i = 1 To 57
        'WT(A)
        'MsgBox (formula(i))
        If Left(formula(i), 5) = "WT(A)" Then
            If Right(formula(i), 1) = "X" Then
                findr = Worksheets("各儀器").Range("B4:B91").Find(Replace(formula(i), "-X", ""), lookat:=xlPart).Row
                Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, "G")
                Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            End If
            If Right(formula(i), 1) = "Y" Then
                findr = Worksheets("各儀器").Range("B4:B91").Find(Replace(formula(i), "-Y", ""), lookat:=xlPart).Row
                Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, "I")
                Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            End If
        End If
        'WT
        If Left(formula(i), 2) = "WT" Then
            If Right(formula(i), 1) = "X" Then
                findr = Worksheets("各儀器").Range("B4:B91").Find(Replace(formula(i), "-X", ""), lookat:=xlPart).Row
                Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, "G")
                Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            End If
            If Right(formula(i), 1) = "Y" Then
                findr = Worksheets("各儀器").Range("B4:B91").Find(Replace(formula(i), "-Y", ""), lookat:=xlPart).Row
                Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, "I")
                Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
            End If
            Else
                '一般
                findr = Worksheets("各儀器").Range("B4:B91").Find(formula(i), lookat:=xlWhole).Row
                findc = Worksheets("各儀器").Range("B4:B91").Find(formula(i), lookat:=xlWhole).Column
                Worksheets("報告書").Cells(1 + i, "F") = Worksheets("各儀器").Cells(findr, findc + 5)
                Worksheets("報告書").Cells(1 + i, "F").Interior.Color = RGB(0, 255, 0)
        End If
    Next i
    '右
    For i = 1 To 44
        'WT(A)
        'MsgBox (formula(i))
        If Left(formula(57 + i), 5) = "WT(A)" Then
            If Right(formula(57 + i), 1) = "X" Then
                findr = Worksheets("各儀器").Range("B4:B91").Find(Replace(formula(57 + i), "-X", ""), lookat:=xlPart).Row
                Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, "G")
                Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            End If
            If Right(formula(57 + i), 1) = "Y" Then
                findr = Worksheets("各儀器").Range("B4:B91").Find(Replace(formula(57 + i), "-Y", ""), lookat:=xlPart).Row
                Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, "I")
                Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            End If
        End If
        'WT
        If Left(formula(57 + i), 2) = "WT" Then
            If Right(formula(57 + i), 1) = "X" Then
                findr = Worksheets("各儀器").Range("B4:B91").Find(Replace(formula(57 + i), "-X", ""), lookat:=xlPart).Row
                Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, "G")
                Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            End If
            If Right(formula(57 + i), 1) = "Y" Then
                findr = Worksheets("各儀器").Range("B4:B91").Find(Replace(formula(57 + i), "-Y", ""), lookat:=xlPart).Row
                Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, "I")
                Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
            End If
            Else
                '一般
                findr = Worksheets("各儀器").Range("B4:B91").Find(formula(57 + i), lookat:=xlWhole).Row
                findc = Worksheets("各儀器").Range("B4:B91").Find(formula(57 + i), lookat:=xlWhole).Column
                Worksheets("報告書").Cells(1 + i, "O") = Worksheets("各儀器").Cells(findr, findc + 5)
                Worksheets("報告書").Cells(1 + i, "O").Interior.Color = RGB(0, 255, 0)
        End If
    Next i
    MsgBox ("本月量測最大值更新完畢")
 
 
End Sub

