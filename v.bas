Sub ShowCRUDSetsWithoutHeaderCheck()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim targetRow As Long
    targetRow = ActiveCell.Row

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' すべての列を非表示
    ws.Columns.Hidden = True

    ' A～C列は常に表示
    ws.Columns("A:C").Hidden = False

    Dim startCol As Long
    startCol = 4 ' D列以降からCRUDセットとみなす

    Dim c As Long
    For c = startCol To lastCol Step 4
        Dim crudCols(1 To 4) As Long
        crudCols(1) = c
        crudCols(2) = c + 1
        crudCols(3) = c + 2
        crudCols(4) = c + 3

        Dim shouldShow As Boolean: shouldShow = False
        Dim i As Long
        For i = 1 To 4
            If crudCols(i) <= lastCol Then
                If Trim(ws.Cells(targetRow, crudCols(i)).Value) = "〇" Then
                    shouldShow = True
                    Exit For
                End If
            End If
        Next i

        If shouldShow Then
            For i = 1 To 4
                If crudCols(i) <= lastCol Then
                    ws.Columns(crudCols(i)).Hidden = False
                End If
            Next i
        End If
    Next c
End Sub
