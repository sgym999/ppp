Sub ShowCRUDColumnsIfAnySelected()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim row As Long
    row = ActiveCell.Row ' 現在の行

    ' CRUD に対応する列番号（必要に応じて変更）
    Dim crudCols As Variant
    crudCols = Array("C", "R", "U", "D") ' 文字 → 後で列番号に変換

    Dim colIndex As Long
    Dim showCRUD As Boolean: showCRUD = False

    ' CRUD 列番号を取得
    Dim crudIndexes() As Long
    ReDim crudIndexes(LBound(crudCols) To UBound(crudCols))
    Dim i As Long
    For i = LBound(crudCols) To UBound(crudCols)
        crudIndexes(i) = Columns(crudCols(i)).Column
    Next i

    ' 1つでも〇があるか判定
    For Each colIndex In crudIndexes
        If Trim(ws.Cells(row, colIndex).Value) = "〇" Then
            showCRUD = True
            Exit For
        End If
    Next colIndex

    ' すべての列を一旦表示
    ws.Columns.Hidden = False

    ' CRUD列以外を非表示（showCRUD=trueのときだけCRUDは残す）
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim isCRUD As Boolean: isCRUD = False
        For Each colIndex In crudIndexes
            If c = colIndex Then isCRUD = True: Exit For
        Next
        If Not isCRUD Then
            ws.Columns(c).Hidden = True
        End If
    Next
End Sub
