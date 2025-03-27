' 罫線方向フラグ（ビットマスク用）
Public Const BORDER_TOP As Integer = 1      ' 上
Public Const BORDER_BOTTOM As Integer = 2   ' 下
Public Const BORDER_LEFT As Integer = 4     ' 左
Public Const BORDER_RIGHT As Integer = 8    ' 右

Sub Main()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 使用されているセルを取得
    Dim usedRng As Range
    Set usedRng = ws.UsedRange

    ' セル結合をすべて解除 + 結合セルの「左上のセル」以外の値をクリア
    UnmergeAllCellsAndClearHiddenCellValues usedRng

    ' 表位置を特定するため、"項番" を含むセル位置を取得
    Dim noCell As Range
    Set noCell = FindNoCell(usedRng)
    
    If noCell Is Nothing Then
        MsgBox "「項番」が見つかりませんでした。", vbExclamation
        Exit Sub
    End If
    
    ' 表データの読み取りを行う。
    ' ・罫線で囲まれた複数セルのまとまりは「論理セル」とする。
    
    ' ・同一行における「各列の高さ」は、「"項番"論理セルの列」を基準とする。
    
    ' ・同一列における「各行の横幅」は、「ヘッダー行の論理セル」を基準とする。
    
    ' ・論理セル内の値は、その中の各セル（「内部セル」とする）を1行目の左端から右方向に進み、
    '   右端まで到達したら、2行目に処理を移す形で読み取りを行う。
    '   なお、例外として、複数行に跨る論理セルについては、一番上の行の時に値を読み取り、
    '   途中の行の場合（"項番"論理セルの上端と同じ位置に罫線がない場合）は、読み取りをスキップし、一番上の行と同じ値を設定する。
    
End Sub

' セル結合をすべて解除 + 結合セルの「左上のセル」以外の値をクリア
Private Sub UnmergeAllCellsAndClearHiddenCellValues(ByRef usedRng As Range)
    Dim cell As Range
    Dim mergedArea As Range
    Dim hiddenCell As Range

    For Each cell In usedRng
        ' 左上セルだけを処理対象にする
        If cell.MergeCells And cell.Address = cell.MergeArea.Cells(1, 1).Address Then
            Set mergedArea = cell.MergeArea

            ' セル結合を解除
            mergedArea.UnMerge

            ' 左上セル以外の値をクリア
            For Each hiddenCell In mergedArea.Cells
                If hiddenCell.Address <> cell.Address Then
                    hiddenCell.ClearContents
                End If
            Next
        End If
    Next
End Sub

' 表位置を特定するため、"項番" を含むセル位置を取得
Private Function FindNoCell(ByRef usedRng As Range) As Range
    Dim noCell As Range
    Set noCell = usedRng.Find(What:="項番", _
                                   LookIn:=xlValues, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlNext, _
                                   MatchCase:=False)
    
    Set FindNoCell = noCell
End Function

Private Function FindTopLeftInnerCell(ByRef startCell As Range) As Range
    Dim ws As Worksheet
    Set ws = startCell.Worksheet

    Dim row As Long
    Dim col As Long
    row = startCell.row
    col = startCell.Column

    Dim topRow As Long
    Dim leftCol As Long
    topRow = row
    leftCol = col

    ' 上方向に移動（罫線がある限り）
    Do While topRow > 1
        If GetBorderFlags(ws.Cells(topRow - 1, col)) And BORDER_TOP <> 0 Then
            Exit Do
        End If
        
        topRow = topRow - 1
    Loop

    ' 左方向に移動（罫線がある限り）
    Do While leftCol > 1
        If GetBorderFlags(ws.Cells(row, leftCol - 1)) And BORDER_LEFT <> 0 Then
            Exit Do
        End If
        
        leftCol = leftCol - 1
    Loop

    Set FindTopLeftBorderedCell = ws.Cells(topRow, leftCol)
End Function


Private Function GetBorderFlags(c As Range) As Integer
    On Error Resume Next
    
    Dim flags As Integer
    flags = 0
    
    If c.Borders(xlEdgeTop).LineStyle <> xlNone Then
        flags = flags Or BORDER_TOP
    End If
    If c.Borders(xlEdgeBottom).LineStyle <> xlNone Then
        flags = flags Or BORDER_BOTTOM
    End If
    If c.Borders(xlEdgeLeft).LineStyle <> xlNone Then
        flags = flags Or BORDER_LEFT
    End If
    If c.Borders(xlEdgeRight).LineStyle <> xlNone Then
        flags = flags Or BORDER_RIGHT
    End If
    
    GetBorderFlags = flags
End Function

