Option Explicit

'-----------------------------------------------------------------------
' 見た目重視で整形されたシートをなんとか再利用したい時に使うマクロ
' Last Midified 2016/12/07
'-----------------------------------------------------------------------


'-----------------------------------------------------------------------
' 指定範囲の空行を削除する
Sub 空行削除()
    Dim sh As Worksheet
    Dim sel As Range
    
    If TypeName(Selection) = "Range" Then
        Set sh = ActiveSheet
        Set sel = Selection
        EmptyRowDelete sh, sel.Row, sel.Row + sel.Rows.Count, 2
    End If
End Sub


'-----------------------------------------------------------------------
' 指定の列にデータが含まれない行を削除するマクロ
' チェックする列は３つまで
Sub EmptyRowDelete(sh As Worksheet, srow As Long, erow As Long, _
            col1 As Long, Optional col2 As Long = 0, Optional col3 As Long = 0)
    
    Dim r As Long, cf1 As String, cf2 As String, cf3 As String
    
    cf1 = ""
    cf2 = ""
    cf3 = ""
    For r = erow To srow Step -1
        If col1 > 0 Then
            cf1 = sh.Cells(r, col1).Value
        End If
        If col2 > 0 Then
            cf2 = sh.Cells(r, col2).Value
        End If
        If col3 > 0 Then
            cf3 = sh.Cells(r, col3).Value
        End If
        If cf1 = "" And cf2 = "" And cf3 = "" Then
            sh.Rows(r).Delete Shift:=xlUp
        End If
    Next r
End Sub


'-----------------------------------------------------------------------
Sub 行列入れ替えテスト()
    RowToColumns ActiveSheet, 1, 3
End Sub


'-----------------------------------------------------------------------
' 指定の列の値を指定の行数ごとに隣の列にコピーして空になった行を削除するマクロ
' 要Rangestr
Sub RowToColumns(sh As Worksheet, col As Long, cols As Long, _
                Optional srow As Long = 1, Optional erow As Long = 0)
    Dim dr As Long, er As Long
    Dim i
        
    For i = 1 To cols ' コピー先の列を作る
        sh.Columns(col + 1).Insert
    Next i
    
    dr = srow
    er = 0
    
    Do While er < 100
        
        If sh.Cells(dr, col).Value = "" Then
            er = er + 1 ' データがない行はスキップ
        Else
            er = 0
            Application.ScreenUpdating = False
            sh.Range(RangeStr(dr + 1, col, dr + (cols - 1), col)).Copy
            sh.Cells(dr, col + 1).PasteSpecial Paste:=xlPasteValues, Transpose:=True
            Application.CutCopyMode = False
            sh.Rows(RowsStr(dr + 1, dr + (cols - 1))).Delete Shift:=xlUp
            Application.ScreenUpdating = True
        End If
        
        dr = dr + 1
        If erow > 0 And dr > erow Then Exit Do
    Loop
                    
End Sub
