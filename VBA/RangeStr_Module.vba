'-----------------------------------------------------------------------
'  RangeStr
' 行と列を数値で指定すると"A1:B2"みたいな文字列を返すVBA関数
' 
' Last modified 2017/01/11
'
'-----------------------------------------------------------------------
Option Explicit
'-----------------------------------------------------------------------
' row,col から、A1 形式の文字列を返す
' r1,c1:開始セル位置 r2,c2:終了セル位置
' f:""じゃなかったら絶対指定にする
' sh:シート or シート名称
Public Function RangeStr(row1 As Variant, col1 As Variant, _
    Optional row2 As Variant = 0, Optional col2 As Variant = 0, _
    Optional flg As Variant = "", Optional sh As Variant = "") As String
    
    Dim rmax As Long
    Dim cmax As Long
    Dim cn1 As String
    Dim cn2 As String
    
    rmax = 1048576 ' 2007の最大行 (2003 は 65536)
    cmax = 16384 ' 2007の最大列(2003は256)
    
    If flg <> "" Then flg = "$"
    row1 = VarIntCheck(row1)
    row2 = VarIntCheck(row2)
    col1 = VarIntCheck(col1)
    col2 = VarIntCheck(col2)
    
    If row1 > rmax Or row1 < 1 Or col1 > cmax Or col1 < 1 Then
        RangeStr = ""
        Exit Function
    End If
    
   
    cn1 = ColStr(col1)
    cn2 = ColStr(col2)
    RangeStr = SheetNameStr(sh) & flg & cn1 & flg & CStr(row1)
    
    If row2 > 0 And row2 < rmax And col2 > 0 And col2 < cmax Then
        RangeStr = RangeStr & ":" & flg & cn2 & flg & CStr(row2)
    End If

End Function

'-----------------------------------------------------------------------
' 複数列の文字列 "A:B"とか
Public Function ColsStr(col1 As Variant, Optional col2 As Variant = 0, _
        Optional flg As Variant = "", Optional sh As Variant = "") As String
    
    If flg <> "" Then flg = "$"
    col1 = VarIntCheck(col1)
    col2 = VarIntCheck(col2)
    
    If col1 > 0 Then
        ColsStr = SheetNameStr(sh) & flg & ColStr(col1)
        If col2 > 0 Then ColsStr = ColsStr & ":" & flg & ColStr(col2)
    End If

End Function

'-----------------------------------------------------------------------
' 複数行の文字列 "1:3"とか
Public Function RowsStr(row1 As Variant, Optional row2 As Variant = 0, _
        Optional flg As Variant = "", Optional sh As Variant = "") As String
    
    If flg <> "" Then flg = "$"
    row1 = VarIntCheck(row1)
    row2 = VarIntCheck(row2)
    
    If row1 > 0 Then
        RowsStr = SheetNameStr(sh) & flg & Format(row1)
        If row2 > 0 Then RowsStr = RowsStr & ":" & flg & Format(row2)
    End If

End Function

'-----------------------------------------------------------------------
' 選択した行を表す文字列を返す
' "シート名!$A$1:$B$5" みたいなの
Public Function SelectionRangeStr() As String
    Dim rng As Range
    Dim sh As Worksheet ' 選択してるシート
    Dim rs As Long  ' 選択エリアの最初の行
    Dim re As Long  ' 選択エリアの最後の行
    Dim cs As Long  ' 選択エリアの最初の列
    Dim ce As Long  ' 選択エリアの最後の列
    
    If TypeName(Selection) = "Range" Then
        Set rng = Selection
        rs = rng.Row
        re = rs + rng.Rows.Count - 1
        cs = rng.Column
        ce = cs + rng.Columns.Count - 1
        SelectionRangeStr = RangeStr(rs, cs, re, ce, "$", rng.Worksheet.Name)
    Else
        SelectionRangeStr = ""
    End If

End Function

'-----------------------------------------------------------------------
' 選択した行を表す文字列を返す
' "シート名!$A:$B" みたいなの
Public Function SelectionRowsStr() As String
    Dim rng As Range
    Dim sh As Worksheet ' 選択してるシート
    Dim rs As Long  ' 選択エリアの最初の行
    Dim re As Long  ' 選択エリアの最後の行
    
    If TypeName(Selection) = "Range" Then
        Set rng = Selection
        rs = rng.Row
        re = rs + rng.Rows.Count - 1
        SelectionRowsStr = RowsStr(rs, re, "$", rng.Worksheet.Name)
    Else
        SelectionRowsStr = ""
    End If

End Function

'-----------------------------------------------------------------------
' 選択範囲を表す文字列が示す範囲を選択する
' "シート名!$A:$B" みたいなの
Public Sub Select_Range(rngtxt As String)
    Dim sstr, rstr, pos, ch
    pos = InStr(rngtxt, "!")
    If pos > 0 Then
        sstr = Left(rngtxt, pos - 1)
        ' シート名の前後に''がついてた場合の対処
        ch = Left(sstr, 1)
        If ch = "'" Then sstr = Right(sstr, Len(sstr) - 1)
        ch = Right(sstr, 1)
        If ch = "'" Then sstr = Left(sstr, Len(sstr) - 1)
        rstr = Right(rngtxt, Len(rngtxt) - pos)
        ActiveWorkbook.Worksheets(sstr).Activate
    Else
        sstr = ""
        rstr = rngtxt
    End If
    'MsgBox "Text = " & rngtxt & " Sheet = " & sstr & " Range = " & rstr
    ActiveSheet.Range(rstr).Select

End Sub

'-----------------------------------------------------------------------
' 文字列から列番号を返す
Public Function StrCol(ctxt As String) As Long
    Dim tlen As Long
    Dim i As Long
    Dim clmno As Long
    Dim chrno As Long
    
    ctxt = StrConv(ctxt, vbUpperCase)
    tlen = Len(ctxt)
    clmno = 0
    If tlen > 0 Then
        For i = 0 To (tlen - 1)
            chrno = Asc(Mid(ctxt, tlen - i, 1)) - &H40
            clmno = clmno + ((26 ^ i) * chrno)
        Next i
    End If
    StrCol = clmno
End Function

'-----------------------------------------------------------------------
' 列数から文字列を返す(A～ZZZ)
Public Function ColStr(clmno As Variant) As String
    Dim ctxt As String
    ctxt = ""
    ColStr2 ctxt, VarIntCheck(clmno)
    ColStr = ctxt
End Function

'-----------------------------------------------------------------------
' ColStr から呼び出される関数
Private Function ColStr2(ByRef ctxt As String, ByVal clmno As Long) As Long
    Dim ret As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim MAX As Long
        
    MAX = 26
    
    If clmno < 1 Then
        ret = 0
    ElseIf clmno > MAX Then
        c1 = clmno Mod MAX
        If c1 = 0 Then
            c2 = Application.RoundDown(clmno / MAX, 0) - 1 'よもやデフォルトが四捨五入だとは…
            ctxt = Chr(&H40 + MAX) & ctxt
        Else
            c2 = Application.RoundDown(clmno / MAX, 0)
            ctxt = Chr(&H40 + c1) & ctxt
        End If
        ret = ColStr2(ctxt, c2)
    Else
        ctxt = Chr(&H40 + clmno) & ctxt
        ret = 0
    End If
    ColStr2 = ret
End Function

'-----------------------------------------------------------------------
' シート名文字列をRangeStr用に加工
Public Function SheetNameStr(sh As Variant) As String
    Dim shname As String

    If IsNull(sh) Then
        shname = ""
    ElseIf TypeName(sh) = "Worksheet" Then
        shname = sh.Name
    Else
        shname = CStr(sh)
    End If

    If shname <> "" Then
        If InStr(shname, " ") Then shname = "'" & shname & "'"
        SheetNameStr = shname & "!"
    Else
        SheetNameStr = ""
    End If

End Function

'-----------------------------------------------------------------------
' Variant変数の中を確認して整数に変換する
Public Function VarIntCheck(vi As Variant) As Long
    If IsNumeric(vi) Then
        VarIntCheck = CLng(vi)
    Else
        VarIntCheck = 0
    End If
End Function

'-----------------------------------------------------------------------
' てすと
Sub RengeStrTest()

    Select_Range "1:2"
    MsgBox SelectionRowsStr
    Select_Range "A1:c2"
    MsgBox SelectionRangeStr
    Select_Range "C3:B4"
    MsgBox SelectionRangeStr

'    MsgBox StrCol("AAZ")

'    MsgBox RangeStr(1, 2, 3, 4)
'    MsgBox RangeStr(0, 2, 3, 4)
'    MsgBox RangeStr(1, 0, 3, 4)
'    MsgBox RangeStr(1, 2, 0, 4)
'    MsgBox RangeStr(1, 2, 3, 0)
'    MsgBox RangeStr(1, 2, 0, 0, "", ActiveSheet)
'    MsgBox RangeStr(1, 2, 0, 0, "f", ActiveSheet)
'    MsgBox RangeStr(1, 2, 3, 4, "", ActiveSheet)
'    MsgBox RangeStr(1, 2, 3, 4, "f", ActiveSheet)

'    MsgBox RowsStr(1, 2)
'    MsgBox RowsStr(0, 2)
'    MsgBox RowsStr(1, 0)
'    MsgBox RowsStr(1, 2, "")
'    MsgBox RowsStr(1, 2, "f", ActiveSheet.Name)
'    MsgBox RowsStr(1, 2, "f", "すぺーすの あるシート名")

'    MsgBox ColsStr(1, 2)
'    MsgBox ColsStr(0, 2)
'    MsgBox ColsStr(1, 0)
'    MsgBox ColsStr(1, 2, "")
'    MsgBox ColsStr(1, 2, "f", ActiveSheet.Name)
'    MsgBox ColsStr(1, 2, "f", "すぺーすの あるシート名")

End Sub
