Attribute VB_Name = "basGodPaste"
Option Explicit

'--------------------------------------------------------------
' マージセルの代表セルの値コピー
'--------------------------------------------------------------
Sub copyMergeCellVal()

    Dim i As Long
    Dim j As Long
    Dim r As Range

    Dim strLine As String
    Dim strBuf As String

    If rlxCheckSelectRange = False Then
        Exit Sub
    End If

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo e
'    Application.ScreenUpdating = False

    For i = Selection(1).Row To Selection(Selection.Count).Row

        For j = Selection(1).Column To Selection(Selection.Count).Column
        
            Set r = Cells(i, j)
        
            'マージセルなら左上のみ処理
            If (r.MergeCells = False Or r.MergeCells = True And r.MergeArea(1, 1).Address = r.Address) Then
        
                If j = Selection(1).Column Then
                    strLine = addQuat(r.Value)
                Else
                    strLine = strLine & vbTab & addQuat(r.Value)
                End If
            End If
        
        Next
        Set r = Cells(i, Selection(1).Column)
        If (r.MergeCells = False Or r.MergeCells = True And r.MergeArea(1, 1).Address = r.Address) Then
            If i = Selection(1).Row Then
                strBuf = strLine
            Else
                strBuf = strBuf & vbCrLf & strLine
            End If
        End If
    Next

    If Len(strBuf) > 0 Then
        SetClipText strBuf
    End If
    
e:
'    Application.ScreenUpdating = True

End Sub
Sub pasteMergeCellValue()
    Call pasteMergeCell(True)
End Sub
Sub pasteMergeCellFormula()
    Call pasteMergeCell(False)
End Sub
'--------------------------------------------------------------
' マージセルの代表セルの値ペースト
'--------------------------------------------------------------
Sub pasteMergeCell(ByVal blnValue As Boolean)

    Dim strBuf As String
    Dim strLine As Variant
    Dim strCell As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim a As Range
    Dim r As Range
    Dim strRange() As String
    
    If rlxCheckSelectRange = False Then
        Exit Sub
    End If

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    

    
    Application.ScreenUpdating = False
    
    On Error GoTo e
    
    '現在のリンク
    Dim bf As Range
    
    
    If Application.CutCopyMode <> xlCopy Then
    
        strBuf = rlxGetFileNameFromCli()
        If Len(strBuf) = 0 Then
            strBuf = GetClipText()
        End If
    
    Else
        'コピー元のRangeを取得
        Set bf = getCopyRange()
        If bf Is Nothing Then
            MsgBox "コピー元の取得に失敗しました。", vbOKOnly + vbExclamation, C_TITLE
            GoTo e
        End If
        
        If bf.CountLarge > 5000 Then
            MsgBox "大量のセルが選択されています。コピーするセルを5,000以下にしてください。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
        
        strBuf = copyMergeCell(bf, blnValue)
    End If
    
    If Len(strBuf) = 0 Then
        Exit Sub
    End If
    
    '貼り付けられる範囲をRangeで取得
    Set r = Selection(1)
    
    strLine = Split(strBuf, vbCrLf)

    l = 0
    For i = LBound(strLine) To UBound(strLine)

        k = 0
        strCell = Split(strLine(i), vbTab)
        
        Do Until r.Offset(l, k).MergeArea(1, 1).Address = r.Offset(l, k).Address
            l = l + 1
        Loop
        For j = LBound(strCell) To UBound(strCell)
            If strCell(j) <> vbNullChar Then
                Do Until r.Offset(l, k).MergeArea(1, 1).Address = r.Offset(l, k).Address
                    k = k + 1
                Loop
                If a Is Nothing Then
                    Set a = r.Offset(l, k)
                Else
                    Set a = Union(a, r.Offset(l, k))
                End If
                k = k + 1
            End If
        Next
        l = l + 1

    Next
    
    '選択セルを数える
    Dim ss As Range
    Dim sr As Range
    
    For Each ss In Selection
        If ss.MergeArea(1, 1).Address = ss.Address Then
            If sr Is Nothing Then
                Set sr = ss
            Else
                Set sr = Union(sr, ss)
            End If
        End If
    Next

    'コピー元が１セルでコピー先が複数セルの場合
    If a.Count = 1 And sr.Count > 1 Then
        sr.Select
'        Range(sr(1), sr(sr.count)).Select
    Else
        a.Select
'        Range(a(1), a(a.count)).Select
    End If
    
    '現在の選択位置をUndo用にバックアップ
    ThisWorkbook.Worksheets("Undo").Cells.Clear

    Set mUndo.sourceRange = Selection
    Set mUndo.destRange = ThisWorkbook.Worksheets("Undo").Range(Selection.Address)
    
    Dim rr As Range
    For Each rr In mUndo.sourceRange.Areas
        rr.Copy mUndo.destRange.Worksheet.Range(rr.Address)
    Next
    
    'コピー元が１セルでコピー先が複数セルの場合
    If a.Count = 1 And sr.Count > 1 Then
        '現在の選択セルにコピー
        Dim p As Range
        For Each p In sr
            p.FormulaLocal = delQuat(strBuf)
        Next
    Else
        '現在の選択位置から右下方面へコピー
        strLine = Split(strBuf, vbCrLf)
    
        l = 0
        For i = LBound(strLine) To UBound(strLine)
    
            k = 0
            strCell = Split(strLine(i), vbTab)
            
            Do Until r.Offset(l, k).MergeArea(1, 1).Address = r.Offset(l, k).Address
                l = l + 1
            Loop
            For j = LBound(strCell) To UBound(strCell)
                If strCell(j) <> vbNullChar Then
                    Do Until r.Offset(l, k).MergeArea(1, 1).Address = r.Offset(l, k).Address
                        k = k + 1
                    Loop
                    r.Offset(l, k).FormulaLocal = delQuat(strCell(j))
                    k = k + 1
                End If
            Next
            l = l + 1
        Next
    End If
    
    '元の選択セルをコピー状態にする。
    If bf Is Nothing Then
    Else
        bf.Copy
    End If
    
    'コピー元が１セルでコピー先が複数セルの場合
    If a.Count = 1 And sr.Count > 1 Then
        sr.Select
'        Range(sr(1), sr(sr.count).MergeArea(sr(sr.count).MergeArea.count)).Select
    Else
        a.Select
'        Range(a(1), a(a.count).MergeArea(a(a.count).MergeArea.count)).Select
    End If
    
    'Undo
    Application.OnUndo "Undo", MacroHelper.BuildPath("execUndo")
e:
    Application.ScreenUpdating = True
End Sub


''--------------------------------------------------------------
'' マージセルの代表セルの式コピー
''--------------------------------------------------------------
'Private Function copyMergeCell(ByRef s As Range, ByVal blnValue As Boolean, ByVal blnRotate As Boolean) As String
'
'    Dim i As Long
'    Dim j As Long
'    Dim iMax As Long
'    Dim jMax As Long
'    Dim r As Range
'
'    Dim strLine As String
'    Dim strBuf As String
'
'    If rlxCheckSelectRange = False Then
'        Exit Function
'    End If
'
'    If ActiveCell Is Nothing Then
'        Exit Function
'    End If
'
'    On Error GoTo e
'
'
''    If blnRotate Then
''        iMax = s(s.count).Column - s(1).Column + 1
''        jMax = s(s.count).Row - s(1).Row + 1
''    Else
''        iMax = s(s.count).Row - s(1).Row + 1
''        jMax = s(s.count).Column - s(1).Column + 1
''    End If
''
''    For i = 1 To iMax
''
''        strLine = ""
''        For j = 1 To jMax
''
''
''            If blnRotate Then
''                Set r = s(j, i)
''            Else
''                Set r = s(i, j)
''            End If
'
'    iMax = s(s.count).Row - s(1).Row + 1
'    jMax = s(s.count).Column - s(1).Column + 1
'
'
'
'    Dim dr As Object
'    Dim dc As Object
'    Dim Key As String
'
'    Set dr = CreateObject("Scripting.Dictionary")
'    Set dc = CreateObject("Scripting.Dictionary")
'
'    For i = 1 To iMax
'        strLine = ""
'        For j = 1 To jMax
'            Set r = s(i, j)
'            If r.MergeArea(1, 1).Address = r.Address Then
'
'                Key = CStr(i)
'                If Not dr.Exists(Key) Then
'                    dr.Add Key, r
'                End If
'
'                Key = CStr(j)
'                If Not dc.Exists(Key) Then
'                    dc.Add Key, r
'                End If
'
'            End If
'        Next
'    Next
'
'
'    ReDim ddd(1 To dr.count, 1 To dc.count)
'
'
'
'    For i = 1 To iMax
'
'        strLine = ""
'        For j = 1 To jMax
'
'            Set r = s(i, j)
'
'            'マージセルなら左上のみ処理
'            If r.MergeArea(1, 1).Address = r.Address Then
'
'                If Len(strLine) = 0 Then
'                    If blnValue Then
'                        strLine = addQuat(r.Value)
'                    Else
'                        strLine = addQuat(r.FormulaLocal)
'                    End If
'                Else
'                    If blnValue Then
'                        strLine = strLine & vbTab & addQuat(r.Value)
'                    Else
'                        strLine = strLine & vbTab & addQuat(r.FormulaLocal)
'                    End If
'                End If
'            End If
'
'        Next
'
'        If strLine <> "" Then
'            If Len(strBuf) = 0 Then
'                strBuf = strLine
'            Else
'                strBuf = strBuf & vbCrLf & strLine
'            End If
'        End If
'    Next
'
'e:
'    If Len(strBuf) > 0 Then
'        copyMergeCell = strBuf
'    End If
'
'End Function
'--------------------------------------------------------------
' マージセルの代表セルの式コピー
'--------------------------------------------------------------
Private Function copyMergeCell(ByRef s As Range, ByVal blnValue As Boolean) As String
    
    Dim i As Long
    Dim j As Long
    
    Dim r As Range
    Dim di As DictionaryEx
    Dim dj As DictionaryEx
    
    Dim Key As String
    
    On Error GoTo e
    
'    Set di = CreateObject("Scripting.Dictionary")
'    Set dj = CreateObject("Scripting.Dictionary")
    Set di = New DictionaryEx
    Set dj = New DictionaryEx
    
    'コピー元セルで最小限必要なセルの数を求める
    For i = 1 To s.Rows.Count
        For j = 1 To s.Columns.Count
            
            Set r = s(i, j)
            
            If r.Address = r.MergeArea(1, 1).Address Then
            
                Key = CStr(i)
                If Not di.Exists(Key) Then
                    di.Add Key, Key
                End If
                
                Key = CStr(j)
                If Not dj.Exists(Key) Then
                    dj.Add Key, Key
                End If
            
            End If
        Next
    Next
    
    Dim tbl() As String
    
    '必要な分だけ配列を作成
    ReDim tbl(0 To di.Count - 1, 0 To dj.Count - 1)

    '初期値（結合セルの代表以外の値になる）
    For i = 0 To di.Count - 1
        For j = 0 To dj.Count - 1
            tbl(i, j) = vbNullChar
        Next
    Next
    
    Dim ai As Variant
    Dim aj As Variant
    
    ai = di.keys
    aj = dj.keys

    'Rangeの値を配列にコピー
    For i = 0 To di.Count - 1
        For j = 0 To dj.Count - 1
            Set r = s(CLng(ai(i)), CLng(aj(j)))
            If blnValue Then
                tbl(i, j) = addQuat(r.Value)
            Else
                tbl(i, j) = addQuat(r.FormulaLocal)
            End If
        Next
    Next
    
    '作成した配列から文字列に変換
    Dim strBuf As String
    Dim strLine As String
    
    strBuf = ""
    For i = 0 To di.Count - 1
        
        strLine = ""
        For j = 0 To dj.Count - 1
        
            If j = 0 Then
                strLine = tbl(i, j)
            Else
                strLine = strLine & vbTab & tbl(i, j)
            End If
        Next
        If i = 0 Then
            strBuf = strLine
        Else
            strBuf = strBuf & vbCrLf & strLine
        End If
    Next
    
e:
    Set di = Nothing
    Set dj = Nothing
    
    If Len(strBuf) > 0 Then
        copyMergeCell = strBuf
    End If

End Function

Private Function addQuat(ByVal strVal As String) As String

    If InStr(strVal, vbLf) > 0 Then
        addQuat = """" & strVal & """"
    Else
        addQuat = strVal
    End If

End Function
Private Function delQuat(ByVal strVal As String) As String

    Dim strBuf As String

    strBuf = strVal

    If Left$(strBuf, 1) = """" Then
        strBuf = Mid$(strBuf, 2)
    End If
    
    If Right$(strBuf, 1) = """" Then
        strBuf = Mid$(strBuf, 1, Len(strBuf) - 1)
    End If

    delQuat = strBuf

End Function
'ObjectLinkの形からRangeオブジェクトを作成
Private Function getCopyRange() As Range

    '現在のリンク
    Dim bf As Range
    Dim strRange As Variant
    
    strRange = Split(getObjectLink(), vbTab)
    If UBound(strRange) = -1 Then
        Exit Function
    End If
    
    If InStr(strRange(0), "Excel") = 0 Then
        Exit Function
    End If
    
    On Error Resume Next
    Err.Clear
    Set bf = Range("'[" & rlxGetFullpathFromFileName(strRange(1)) & "]" & Mid$(strRange(2), 1, InStr(strRange(2), "!") - 1) & "'!" & Application.ConvertFormula(Mid$(strRange(2), InStr(strRange(2), "!") + 1), xlR1C1, xlA1))
    If Err.Number <> 0 Then
        Set bf = Nothing
    End If

    Set getCopyRange = bf
    
End Function

