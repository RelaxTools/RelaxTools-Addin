Attribute VB_Name = "basMergeCell"
Option Explicit
'Sub key()
'    Application.OnKey "^%{RIGHT}", "SizeToWidest"
'    Application.OnKey "^%{LEFT}", "SizeToNarrowest"
'    Application.OnKey "^%{UP}", "SizeToShortest"
'    Application.OnKey "^%{DOWN}", "SizeToTallest"
'
'End Sub

'SizeToWidest
Sub SizeToWidest()

    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    Dim i As Long
    Dim j As Long

    Dim blnMerge As Boolean
    Dim blnValue As Boolean
    
    Dim strSel As String
    
    On Error GoTo e
    
    Dim blnOnly1 As Boolean
    
    Application.CutCopyMode = False
    
    If Selection(1).MergeArea.Columns.count = 1 Then
        blnOnly1 = True
    End If
    
    On Error GoTo e
    
    strSel = Selection.Address

    lngLeft = Selection(1).Column
    lngTop = Selection(1).Row
    lngBottom = Selection(Selection.count).Row
    lngRight = Selection(Selection.count).Column + 1

    For j = lngRight To Cells.Columns.count
        blnMerge = False
        blnValue = False
        For i = lngTop To lngBottom

            If Cells(i, j).MergeCells Then
                blnMerge = True
                Exit For
            End If
            If Cells(i, j).Value <> "" Then
                blnValue = True
                Exit For
            End If

        Next
        
        If blnMerge = False And blnValue = False Then

            Dim r As Range
            Dim s As Range
            
            Set r = Range(Cells(lngTop, lngLeft), Cells(lngBottom, j - 1))

            Application.ScreenUpdating = False

            With ThisWorkbook.Worksheets("Work")

                .Cells.Clear

                r.Cut Destination:=.Range(r.Address)

                .Columns(r.Columns(2).Column).Insert Shift:=xlToRight
                
                If blnOnly1 Then
                    For i = lngTop To lngBottom
            
                        If .Cells(i, lngLeft).Address = .Cells(i, lngLeft).MergeArea(1).Address Then
'                            '書式維持のためのコピー
'                            .Cells(i, lngLeft).MergeArea.Copy .Cells(i, lngLeft).MergeArea.Offset(, 1)
'                            Application.DisplayAlerts = False
                            .Cells(i, lngLeft).MergeArea.Resize(, 2).Merge
'                            Application.DisplayAlerts = True
                        End If
            
                    Next
                End If

                Set s = .Range(r.Address).Resize(, r.Columns.count + 1)

                s.Cut Destination:=Range(s.Address)
                
                Set s = Range(strSel)
                s.Resize(, s.Columns.count + 1).Select

            End With

            Application.ScreenUpdating = True

            Exit For

        End If
    Next
    
    Exit Sub
e:
    MsgBox "他の結合セルに影響するため実行できません。", vbOKOnly + vbExclamation, C_TITLE

End Sub
'SizeToNarrowest
Sub SizeToNarrowest()

    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    Dim i As Long
    Dim j As Long

    Dim blnMerge As Boolean
    
    Dim strSel As String
    
    On Error GoTo e
    
    If Selection(1).MergeArea.Columns.count <= 1 Then
        Exit Sub
    End If
    
    Application.CutCopyMode = False

    strSel = Selection.Address

    lngLeft = Selection(1).Column
    lngTop = Selection(1).Row
    lngBottom = Selection(Selection.count).Row
    lngRight = Selection(Selection.count).Column + 1

    For j = lngRight To Cells.Columns.count
        blnMerge = False
        For i = lngTop To lngBottom

            If Cells(i, j).MergeCells Then
                blnMerge = True
                Exit For
            End If

        Next
        If blnMerge = False Then

            Dim r As Range
            Dim s As Range
            
            Set r = Range(Cells(lngTop, lngLeft), Cells(lngBottom, j - 1))

            Application.ScreenUpdating = False

            With ThisWorkbook.Worksheets("Work")

                .Cells.Clear

                r.Cut Destination:=.Range(r.Address)

                .Columns(r.Columns(2).Column).Delete Shift:=xlToLeft

                Set s = .Range(r.Address).Resize(, r.Columns.count - 1)

                s.Cut Destination:=Range(s.Address)
                
            
            End With

            Set s = Range(strSel)
            s.Resize(, s.Columns.count - 1).Select
            
            Application.ScreenUpdating = True

            Exit For

        End If
    Next

    Exit Sub
e:
    MsgBox "他の結合セルに影響するため実行できません。", vbOKOnly + vbExclamation, C_TITLE
End Sub
'SizeToTallest
Sub SizeToTallest()

    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    Dim i As Long
    Dim j As Long

    Dim blnMerge As Boolean
    Dim blnValue As Boolean

    Dim strSel As String
    
    On Error GoTo e
    
    Dim blnOnly1 As Boolean
    
    Application.CutCopyMode = False
    
    If Selection(1).MergeArea.Rows.count = 1 Then
        blnOnly1 = True
    End If
    
    strSel = Selection.Address

    lngLeft = Selection(1).Column
    lngTop = Selection(1).Row
    lngBottom = Selection(Selection.count).Row + 1
    lngRight = Selection(Selection.count).Column

    For i = lngBottom To Cells.Rows.count
        blnMerge = False
        blnValue = False
        For j = lngLeft To lngRight

            If Cells(i, j).MergeCells Then
                blnMerge = True
                Exit For
            End If
            If Cells(i, j).Value <> "" Then
                blnValue = True
                Exit For
            End If

        Next
        If blnMerge = False And blnValue = False Then

            Dim r As Range
            Dim s As Range
            
            Set r = Range(Cells(lngTop, lngLeft), Cells(i - 1, lngRight))

            Application.ScreenUpdating = False

            With ThisWorkbook.Worksheets("Work")

                .Cells.Clear

                r.Cut Destination:=.Range(r.Address)

                .Rows(r.Rows(2).Row).Insert Shift:=xlDown
                
                If blnOnly1 Then
                    For j = lngLeft To lngRight
            
                        If .Cells(lngTop, j).Address = .Cells(lngTop, j).MergeArea(1).Address Then
'                            '書式維持のためのコピー
'                            .Cells(lngTop, j).MergeArea.Copy .Cells(lngTop, j).MergeArea.Offset(1)
'                            Application.DisplayAlerts = False
                            .Cells(lngTop, j).MergeArea.Resize(2).Merge
'                            Application.DisplayAlerts = True
                        End If
            
                    Next
                End If

                Set s = .Range(r.Address).Resize(r.Rows.count + 1)

                s.Cut Destination:=Range(s.Address)

            End With
            
            Set s = Range(strSel)
            s.Resize(s.Rows.count + 1).Select

            Application.ScreenUpdating = True

            Exit For

        End If
    Next
    Exit Sub
e:
    MsgBox "他の結合セルに影響するため実行できません。", vbOKOnly + vbExclamation, C_TITLE
End Sub
'SizeToShortest
Sub SizeToShortest()

    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    Dim i As Long
    Dim j As Long

    Dim blnMerge As Boolean
    
    Dim strSel As String
    
    On Error GoTo e
    
    If Selection(1).MergeArea.Rows.count <= 1 Then
        Exit Sub
    End If
    
    Application.CutCopyMode = False
    
    strSel = Selection.Address

    lngLeft = Selection(1).Column
    lngTop = Selection(1).Row
    lngBottom = Selection(Selection.count).Row + 1
    lngRight = Selection(Selection.count).Column

    For i = lngBottom To Cells.Rows.count
        blnMerge = False
        For j = lngLeft To lngRight

            If Cells(i, j).MergeCells Then
                blnMerge = True
                Exit For
            End If

        Next
        If blnMerge = False Then

            Dim r As Range
            Dim s As Range
            
            Set r = Range(Cells(lngTop, lngLeft), Cells(i - 1, lngRight))

            Application.ScreenUpdating = False

            With ThisWorkbook.Worksheets("Work")

                .Cells.Clear

                r.Cut Destination:=.Range(r.Address)

                .Rows(r.Rows(2).Row).Delete Shift:=xlUp

                Set s = .Range(r.Address).Resize(r.Rows.count - 1)

                s.Cut Destination:=Range(s.Address)
                
            End With

            Set s = Range(strSel)
            s.Resize(s.Rows.count - 1).Select
            
            Application.ScreenUpdating = True

            Exit For

        End If
    Next
    Exit Sub
e:
    MsgBox "他の結合セルに影響するため実行できません。", vbOKOnly + vbExclamation, C_TITLE
End Sub
