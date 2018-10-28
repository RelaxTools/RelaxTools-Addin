VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmComp 
   Caption         =   "シート比較"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   OleObjectBlob   =   "frmComp.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [RelaxTools-Addin] v4
'
' Copyright (c) 2009 Yasuhiro Watanabe
' https://github.com/RelaxTools/RelaxTools-Addin
' author:relaxtools@opensquare.net
'
' The MIT License (MIT)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'-----------------------------------------------------------------------------------------------------


Option Explicit
Private mblnCancel As Boolean
Private Const C_START_ROW As Long = 8

Private Const C_COMP_NO As Long = 1
Private Const C_COMP_RESULT As Long = 2
Private Const C_COMP_SRCSTR As Long = 3
Private Const C_COMP_DSTSTR As Long = 4
Private Const C_COMP_BOOK As Long = 5
Private Const C_COMP_SHEET As Long = 6
Private Const C_COMP_ADDRESS As Long = 7

Private Sub cboDstBook_Change()
    Dim s As Worksheet
    
    cboDstSheet.Clear
    If cboDstBook.ListIndex <> -1 Then
        For Each s In Workbooks(cboDstBook.Text).Worksheets
            cboDstSheet.AddItem s.Name
        Next
    End If
End Sub

Private Sub cboSrcBook_Change()
    Dim s As Worksheet
    
    cboSrcSheet.Clear
    If cboSrcBook.ListIndex <> -1 Then
        For Each s In Workbooks(cboSrcBook.Text).Worksheets
            cboSrcSheet.AddItem s.Name
        Next
    End If
End Sub

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = "閉じる" Then
        Unload Me
    Else
        mblnCancel = True
    End If
End Sub



'Private Sub cmdOk_Click()
'
'    Dim srcSheet As Worksheet
'    Dim dstSheet As Worksheet
'
'    If cboSrcSheet.ListIndex = -1 Then
'        MsgBox "比較元のシートを入力してください。", vbOKOnly + vbExclamation, C_TITLE
'        Exit Sub
'    End If
'    If cboDstSheet.ListIndex = -1 Then
'        MsgBox "比較先のシートを入力してください。", vbOKOnly + vbExclamation, C_TITLE
'        Exit Sub
'    End If
'
'    Set srcSheet = Workbooks(cboSrcBook.Text).Worksheets(cboSrcSheet.Text)
'    Set dstSheet = Workbooks(cboDstBook.Text).Worksheets(cboDstSheet.Text)
'
'    Dim srcBook As Workbook
'    Dim dstBook As Workbook
'
'    Dim r1 As Range
'    Dim d1 As Variant
'    Dim r2 As Range
'    Dim d2 As Variant
'
'    Dim i As Long
'    Dim j As Long
'    Dim lngCount As Long
'
'    Dim strSrcAddress As String
'    Dim strDstAddress As String
'
'    Set srcBook = srcSheet.Parent
'    Set dstBook = dstSheet.Parent
'
'    strSrcAddress = srcSheet.UsedRange.Address
'    strDstAddress = dstSheet.UsedRange.Address
'
'    Set r1 = Union(dstSheet.Range(strSrcAddress), dstSheet.Range(strDstAddress))
'    d1 = r1
'
'    Set r2 = Union(srcSheet.Range(strSrcAddress), srcSheet.Range(strDstAddress))
'    d2 = r2
'
'
'
'    Dim ResultWS As Worksheet
'    ThisWorkbook.Worksheets("比較結果").Copy
'    Set ResultWS = ActiveSheet
'
'    ResultWS.Name = "比較結果"
'
'    ResultWS.Cells(1, C_COMP_NO).Value = "シートの比較"
'    ResultWS.Cells(2, C_COMP_NO).Value = "比較元：" & cboSrcBook.Text & "!" & cboSrcSheet.Text
'    ResultWS.Cells(3, C_COMP_NO).Value = "比較先：" & cboDstBook.Text & "!" & cboDstSheet.Text
'    ResultWS.Cells(4, C_COMP_NO).Value = "不一致の比較元の背景色を変更する（黄）：" & chkSrcColor.Value
'    ResultWS.Cells(5, C_COMP_NO).Value = "不一致の比較先の背景色を変更する（赤）：" & chkDstColor.Value
'
'    ResultWS.Cells(7, C_COMP_NO).Value = "No."
'    ResultWS.Cells(7, C_COMP_RESULT).Value = "結果"
'    ResultWS.Cells(7, C_COMP_SRCSTR).Value = "比較元文字列"
'    ResultWS.Cells(7, C_COMP_DSTSTR).Value = "比較先文字列"
'    ResultWS.Cells(7, C_COMP_BOOK).Value = "比較先ブック"
'    ResultWS.Cells(7, C_COMP_SHEET).Value = "比較先シート"
'    ResultWS.Cells(7, C_COMP_ADDRESS).Value = "アドレス"
'    lngCount = C_START_ROW
'
'    If IsEmpty(d1) Or IsEmpty(d2) Then
'        GoTo e
'    End If
'
'    Dim mm As MacroManager
'    Set mm = New MacroManager
'    Set mm.Form = Me
'    mm.Disable
'    mm.DispGuidance "セル数をカウントしています..."
'
'    mm.StartGauge (UBound(d1, 1) - LBound(d1, 1) + 1) * (UBound(d1, 2) - LBound(d1, 2) + 1)
'    Dim lngCnt As Long
'    lngCnt = 0
'
'    For i = LBound(d1, 1) To UBound(d1, 1)
'        For j = LBound(d1, 2) To UBound(d1, 2)
'            If mblnCancel Then
'                GoTo e
'            End If
'            If d1(i, j) <> d2(i, j) Then
'                makeResult ResultWS, srcSheet, dstSheet, lngCount, i, j
'
'            Else
'                '空セル対策
'                If IsEmpty(d1(i, j)) <> IsEmpty(d2(i, j)) Then
'                    makeResult ResultWS, srcSheet, dstSheet, lngCount, i, j
'                End If
'            End If
'            lngCnt = lngCnt + 1
'            mm.DisplayGauge lngCnt
'        Next
'    Next
'
'
'    ResultWS.Columns("B:G").AutoFit
'    Dim r As Range
'    Set r = ResultWS.Cells(C_START_ROW, 1).CurrentRegion
'
'    r.VerticalAlignment = xlTop
'    r.Select
'
'    execSelectionRowDrawGrid
'e:
'    Set ResultWS = Nothing
'    Unload Me
'
'End Sub
'Sub makeResult(ByRef ResultWS As Worksheet, ByRef srcSheet As Worksheet, ByRef dstSheet As Worksheet, ByRef lngCount As Long, ByVal i As Long, ByVal j As Long)
'
'    ResultWS.Cells(lngCount, C_COMP_NO).Value = lngCount - C_START_ROW + 1
'    ResultWS.Cells(lngCount, C_COMP_RESULT).Value = "不一致"
'    ResultWS.Cells(lngCount, C_COMP_BOOK).Value = dstSheet.Parent.Name
'    ResultWS.Cells(lngCount, C_COMP_SHEET).Value = dstSheet.Name
'    ResultWS.Cells(lngCount, C_COMP_ADDRESS).Value = dstSheet.Cells(i, j).Address
'    ResultWS.Cells(lngCount, C_COMP_SRCSTR).Value = srcSheet.Cells(i, j).Value
'    ResultWS.Cells(lngCount, C_COMP_DSTSTR).Value = dstSheet.Cells(i, j).Value
'
'    ResultWS.Hyperlinks.Add _
'        Anchor:=ResultWS.Cells(lngCount, C_COMP_ADDRESS), _
'        Address:="", _
'        SubAddress:=ResultWS.Cells(lngCount, C_COMP_ADDRESS).Address, _
'        TextToDisplay:=dstSheet.Cells(i, j).Address
'
'    On Error Resume Next
'    If chkSrcColor Then
'        srcSheet.Cells(i, j).Interior.Color = vbYellow
'    End If
'    If chkDstColor Then
'        dstSheet.Cells(i, j).Interior.Color = vbRed
'    End If
'
'    lngCount = lngCount + 1
'
'End Sub
Private Sub UserForm_Initialize()

    Dim b As Workbook
    
    For Each b In Workbooks
        cboSrcBook.AddItem b.Name
        cboDstBook.AddItem b.Name
    Next
    
    cboSrcSheet.Clear
    cboDstSheet.Clear
    
    chkSrcColor.Value = True
    chkDstColor.Value = True

    lblGauge.visible = False
    mblnCancel = False
    
End Sub


Private Sub cmdOk_Click()

    Dim srcSheet As Worksheet
    Dim dstSheet As Worksheet

    If cboSrcSheet.ListIndex = -1 Then
        MsgBox "比較元のシートを入力してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    If cboDstSheet.ListIndex = -1 Then
        MsgBox "比較先のシートを入力してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    If cboSrcBook.Text = cboDstBook.Text And cboSrcSheet.Text = cboDstSheet.Text Then
        MsgBox "比較元と比較先が同じです。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If

    Set srcSheet = Workbooks(cboSrcBook.Text).Worksheets(cboSrcSheet.Text)
    Set dstSheet = Workbooks(cboDstBook.Text).Worksheets(cboDstSheet.Text)

    Dim srcBook As Workbook
    Dim dstBook As Workbook

    Dim r1 As Range
    Dim r2 As Range
    
    Dim i As Long
    Dim j As Long
    Dim lngCount As Long
    
    Dim strSrcAddress As String
    Dim strDstAddress As String

    Set srcBook = srcSheet.Parent
    Set dstBook = dstSheet.Parent
    
    strSrcAddress = srcSheet.UsedRange.Address
    strDstAddress = dstSheet.UsedRange.Address

'    Set r1 = Union(dstSheet.Range(strSrcAddress), dstSheet.Range(strDstAddress))
'    Set r2 = Union(srcSheet.Range(strSrcAddress), srcSheet.Range(strDstAddress))
    Set r1 = Union(srcSheet.Range(strSrcAddress), srcSheet.Range(strDstAddress))
    Set r2 = Union(dstSheet.Range(strSrcAddress), dstSheet.Range(strDstAddress))
    
    Dim ResultWS As Worksheet
    ThisWorkbook.Worksheets("比較結果").Copy
    Set ResultWS = Application.Workbooks(Application.Workbooks.Count).Worksheets(1)
    
    ResultWS.Name = "比較結果"
    
    ResultWS.Cells(1, C_COMP_NO).Value = "シートの比較"
    ResultWS.Cells(2, C_COMP_NO).Value = "比較元：" & cboSrcBook.Text & "!" & cboSrcSheet.Text
    ResultWS.Cells(3, C_COMP_NO).Value = "比較先：" & cboDstBook.Text & "!" & cboDstSheet.Text
    ResultWS.Cells(4, C_COMP_NO).Value = "不一致の比較「元」の背景色を変更する（黄）：" & chkSrcColor.Value
    ResultWS.Cells(5, C_COMP_NO).Value = "不一致の比較「先」の背景色を変更する（赤）：" & chkDstColor.Value
    
    ResultWS.Cells(7, C_COMP_NO).Value = "No."
    ResultWS.Cells(7, C_COMP_RESULT).Value = "結果"
    ResultWS.Cells(7, C_COMP_SRCSTR).Value = "比較元文字列"
    ResultWS.Cells(7, C_COMP_DSTSTR).Value = "比較先文字列"
    ResultWS.Cells(7, C_COMP_BOOK).Value = "比較先ブック"
    ResultWS.Cells(7, C_COMP_SHEET).Value = "比較先シート"
    ResultWS.Cells(7, C_COMP_ADDRESS).Value = "アドレス"
    lngCount = C_START_ROW
    
    If r1 Is Nothing Or r2 Is Nothing Then
        GoTo e
    End If
    
    Dim mm As MacroManager
    Set mm = New MacroManager
    Set mm.Form = Me
    
    mm.Disable
    mm.DispGuidance "セル数をカウントしています..."
    
    mm.StartGauge r1.Count
    
    For i = 1 To r1.Count
            
        If mblnCancel Then
            GoTo e
        End If
        
        If IsError(r1(i).Value) Or IsError(r2(i).Value) Then
            makeResult ResultWS, srcSheet, dstSheet, lngCount, r1(i), r2(i)
        Else
            '空セル対策
            If IsEmpty(r1(i).Value) And IsEmpty(r2(i).Value) Then
            Else
                If r1(i).Value <> r2(i).Value Then
                    makeResult ResultWS, srcSheet, dstSheet, lngCount, r1(i), r2(i)
                End If
            End If
        End If
        
        mm.DisplayGauge i
    
    Next
    
    ResultWS.Columns("B:G").AutoFit
    ResultWS.Select
    
    Dim r As Range
    Set r = ResultWS.Cells(C_START_ROW, 1).CurrentRegion
    
    r.VerticalAlignment = xlTop
    r.Select
    
    execSelectionRowDrawGrid
e:
    Unload Me
    
    ResultWS.Parent.Activate
    Set ResultWS = Nothing
    
End Sub
Sub makeResult(ByRef ResultWS As Worksheet, ByRef srcSheet As Worksheet, ByRef dstSheet As Worksheet, ByRef lngCount As Long, ByRef r1 As Range, ByRef r2 As Range)
                
    ResultWS.Cells(lngCount, C_COMP_NO).Value = lngCount - C_START_ROW + 1
    ResultWS.Cells(lngCount, C_COMP_RESULT).Value = "不一致"
    ResultWS.Cells(lngCount, C_COMP_BOOK).Value = dstSheet.Parent.Name
    ResultWS.Cells(lngCount, C_COMP_SHEET).Value = dstSheet.Name
    ResultWS.Cells(lngCount, C_COMP_ADDRESS).Value = r1.Address
    ResultWS.Cells(lngCount, C_COMP_SRCSTR).Value = r1.Value
    ResultWS.Cells(lngCount, C_COMP_DSTSTR).Value = r2.Value

    ResultWS.Hyperlinks.Add _
        Anchor:=ResultWS.Cells(lngCount, C_COMP_ADDRESS), _
        Address:="", _
        SubAddress:=ResultWS.Cells(lngCount, C_COMP_ADDRESS).Address, _
        TextToDisplay:=r1.Address

    On Error Resume Next
    If chkSrcColor Then
        r1.Interior.Color = vbYellow
    End If
    If chkDstColor Then
        r2.Interior.Color = vbRed
    End If

    lngCount = lngCount + 1

End Sub

