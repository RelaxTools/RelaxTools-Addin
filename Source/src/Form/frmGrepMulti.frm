VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrepMulti 
   Caption         =   "ExcelファイルのGrep(マルチプロセス版)"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   OleObjectBlob   =   "frmGrepMulti.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmGrepMulti"
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
Private mRange As Range
Private mblnSelectMode As Boolean
Private mobjRegx As Object
Private mlngCount As Long
Private mblnCancel As Boolean

Private mblnRefresh As Boolean

Private Const C_START_ROW As Long = 11
Private Const C_SEARCH_NO As Long = 1
Private Const C_SEARCH_BOOK As Long = 2
Private Const C_SEARCH_SHEET As Long = 3
Private Const C_SEARCH_ADDRESS As Long = 4
Private Const C_SEARCH_STR As Long = 5
'Private Const C_SEARCH_ID As Long = 6

Private Const C_SEARCH_OBJECT_CELL = "セルのみ"
Private Const C_SEARCH_OBJECT_SHAPE = "シェイプのみ"
Private Const C_SEARCH_OBJECT_CELL_AND_SHAPE = "セル＆シェイプ"
Private Const C_SEARCH_VALUE_VALUE = "値"
Private Const C_SEARCH_VALUE_FORMULA = "式"

Private mMm As MacroManager

Private mtray As TaskTrayView

Private Sub chkOffset_Change()

    txtRow.enabled = chkOffset.Value
    txtCol.enabled = chkOffset.Value

End Sub

Private Sub chkRegEx_Change()
'    chkZenHan.enabled = Not (chkRegEx.Value)
End Sub

'Private Sub chkRegEx_Change()
'    chkCase.enabled = chkRegEx.Value
'    If chkRegEx.Value = False Then
'        chkCase.Value = False
'    End If
'End Sub

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = "閉じる" Then
        Unload Me
        If ThisWorkbook.IsAddin Then
            Application.Quit
        End If
    Else
        mblnCancel = True
    End If
End Sub

Private Sub cmdFolder_Click()

    Dim strFile As String

    'フォルダ名取得
    strFile = rlxSelectFolder()
    
    If Trim(strFile) <> "" Then
        cboFolder.Text = strFile
    End If
    
End Sub

Private Sub cmdHelp_Click()
    
    If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    Call GotoRegExpHelp
    
End Sub

Private Sub cmdOk_Click()

    Dim XL As Excel.Application
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim colBook As Collection
    Dim varBook As Variant
    Dim objFs As Object
    Dim lngBookCount As Long
    Dim lngBookMax As Long
    
    Dim ResultWS As Worksheet
    
    Dim strPath As String
    Dim strPatterns() As String
    
    Dim o As Object
    Set o = CreateObject("VBScript.RegExp")
    
    If Len(Trim(cboSearch.Text)) = 0 Then
        MsgBox "検索文字列を指定してください...", vbExclamation, C_TITLE
        cboSearch.SetFocus
        Exit Sub
    End If
    
    
    '正規表現の場合
    If chkRegEx.Value Then
        Err.Clear
        On Error Resume Next
        o.Pattern = cboSearch.Text
        o.IgnoreCase = Not (chkCase.Value)
        o.Global = True
        o.Execute ""
        If Err.Number <> 0 Then
            MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
            cboSearch.SetFocus
            Exit Sub
        End If
    End If
    
    If chkOffset.Value Then
        
        o.Pattern = "([0-9]+|\-[0-9]+)"
        o.IgnoreCase = False
        o.Global = True
        
        If Not o.Test(txtRow.Value) Then
            MsgBox "行には数値を入れてください。", vbExclamation + vbOKOnly, C_TITLE
            txtRow.SetFocus
            Exit Sub
        End If
        If Not o.Test(txtCol.Value) Then
            MsgBox "列には数値を入れてください。", vbExclamation + vbOKOnly, C_TITLE
            txtCol.SetFocus
            Exit Sub
        End If
        
        Dim lngRow As Long
        Dim lngCol As Long
        
        lngRow = Val(txtRow.Value)
        lngCol = Val(txtCol.Value)
        
        Select Case lngRow
            Case -100 To 100
            Case Else
                MsgBox "行には-100～100を入れてください。", vbExclamation + vbOKOnly, C_TITLE
                txtRow.SetFocus
                Exit Sub
        End Select
        Select Case lngCol
            Case -100 To 100
            Case Else
                MsgBox "列には-100～100を入れてください。", vbExclamation + vbOKOnly, C_TITLE
                txtCol.SetFocus
                Exit Sub
        End Select
        
        
    End If
    
    strPath = cboFolder.Text
    strPatterns = Split(cboPattern.Text, ";")

    Set colBook = New Collection
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    Set mMm = New MacroManager
    Set mMm.Form = Me
    mMm.Disable
    mMm.DispGuidance "ファイルの数をカウントしています..."
    
    FileSearch objFs, strPath, strPatterns(), colBook
    Select Case Err.Number
    Case 75, 76
        mMm.Enable
        Set mMm = Nothing
        MsgBox "フォルダが存在しません。", vbExclamation, "ExcelGrep"
        cboFolder.SetFocus
        Exit Sub
    End Select
    
    
    Set objFs = Nothing
    
    ThisWorkbook.Worksheets("Grep結果").Copy
    DoEvents
    Set ResultWS = Application.Workbooks(Application.Workbooks.count).Worksheets(1)
    ResultWS.Name = "Grep結果"
    
    ResultWS.Cells(1, C_SEARCH_NO).Value = "ExcelファイルのGrep"
    ResultWS.Cells(2, C_SEARCH_NO).Value = "条件：" & cboSearch.Text
    ResultWS.Cells(3, C_SEARCH_NO).Value = "ファイル：" & cboPattern.Text
    ResultWS.Cells(4, C_SEARCH_NO).Value = "フォルダ：" & cboFolder.Text
    ResultWS.Cells(5, C_SEARCH_NO).Value = "検索オブジェクト：" & cboObj.Text
    ResultWS.Cells(6, C_SEARCH_NO).Value = "検索対象：" & cboValue.Text
    ResultWS.Cells(7, C_SEARCH_NO).Value = "正規表現：" & chkRegEx.Value
    ResultWS.Cells(8, C_SEARCH_NO).Value = "読取パスワード：" & txtPassword.Text
    
    ResultWS.Cells(10, C_SEARCH_NO).Value = "No."
    ResultWS.Cells(10, C_SEARCH_BOOK).Value = "ブック名"
    ResultWS.Cells(10, C_SEARCH_BOOK).ColumnWidth = 60
    ResultWS.Cells(10, C_SEARCH_SHEET).Value = "シート名"
    ResultWS.Cells(10, C_SEARCH_ADDRESS).Value = "セル/シェイプ"
    ResultWS.Cells(10, C_SEARCH_STR).Value = "検索文字列"
    mlngCount = C_START_ROW

    cmdCancel.Caption = "キャンセル"
    
    Set XL = New Excel.Application
    
    AppActivate Me.Caption
    
    lngBookCount = 0
    lngBookMax = colBook.count
    mMm.StartGauge lngBookMax
    
    XL.DisplayAlerts = False
    XL.EnableEvents = False
    
    Dim varPassword As Variant
    Dim pass As Variant
    
    If Len(txtPassword.Text) <> 0 Then
        varPassword = Split(txtPassword.Text, ",")
    Else
        varPassword = Array("")
    End If
 
    For Each varBook In colBook
    
        If mblnCancel Then
            Exit For
        End If
    
        For Each pass In varPassword
            Err.Clear
            Set WB = XL.Workbooks.Open(filename:=varBook, ReadOnly:=True, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True, notify:=False, Password:=pass, local:=True)
            If Err.Number = 0 Then
                Exit For
            End If
        Next
        
        If Err.Number = 0 Then
            For Each WS In WB.Worksheets
                If WS.visible = xlSheetVisible Then
                    Select Case cboObj.Text
                        Case C_SEARCH_OBJECT_CELL
                            Call seachCell(WS, ResultWS)
                            
                        Case C_SEARCH_OBJECT_SHAPE
                            Call searchShape(WS, ResultWS)
                            
                        Case C_SEARCH_OBJECT_CELL_AND_SHAPE
                            Call seachCell(WS, ResultWS)
                            Call searchShape(WS, ResultWS)
                    End Select
                    
                    '１シート目のみ検索
                    If chkFirstSheet.Value Then
                        Exit For
                    End If
                End If
                Set WS = Nothing
            Next
        Else
            ResultWS.Cells(mlngCount, C_SEARCH_NO).Value = mlngCount - C_START_ROW + 1
            ResultWS.Cells(mlngCount, C_SEARCH_BOOK).Value = varBook
            ResultWS.Cells(mlngCount, C_SEARCH_SHEET).Value = "ブックを開けませんでした"
            ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Value = ""
    
            ResultWS.Cells(mlngCount, C_SEARCH_STR).NumberFormatLocal = "@"
            ResultWS.Cells(mlngCount, C_SEARCH_STR).Value = Err.Description
            mlngCount = mlngCount + 1
        End If
        WB.Close SaveChanges:=False
        Set WB = Nothing
        lngBookCount = lngBookCount + 1
        mMm.DisplayGauge lngBookCount
        
        DoEvents
    Next
    
    XL.EnableEvents = True
    XL.DisplayAlerts = True
    XL.Quit
    Set XL = Nothing
    
    Dim r As Range
    Set r = ResultWS.Cells(C_START_ROW, 1).CurrentRegion
    
    r.VerticalAlignment = xlTop
    r.Select
    
    Dim strBuf As String
    Dim i As Long
    Dim lngCount As Long
    
    strBuf = cboPattern.Text
    lngCount = 1
    For i = 0 To cboPattern.ListCount - 1
        If cboPattern.List(i) <> cboPattern.Text Then
            strBuf = strBuf & vbTab & cboPattern.List(i)
            lngCount = lngCount + 1
            'リストは最大１０
            If lngCount >= 10 Then
                Exit For
            End If
        End If
    Next
    SaveSetting C_TITLE, "ExcelGrep", "PatternStr", strBuf
    
    strBuf = cboSearch.Text
    lngCount = 1
    For i = 0 To cboSearch.ListCount - 1
        If cboSearch.List(i) <> cboSearch.Text Then
            strBuf = strBuf & vbTab & cboSearch.List(i)
            lngCount = lngCount + 1
            'リストは最大１０
            If lngCount >= 10 Then
                Exit For
            End If
        End If
    Next
    SaveSetting C_TITLE, "ExcelGrep", "SearchStr", strBuf
    
    strBuf = cboFolder.Text
    lngCount = 1
    For i = 0 To cboFolder.ListCount - 1
        If cboFolder.List(i) <> cboFolder.Text Then
            strBuf = strBuf & vbTab & cboFolder.List(i)
            lngCount = lngCount + 1
            'リストは最大１０
            If lngCount >= 10 Then
                Exit For
            End If
        End If
    Next
    SaveSetting C_TITLE, "ExcelGrep", "FolderStr", strBuf
    
    SaveSetting C_TITLE, "ExcelGrep", "cboObj", cboObj.ListIndex
    SaveSetting C_TITLE, "ExcelGrep", "chkRegEx", chkRegEx.Value
    SaveSetting C_TITLE, "ExcelGrep", "chkCase", chkCase.Value
    SaveSetting C_TITLE, "ExcelGrep", "chkSubFolder", chkSubFolder.Value
    SaveSetting C_TITLE, "ExcelGrep", "cboValue", cboValue.ListIndex
    SaveSetting C_TITLE, "ExcelGrep", "chkZenHan", chkZenHan.Value
    SaveSetting C_TITLE, "ExcelGrep", "Password", txtPassword.Text
    SaveSetting C_TITLE, "ExcelGrep", "chkFirstSheet", chkFirstSheet.Value
    
    
    SaveSetting C_TITLE, "ExcelGrep", "chkOffset", chkOffset.Value
    
    SaveSetting C_TITLE, "ExcelGrep", "txtRow", txtRow.Value
    SaveSetting C_TITLE, "ExcelGrep", "txtCol", txtCol.Value
    
    Set mMm = Nothing
    
    Unload Me
    
    Application.visible = True
    
    AppActivate ResultWS.Application.Caption
    execSelectionRowDrawGrid

    Set ResultWS = Nothing
    
    If mlngCount - C_START_ROW = 0 Then
        MsgBox "検索対象が見つかりませんでした。", vbInformation + vbOKOnly, C_TITLE
    End If
    
End Sub
Private Sub FileSearch(objFs As Object, strPath As String, strPatterns() As String, objCol As Collection)

    Dim objfld As Object
    Dim objfl As Object
    Dim objSub As Object
    Dim f As Variant
    
    Dim lngCol2 As Long

    Set objfld = objFs.GetFolder(strPath)
    
    'ファイル名取得
    For Each objfl In objfld.files
    
        Dim blnFind As Boolean
        blnFind = False
        DoEvents
        DoEvents
        DoEvents
        For Each f In strPatterns
            If LCase(objfl.Name) Like LCase(f) And Left$(objfl.Name, 2) <> "~$" Then
                blnFind = True
                Exit For
            End If
        Next
        
        If blnFind Then
            objCol.Add rlxAddFileSeparator(objfl.ParentFolder.Path) & objfl.Name
        End If
    Next
    
    'サブフォルダ検索あり
    If chkSubFolder.Value Then
        For Each objSub In objfld.SubFolders
            DoEvents
            DoEvents
            DoEvents
            FileSearch objFs, objSub.Path, strPatterns(), objCol
        Next
    End If
End Sub

'Private Sub seachCell(ByRef objSheet As Worksheet, ByRef ResultWS As Worksheet)
'
'    Dim objRegx As Object
'    Dim matchCount As Long
'    Dim objMatch As Object
'    Dim strPattern As String
'    Dim c As Range
'
'    strPattern = cboSearch.Text
'
'    '正規表現の場合
'    If chkRegEx Then
'        Set mobjRegx = CreateObject("VBScript.RegExp")
'        mobjRegx.Pattern = strPattern
'        mobjRegx.IgnoreCase = Not (chkCase.Value)
'        mobjRegx.Global = True
'    End If
'
'    For Each c In objSheet.UsedRange
'
'        '正規表現の場合
'        If chkRegEx Then
'            err.Clear
'            On Error Resume Next
'            Set objMatch = mobjRegx.Execute(c.Value)
'            If err.Number <> 0 Then
'                MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
'                cboSearch.SetFocus
'                Exit Sub
'            End If
'            matchCount = objMatch.count
'        Else
'            matchCount = InStr(c.Value, strPattern)
'        End If
'
'        If matchCount > 0 Then
'            ResultWS.Cells(mlngCount, C_SEARCH_NO).Value = mlngCount - C_START_ROW + 1
'            ResultWS.Cells(mlngCount, C_SEARCH_BOOK).Value = objSheet.Parent.FullName
'            ResultWS.Cells(mlngCount, C_SEARCH_SHEET).Value = objSheet.Name
'            ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Value = c.Address
'            ResultWS.Cells(mlngCount, C_SEARCH_STR).NumberFormatLocal = "@"
'            ResultWS.Cells(mlngCount, C_SEARCH_STR).Value = c.Value
'            mlngCount = mlngCount + 1
'        End If
'    Next
'
'
'End Sub
'Private Sub seachCell(ByRef objSheet As Worksheet, ByRef ResultWS As Worksheet)
'
'    Dim objRegx As Object
'    Dim strPattern As String
'    Dim c As Range
'
'    Dim d As Variant
'    Dim i As Long
'    Dim j As Long
'
'    strPattern = cboSearch.Text
'
'    '正規表現の場合
'    If chkRegEx Then
'        Set mobjRegx = CreateObject("VBScript.RegExp")
'        mobjRegx.Pattern = strPattern
'        mobjRegx.IgnoreCase = Not (chkCase.Value)
'        mobjRegx.Global = True
'    End If
'
'    d = objSheet.UsedRange
'    If IsEmpty(d) Then
'        Exit Sub
'    End If
'
'    If IsArray(d) Then
'        For i = LBound(d, 1) To UBound(d, 1)
'            For j = LBound(d, 2) To UBound(d, 2)
'
'                Call searchStr(objSheet, ResultWS, d(i, j), strPattern, i, j)
'
'            Next
'        Next
'    Else
'        Call searchStr(objSheet, ResultWS, d, strPattern, 1, 1)
'    End If
'
'    Erase d
'
'End Sub
'Private Sub searchStr(ByRef objSheet As Worksheet, ByRef ResultWS As Worksheet, ByVal strSearch As Variant, ByVal strPattern As String, ByVal i As Long, ByVal j As Long)
'
'    Dim objMatch As Object
'    Dim matchCount As Long
'
'    If IsError(strSearch) Then
'        Exit Sub
'    End If
'
'    '正規表現の場合
'    If chkRegEx Then
'        err.Clear
'        On Error Resume Next
'        Set objMatch = mobjRegx.Execute(strSearch)
'        If err.Number <> 0 Then
'            MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
'            cboSearch.SetFocus
'            Exit Sub
'        End If
'        matchCount = objMatch.count
'    Else
'        matchCount = InStr(strSearch, strPattern)
'    End If
'
'    If matchCount > 0 Then
'        ResultWS.Cells(mlngCount, C_SEARCH_NO).Value = mlngCount - C_START_ROW + 1
'        ResultWS.Cells(mlngCount, C_SEARCH_BOOK).Value = objSheet.Parent.FullName
'        ResultWS.Cells(mlngCount, C_SEARCH_SHEET).Value = objSheet.Name
'        ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Value = objSheet.UsedRange(i, j).Address
'        ResultWS.Hyperlinks.Add _
'            Anchor:=Cells(mlngCount, C_SEARCH_ADDRESS), _
'            Address:="", _
'            SubAddress:=Cells(mlngCount, C_SEARCH_ADDRESS).Address, _
'            TextToDisplay:=objSheet.UsedRange(i, j).Address
'
'        ResultWS.Cells(mlngCount, C_SEARCH_STR).NumberFormatLocal = "@"
'        ResultWS.Cells(mlngCount, C_SEARCH_STR).Value = strSearch
'        mlngCount = mlngCount + 1
'    End If
'End Sub
Private Sub seachCell(ByRef objSheet As Worksheet, ByRef ResultWS As Worksheet)

    Dim strPattern As String
    Dim objFind As Range
    Dim strFirstAddress As String
    Dim r As Range
    
    Dim lngRow As Long
    Dim lngCol As Long
    
'    On Error GoTo 0
    
    lngRow = Val(txtRow.Value)
    lngCol = Val(txtCol.Value)
    
    strPattern = cboSearch.Text
        
    '正規表現の場合
    If chkRegEx Then
    
        Dim objRegx As Object
        Set objRegx = CreateObject("VBScript.RegExp")
        
        objRegx.Pattern = strPattern
        objRegx.IgnoreCase = Not (chkCase.Value)
        objRegx.Global = True
    
        If cboValue.Value = C_SEARCH_VALUE_VALUE Then
            Set objFind = objSheet.UsedRange.Find("*", , xlValues, xlPart, xlByRows, xlNext, chkCase.Value, chkZenHan.Value)
        Else
            Set objFind = objSheet.UsedRange.Find("*", , xlFormulas, xlPart, xlByRows, xlNext, chkCase.Value, chkZenHan.Value)
        End If
        
        If Not objFind Is Nothing Then
        
            strFirstAddress = objFind.Address
    
            Do
    
                Dim schStr As Variant
                
                If cboValue.Value = C_SEARCH_VALUE_VALUE Then
                    schStr = objFind.Value
                Else
                    schStr = objFind.FormulaLocal
                End If
                
                If IsError(schStr) Then
                    GoTo pass
                End If
                
                Dim objMatch As Object
                Set objMatch = objRegx.Execute(schStr)
    
                If objMatch.count > 0 Then
                    ResultWS.Cells(mlngCount, C_SEARCH_NO).Value = mlngCount - C_START_ROW + 1
                    ResultWS.Cells(mlngCount, C_SEARCH_BOOK).Value = objSheet.Parent.FullName
                    ResultWS.Cells(mlngCount, C_SEARCH_SHEET).Value = objSheet.Name
                    
                    
                    If chkOffset.Value Then
                    
                        Set r = GetOffset(objFind, lngRow, lngCol)
                        
                        If r Is Nothing Then
                            GoTo pass
                        End If
                        
                        If cboValue.Value = C_SEARCH_VALUE_VALUE Then
                            schStr = r.Value
                        Else
                            schStr = r.FormulaLocal
                        End If
                        
                    Else
                        Set r = objFind
                        
                    End If
                    
                    ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Value = r.Address
        
'                    ResultWS.Hyperlinks.Add _
'                        Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
'                        Address:="", _
'                        SubAddress:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Address(False, False), _
'                        TextToDisplay:=r.Address(False, False)
                    ResultWS.Hyperlinks.Add _
                        Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
                        Address:="", _
                        TextToDisplay:=r.Address
        
                    ResultWS.Cells(mlngCount, C_SEARCH_STR).NumberFormatLocal = "@"
                    ResultWS.Cells(mlngCount, C_SEARCH_STR).Value = schStr
                    
                    mlngCount = mlngCount + 1
                End If
                Set objMatch = Nothing
pass:
                Set objFind = objSheet.UsedRange.FindNext(objFind)
                
                If objFind Is Nothing Then
                    Exit Do
                End If
        
            Loop Until strFirstAddress = objFind.Address
            Set objRegx = Nothing
            
        End If
        
    Else
        
        If cboValue.Value = C_SEARCH_VALUE_VALUE Then
            Set objFind = objSheet.UsedRange.Find(strPattern, , xlValues, xlPart, xlByColumns, xlNext, chkCase.Value, chkZenHan.Value)
        Else
            Set objFind = objSheet.UsedRange.Find(strPattern, , xlFormulas, xlPart, xlByColumns, xlNext, chkCase.Value, chkZenHan.Value)
        End If
        
        If Not objFind Is Nothing Then
        
            strFirstAddress = objFind.Address
    
            Do
            
                If chkOffset.Value Then
                    Set r = GetOffset(objFind, lngRow, lngCol)
                    If r Is Nothing Then
                        GoTo pass
                    End If
                Else
                    Set r = objFind
                End If
            
                ResultWS.Cells(mlngCount, C_SEARCH_NO).Value = mlngCount - C_START_ROW + 1
                ResultWS.Cells(mlngCount, C_SEARCH_BOOK).Value = objSheet.Parent.FullName
                ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Value = r.Address
                ResultWS.Cells(mlngCount, C_SEARCH_SHEET).Value = objSheet.Name
                
'                ResultWS.Hyperlinks.Add _
'                    Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
'                    Address:="", _
'                    SubAddress:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Address(False, False), _
'                    TextToDisplay:=r.Address(False, False)
                ResultWS.Hyperlinks.Add _
                    Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
                    Address:="", _
                    SubAddress:="", _
                    TextToDisplay:=r.Address
        
        
                ResultWS.Cells(mlngCount, C_SEARCH_STR).NumberFormatLocal = "@"
                
                If cboValue.Value = C_SEARCH_VALUE_VALUE Then
                    ResultWS.Cells(mlngCount, C_SEARCH_STR).Value = r.Value
                Else
                    ResultWS.Cells(mlngCount, C_SEARCH_STR).Value = r.FormulaLocal
                End If

                mlngCount = mlngCount + 1
        
                Set objFind = objSheet.UsedRange.FindNext(objFind)
                
                If objFind Is Nothing Then
                    Exit Do
                End If
                
            Loop Until strFirstAddress = objFind.Address
            
        End If
    End If
    
End Sub
Function GetOffset(r As Range, ByVal lngRow As Long, ByVal lngCol As Long) As Range
    
    Set GetOffset = Nothing

    On Error GoTo e
    
    Set GetOffset = r.Offset(lngRow, lngCol)
e:

End Function
Private Sub searchShape(ByRef objSheet As Worksheet, ByRef ResultWS As Worksheet)

    Dim matchCount As Long
    Dim objMatch As Object
    Dim strPattern As String

    Dim objShape As Shape
    Dim objAct As Worksheet
    Dim c As Shape
    
    Dim strBuf As String

    Dim colShapes As Collection
    Set colShapes = New Collection

    Const C_RESULT_NAME As String = "シェイプ検索Result"
    
    strPattern = cboSearch.Text
    
    '正規表現の場合
    If chkRegEx Then
        Set mobjRegx = CreateObject("VBScript.RegExp")
        mobjRegx.Pattern = strPattern
        mobjRegx.IgnoreCase = Not (chkCase.Value)
        mobjRegx.Global = True
    End If
    
    For Each c In objSheet.Shapes
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                'シェイプに文字があるかないか判断がつかないためエラー検出にて処理
                On Error Resume Next
                strBuf = c.TextFrame.Characters.Text
                If Err.Number = 0 Then
                    On Error GoTo 0
                    
                    '正規表現の場合
                    If chkRegEx Then
                        Err.Clear
                        On Error Resume Next
                        Set objMatch = mobjRegx.Execute(strBuf)
                        If Err.Number <> 0 Then
                            MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
                            cboSearch.SetFocus
                            Exit Sub
                        End If
                        matchCount = objMatch.count
                    Else
'                        matchCount = InStr(strBuf, strPattern)
                        If chkCase.Value Then
                            matchCount = InStr(strBuf, strPattern)
                        Else
                            matchCount = InStr(UCase(strBuf), UCase(strPattern))
                        End If
                    End If
                    
                    If matchCount > 0 Then
                    
                        ResultWS.Cells(mlngCount, C_SEARCH_NO).Value = mlngCount - C_START_ROW + 1
                        ResultWS.Cells(mlngCount, C_SEARCH_BOOK).Value = objSheet.Parent.FullName
                        ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Value = c.Name & ":" & c.id
'                        ResultWS.Cells(mlngCount, C_SEARCH_ID).Value = "Shape:" & c.ID
                        
'                        ResultWS.Hyperlinks.Add _
'                            Anchor:=Cells(mlngCount, C_SEARCH_ADDRESS), _
'                            Address:=objSheet.Parent.FullName, _
'                            SubAddress:="'" & objSheet.Name & "'!" & c.TopLeftCell.Address(0, 0), _
'                            TextToDisplay:=c.Name
'                ResultWS.Hyperlinks.Add _
'                    Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
'                    Address:="", _
'                    SubAddress:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Address, _
'                    TextToDisplay:=c.Name
'                ResultWS.Hyperlinks.Add _
'                    Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
'                    Address:="", _
'                    SubAddress:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Address, _
'                    TextToDisplay:=c.Name & ":" & c.id
                ResultWS.Hyperlinks.Add _
                    Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
                    Address:="", _
                    TextToDisplay:=c.Name & ":" & c.id
                    
                        ResultWS.Cells(mlngCount, C_SEARCH_SHEET).Value = objSheet.Name
                        ResultWS.Cells(mlngCount, C_SEARCH_STR).NumberFormatLocal = "@"
                        ResultWS.Cells(mlngCount, C_SEARCH_STR).Value = strBuf
                        mlngCount = mlngCount + 1
                        
                    End If
                Else
                    On Error GoTo 0
                    Err.Clear
                End If
            Case msoGroup
                grouprc c, c, colShapes, ResultWS

        End Select
    Next

End Sub
'再帰にてグループ以下のシェイプを検索
Private Sub grouprc(ByRef objTop As Shape, ByRef objShape As Shape, ByRef colShapes As Collection, ByRef ResultWS As Worksheet)

    Dim matchCount As Long
    Dim c As Shape
    Dim strBuf As String
    Dim objMatch As Object
    Dim strPattern As String
    strPattern = cboSearch.Text
    
    For Each c In objShape.GroupItems
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                'シェイプに文字があるかないか判断がつかないためエラー検出にて処理
                On Error Resume Next
                strBuf = c.TextFrame.Characters.Text
                If Err.Number = 0 Then
                    On Error GoTo 0
                    
                    '正規表現の場合
                    If chkRegEx Then
                        Err.Clear
                        On Error Resume Next
                        Set objMatch = mobjRegx.Execute(strBuf)
                        If Err.Number <> 0 Then
                            MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
                            cboSearch.SetFocus
                            Exit Sub
                        End If
                        matchCount = objMatch.count
                    Else
                        matchCount = InStr(strBuf, strPattern)
                    End If
                    
                    If matchCount > 0 Then
                    
                        ResultWS.Cells(mlngCount, C_SEARCH_NO).Value = mlngCount - C_START_ROW + 1
                        ResultWS.Cells(mlngCount, C_SEARCH_BOOK).Value = objShape.Parent.Parent.FullName
                        ResultWS.Cells(mlngCount, C_SEARCH_SHEET).Value = objShape.Parent.Name
                        ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS).Value = c.Name & ":" & c.id
'                        ResultWS.Cells(mlngCount, C_SEARCH_ID).Value = "Shape:" & c.ID
                        
'                        ResultWS.Hyperlinks.Add _
'                            Anchor:=Cells(mlngCount, C_SEARCH_ADDRESS), _
'                            Address:=objShape.Parent.Parent.FullName, _
'                            SubAddress:="'" & objShape.Parent.Name & "'!" & c.TopLeftCell.Address(0, 0), _
'                            TextToDisplay:=c.Name
'                ResultWS.Hyperlinks.Add _
'                    Anchor:=Cells(mlngCount, C_SEARCH_ADDRESS), _
'                    Address:="", _
'                    SubAddress:=Cells(mlngCount, C_SEARCH_ADDRESS).Address, _
'                    TextToDisplay:=c.Name
'                ResultWS.Hyperlinks.Add _
'                    Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
'                    Address:="", _
'                    SubAddress:=Cells(mlngCount, C_SEARCH_ADDRESS).Address, _
'                    TextToDisplay:=c.Name & ":" & c.id
                ResultWS.Hyperlinks.Add _
                    Anchor:=ResultWS.Cells(mlngCount, C_SEARCH_ADDRESS), _
                    Address:="", _
                    TextToDisplay:=c.Name & ":" & c.id
                        
                        ResultWS.Cells(mlngCount, C_SEARCH_STR).NumberFormatLocal = "@"
                        ResultWS.Cells(mlngCount, C_SEARCH_STR).Value = strBuf
                        mlngCount = mlngCount + 1
                    
                    End If
                Else
                    On Error GoTo 0
                    Err.Clear
                End If
            Case msoGroup
                '再帰呼出
                grouprc objTop, c, colShapes, ResultWS
            
        End Select
    Next

End Sub




Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Activate()
    AppActivate Me.Caption
End Sub

Private Sub UserForm_Initialize()
    
    Dim strBuf As String
    Dim strSearch() As String
    Dim strFolder() As String
    Dim i As Long
    
    mblnRefresh = True
    
    cboObj.AddItem C_SEARCH_OBJECT_CELL
    cboObj.AddItem C_SEARCH_OBJECT_SHAPE
    cboObj.AddItem C_SEARCH_OBJECT_CELL_AND_SHAPE
    cboObj.ListIndex = GetSetting(C_TITLE, "ExcelGrep", "cboObj", 0)
    
    cboValue.AddItem C_SEARCH_VALUE_FORMULA
    cboValue.AddItem C_SEARCH_VALUE_VALUE
    cboValue.ListIndex = GetSetting(C_TITLE, "ExcelGrep", "cboValue", 0)
    
    chkSubFolder.Value = GetSetting(C_TITLE, "ExcelGrep", "chkSubFolder", False)
    
    chkRegEx.Value = GetSetting(C_TITLE, "ExcelGrep", "chkRegEx", False)
    chkCase.Value = GetSetting(C_TITLE, "ExcelGrep", "chkCase", False)
    chkZenHan.Value = GetSetting(C_TITLE, "ExcelGrep", "chkZenHan", False)
    
    chkFirstSheet.Value = GetSetting(C_TITLE, "ExcelGrep", "chkFirstSheet", False)
    chkOffset.Value = GetSetting(C_TITLE, "ExcelGrep", "chkOffset", False)
    
    txtRow.enabled = chkOffset.Value
    txtCol.enabled = chkOffset.Value
    
    txtRow.Value = GetSetting(C_TITLE, "ExcelGrep", "txtRow", 0)
    txtCol.Value = GetSetting(C_TITLE, "ExcelGrep", "txtCol", 0)
    
    Dim strdef As String
    strdef = "*.xlsx;*.xlsm;*.xls"
    
    cboPattern.AddItem strdef
    strBuf = GetSetting(C_TITLE, "ExcelGrep", "PatternStr", strdef)
    strSearch = Split(strBuf, vbTab)
    For i = LBound(strSearch) To UBound(strSearch)
        If strSearch(i) <> strdef Then
            cboPattern.AddItem strSearch(i)
        End If
    Next
    If cboPattern.ListCount > 0 Then
        cboPattern.ListIndex = 0
    End If
    
    
    strBuf = GetSetting(C_TITLE, "ExcelGrep", "SearchStr", "")
    strSearch = Split(strBuf, vbTab)
    
    txtPassword.Text = GetSetting(C_TITLE, "ExcelGrep", "Password", "")
    
    For i = LBound(strSearch) To UBound(strSearch)
        cboSearch.AddItem strSearch(i)
    Next
    If cboSearch.ListCount > 0 Then
        cboSearch.ListIndex = 0
    End If
    
    strBuf = GetSetting(C_TITLE, "ExcelGrep", "FolderStr", "")
    strFolder = Split(strBuf, vbTab)
    
    For i = LBound(strFolder) To UBound(strFolder)
        cboFolder.AddItem strFolder(i)
    Next
    If cboFolder.ListCount > 0 Then
        cboFolder.ListIndex = 0
    End If

    lblGauge.visible = False

   ' txtBack.Value = "ExcelブックのGrepを行います"

'    Me.Top = (Application.Top + Application.Height - Me.Height) - 20
'    Me.Left = (Application.Left + Application.Width - Me.Width) - 20

    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    'コントロールボックスからClose
    If CloseMode = vbFormControlMenu Then
        If cmdCancel.Caption = "閉じる" Then
            Unload Me
            If ThisWorkbook.IsAddin Then
                Application.Quit
            End If
        Else
            mblnCancel = True
        End If
    End If
    
End Sub
