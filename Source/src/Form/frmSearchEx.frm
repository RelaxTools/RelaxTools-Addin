VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchEx 
   Caption         =   "セル・シェイプの正規表現検索／置換"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   OleObjectBlob   =   "frmSearchEx.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmSearchEx"
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

Private mblnRefresh As Boolean

Private Const C_SEARCH_NO As Long = 0
Private Const C_SEARCH_STR As Long = 1
Private Const C_SEARCH_ADDRESS As Long = 2
Private Const C_SEARCH_SHEET As Long = 3
Private Const C_SEARCH_ID As Long = 4
Private Const C_SEARCH_BOOK As Long = 5

Private Const C_SEARCH_PLACE_SHEET = "現在のシート"
Private Const C_SEARCH_PLACE_SELECT = "選択したシート"
Private Const C_SEARCH_PLACE_BOOK = "ブック全体"
Private Const C_SEARCH_OBJECT_CELL = "セルのみ"
Private Const C_SEARCH_OBJECT_SHAPE = "シェイプのみ"
Private Const C_SEARCH_OBJECT_CELL_AND_SHAPE = "セル＆シェイプ"
Private Const C_SEARCH_VALUE_VALUE = "値"
Private Const C_SEARCH_VALUE_FORMULA = "式"
Private Const C_SEARCH_ID_CELL As String = "Cell:"
Private Const C_SEARCH_ID_SHAPE As String = "Shape"
Private Const C_SEARCH_ID_SMARTART As String = "SmartArt"

Private Const C_SIZE As Long = 256

Private RW As ResizeWindow
Private mlngListWidth As Long
Private mlngListHeight As Long
Private mlngTabWidth As Long
Private mlngTabHeight As Long
Private mlngLblWidth As Long
Private mlngLblObjLeft As Long
Private mlngLblPlaceLeft As Long
Private mlngColumnWidth As Long
Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1
'Private TR As Transparent

'Private Tr As Transparent

Private Sub chkRegEx_Change()

'    chkZenHan.enabled = Not (chkRegEx.Value)

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
        
    If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    Call GotoRegExpHelp

End Sub

Private Sub cmdOk_Click()

    If Len(Trim(txtSearch.Text)) = 0 Then
        MsgBox "検索文字列を指定してください。", vbExclamation, C_TITLE
        txtSearch.SetFocus
        Exit Sub
    End If
    
    '正規表現の場合
    If chkRegEx.Value Then
        Dim o As Object
        Set o = CreateObject("VBScript.RegExp")
        o.Pattern = txtSearch.Text
        o.IgnoreCase = Not (chkCase.Value)
        o.Global = True
        Err.Clear
        On Error Resume Next
        o.Execute ""
        If Err.Number <> 0 Then
            MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
            txtSearch.SetFocus
            Exit Sub
        End If
    End If
    
    Logger.LogBegin TypeName(Me) & ".cmdOk_Click"
    
    Static sw
    
    If sw Then
        Exit Sub
    End If
    
    sw = True
    
    Call searchStart
    
    Dim strBuf As String
    Dim i As Long
    Dim lngCount As Long
    Dim strSearch() As String
    
    strBuf = txtSearch.Text
    lngCount = 1
    For i = 0 To txtSearch.ListCount - 1
        If txtSearch.List(i) <> txtSearch.Text Then
            strBuf = strBuf & vbTab & txtSearch.List(i)
            lngCount = lngCount + 1
            'リストは最大１０
            If lngCount >= 10 Then
                Exit For
            End If
        End If
    Next
    SaveSetting C_TITLE, "Search", "SearchStr", strBuf

    txtSearch.Clear
    strSearch = Split(strBuf, vbTab)
    
    For i = LBound(strSearch) To UBound(strSearch)
        txtSearch.AddItem strSearch(i)
    Next
    If txtSearch.ListCount > 0 Then
        txtSearch.ListIndex = 0
    End If
    
    SaveSetting C_TITLE, "Search", "cboPlace", cboPlace.ListIndex
    SaveSetting C_TITLE, "Search", "cboObj", cboObj.ListIndex
    SaveSetting C_TITLE, "Search", "chkRegEx", chkRegEx.Value
    SaveSetting C_TITLE, "Search", "chkCase", chkCase.Value
    SaveSetting C_TITLE, "Search", "chkZenHan", chkZenHan.Value
    SaveSetting C_TITLE, "Search", "cboValue", cboValue.ListIndex
    SaveSetting C_TITLE, "Search", "chkSmartArt", chkSmartArt.Value
    
    Logger.LogFinish TypeName(Me) & ".cmdOk_Click"
    
    If lstResult.ListCount = 0 Then
        MsgBox "検索対象が見つかりませんでした。", vbInformation + vbOKOnly, C_TITLE
    End If
    
    sw = False

End Sub


Private Sub cmdReplace_Click()

    Dim lngRet As Long

    If Len(Trim(txtSearch.Text)) = 0 Then
        MsgBox "検索文字列を指定してください。", vbExclamation, C_TITLE
        txtSearch.SetFocus
        Exit Sub
    End If
    
    '正規表現の場合
    If chkRegEx.Value Then
        Dim o As Object
        Set o = CreateObject("VBScript.RegExp")
        o.Pattern = txtSearch.Text
        o.IgnoreCase = Not (chkCase.Value)
        o.Global = True
        Err.Clear
        On Error Resume Next
        o.Execute ""
        If Err.Number <> 0 Then
            MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
            txtSearch.SetFocus
            Exit Sub
        End If
    End If
    
    Call searchStart
    
    If lstResult.ListCount = 0 Then
        MsgBox "検索対象が見つかりませんでした。", vbInformation + vbOKOnly, C_TITLE
        Exit Sub
    End If
    
    lngRet = replaceStart(True)
    If lngRet < 0 Then
        Exit Sub
    End If
        
    MsgBox lngRet & " 個の文字列を置換しました。", vbInformation + vbOKOnly, C_TITLE
    
    Dim strBuf As String
    Dim i As Long
    Dim lngCount As Long
    Dim strSearch() As String
    Dim strReplace() As String
    
    strBuf = txtSearch.Text
    lngCount = 1
    For i = 0 To txtSearch.ListCount - 1
        If txtSearch.List(i) <> txtSearch.Text Then
            strBuf = strBuf & vbTab & txtSearch.List(i)
            lngCount = lngCount + 1
            'リストは最大１０
            If lngCount >= 10 Then
                Exit For
            End If
        End If
    Next
    SaveSetting C_TITLE, "Search", "SearchStr", strBuf
    
    txtSearch.Clear
    strSearch = Split(strBuf, vbTab)
    
    For i = LBound(strSearch) To UBound(strSearch)
        txtSearch.AddItem strSearch(i)
    Next
    If txtSearch.ListCount > 0 Then
        txtSearch.ListIndex = 0
    End If
    
    strBuf = txtReplace.Text
    lngCount = 1
    For i = 0 To txtReplace.ListCount - 1
        If txtReplace.List(i) <> txtReplace.Text Then
            strBuf = strBuf & vbTab & txtReplace.List(i)
            lngCount = lngCount + 1
            'リストは最大１０
            If lngCount >= 10 Then
                Exit For
            End If
        End If
    Next
    SaveSetting C_TITLE, "Search", "ReplaceStr", strBuf

    txtReplace.Clear
    strReplace = Split(strBuf, vbTab)
    
    For i = LBound(strReplace) To UBound(strReplace)
        txtReplace.AddItem strReplace(i)
    Next
    If txtReplace.ListCount > 0 Then
        txtReplace.ListIndex = 0
    End If
    
End Sub

Private Function replaceStart(ByVal blnAsk As Boolean) As Long

    Dim objMatch As Object
    Dim matchCount As Long
    Dim i As Long
    Dim j As Long
    Dim ret As Long
    Dim lngRet As Long
    Dim strAddress As String
    Dim strSheet As String
    Dim objRegx As Object
    Dim strBuf As String
    
    Dim strPattern As String
    Dim strReplace As String
    
    strPattern = txtSearch.Text
    strReplace = txtReplace.Text
    
    If lstResult.ListCount = 0 Then
        Exit Function
    End If
    
    For i = 0 To lstResult.ListCount - 1
    
        strAddress = lstResult.List(i, C_SEARCH_ID)
        strSheet = lstResult.List(i, C_SEARCH_SHEET)
        If blnAsk Then
        
            mblnRefresh = False
            For j = 0 To lstResult.ListCount - 1
                If j = i Then
                Else
                    lstResult.Selected(j) = False
                End If
            Next
            mblnRefresh = True
            lstResult.Selected(i) = True
            lstResult.TopIndex = i
            
            DoEvents
            
            ret = frmSearchBox.Start()
            Select Case ret
                Case 2
                    blnAsk = False
                Case 4
                    GoTo pass
                Case 8
                    lngRet = -1
                    Exit For
            End Select
        End If
        
        If InStr(strAddress, "$") > 0 Then
        
            Dim r As Range
            Set r = Worksheets(strSheet).Range(strAddress)
           
            '正規表現の場合
            If chkRegEx.Value Then
            
                Set objRegx = CreateObject("VBScript.RegExp")
                
                objRegx.Pattern = strPattern
                objRegx.IgnoreCase = Not (chkCase.Value)
                objRegx.Global = True
                
                If r.HasFormula Then
                    r.FormulaLocal = objRegx.Replace(r.FormulaLocal, convEscSeq(strReplace))
                Else
                    r.Value = objRegx.Replace(r.Value, convEscSeq(strReplace))
                End If
                Set objRegx = Nothing
               
            Else
                Call r.Replace(strPattern, strReplace, xlPart, xlByColumns, chkCase.Value, False)
            End If
            
            If r.HasFormula Then
                lstResult.List(i, C_SEARCH_STR) = r.FormulaLocal
            Else
                lstResult.List(i, C_SEARCH_STR) = Left(r.Value, C_SIZE)
            End If
            Set r = Nothing
        Else
        
            Dim s As Object

            Set s = getObjFromID(Worksheets(strSheet), Mid$(strAddress, InStrRev(strAddress, ":") + 1))
           
            '正規表現の場合
            If chkRegEx.Value Then
              
                Set objRegx = CreateObject("VBScript.RegExp")
                
                objRegx.Pattern = strPattern
                objRegx.IgnoreCase = Not (chkCase.Value)
                objRegx.Global = True
                
                If InStr(strAddress, C_SEARCH_ID_SMARTART) > 0 Then
                    strBuf = objRegx.Replace(s.TextFrame2.TextRange.Text, convEscSeq(strReplace))
'                    s.TextFrame2.TextRange.Text = strBuf
                Else
                    strBuf = objRegx.Replace(s.TextFrame.Characters.Text, convEscSeq(strReplace))
'                    s.TextFrame.Characters.Text = strBuf
                End If
                
                Set objRegx = Nothing
               
            Else
                Dim strL As String
                Dim strR As String
                Dim lngPos As Long
                Dim lngLen As Long
                
                If InStr(strAddress, C_SEARCH_ID_SMARTART) > 0 Then
                    strBuf = s.TextFrame2.TextRange.Text
                Else
                    strBuf = s.TextFrame.Characters.Text
                End If
                
                lngLen = Len(strPattern)
                If chkCase.Value Then
                    lngPos = InStr(strBuf, strPattern)
                Else
                    lngPos = InStr(UCase(strBuf), UCase(strPattern))
                End If
                
                Do Until lngPos = 0
                
                    strL = Mid$(strBuf, 1, lngPos - 1)
                    strR = Mid$(strBuf, lngPos + lngLen)
                    
                    strBuf = strL & strReplace & strR
                    
                    If chkCase.Value Then
                        lngPos = InStr(Len(strL & strReplace) + 1, strBuf, strPattern)
                    Else
                        lngPos = InStr(Len(strL & strReplace) + 1, UCase(strBuf), UCase(strPattern))
                    End If
                Loop
            End If
            
            lstResult.List(i, C_SEARCH_STR) = Left(strBuf, C_SIZE)
            
            If InStr(strAddress, C_SEARCH_ID_SMARTART) > 0 Then
                s.TextFrame2.TextRange.Text = strBuf
            Else
                s.TextFrame.Characters.Text = strBuf
            End If
            
            Set s = Nothing
        
        End If
        lngRet = lngRet + 1
pass:
    Next
    
    Me.Show
    replaceStart = lngRet

End Function
Private Sub searchStart()
    
    Dim colSheet As Collection
    Dim objSheet1 As Worksheet
    Dim objSheet2 As Worksheet
    
    lstResult.Clear
    mlngCount = 0

    Set colSheet = New Collection

    Select Case cboPlace.Text
        Case C_SEARCH_PLACE_SHEET
            colSheet.Add ActiveSheet
            
        Case C_SEARCH_PLACE_SELECT
            For Each objSheet1 In ActiveWorkbook.Windows(1).SelectedSheets
                If objSheet1.visible = xlSheetVisible Then
                    colSheet.Add objSheet1
                End If
            Next
            
        Case Else
            For Each objSheet1 In ActiveWorkbook.Worksheets
                If objSheet1.visible = xlSheetVisible Then
                    colSheet.Add objSheet1
                End If
            Next
    End Select

    For Each objSheet2 In colSheet

        Select Case cboObj.Text
            Case C_SEARCH_OBJECT_CELL
                Call seachCell(objSheet2)
                
            Case C_SEARCH_OBJECT_SHAPE
                Call searchShape(objSheet2)
                
            Case C_SEARCH_OBJECT_CELL_AND_SHAPE
                Call seachCell(objSheet2)
                Call searchShape(objSheet2)
                
        End Select
    
    Next

    Set colSheet = Nothing
    
End Sub

Private Sub seachCell(ByRef objSheet As Worksheet)

    Dim strPattern As String
    Dim objFind As Range
    Dim strFirstAddress As String
    
    strPattern = txtSearch.Text
        
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
                    If objFind.HasFormula Then
                        schStr = objFind.FormulaLocal
                    Else
                        schStr = objFind.Value
                    End If
                End If
                
                Dim objMatch As Object
                Set objMatch = objRegx.Execute(schStr)
    
                If objMatch.count > 0 Then
                
                    lstResult.AddItem ""
                    lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                    
                    lstResult.List(mlngCount, C_SEARCH_STR) = Left(schStr, C_SIZE)
                    
                    lstResult.List(mlngCount, C_SEARCH_ADDRESS) = objFind.Address
                    lstResult.List(mlngCount, C_SEARCH_ID) = objFind.Address
                    lstResult.List(mlngCount, C_SEARCH_SHEET) = objSheet.Name
                    lstResult.List(mlngCount, C_SEARCH_BOOK) = objSheet.Parent.Name
    
                    mlngCount = mlngCount + 1
                End If
                
                Set objMatch = Nothing
                Set objFind = objSheet.UsedRange.FindNext(objFind)
                
                DoEvents
                
                If objFind Is Nothing Then
                    Exit Do
                End If
                
            Loop Until strFirstAddress = objFind.Address
            Set objRegx = Nothing
        End If
    Else
        
        If cboValue.Value = C_SEARCH_VALUE_VALUE Then
            Set objFind = objSheet.UsedRange.Find(strPattern, , xlValues, xlPart, xlByRows, xlNext, chkCase.Value, chkZenHan.Value)
        Else
            Set objFind = objSheet.UsedRange.Find(strPattern, , xlFormulas, xlPart, xlByRows, xlNext, chkCase.Value, chkZenHan.Value)
        End If
        
        If Not objFind Is Nothing Then
        
            strFirstAddress = objFind.Address
    
            Do
            
                lstResult.AddItem ""
                lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                
                If cboValue.Value = C_SEARCH_VALUE_VALUE Then
                    lstResult.List(mlngCount, C_SEARCH_STR) = Left(objFind.Value, C_SIZE)
                Else
                    If objFind.HasFormula Then
                        lstResult.List(mlngCount, C_SEARCH_STR) = objFind.FormulaLocal
                    Else
                        lstResult.List(mlngCount, C_SEARCH_STR) = Left(objFind.Value, C_SIZE)
                    End If
                End If
                
                lstResult.List(mlngCount, C_SEARCH_ADDRESS) = objFind.Address
                lstResult.List(mlngCount, C_SEARCH_ID) = objFind.Address
                
                lstResult.List(mlngCount, C_SEARCH_SHEET) = objSheet.Name
                lstResult.List(mlngCount, C_SEARCH_BOOK) = objSheet.Parent.Name

                mlngCount = mlngCount + 1
        
                Set objFind = objSheet.UsedRange.FindNext(objFind)
                
                If objFind Is Nothing Then
                    Exit Do
                End If
                
                DoEvents
                
            Loop Until strFirstAddress = objFind.Address
            
        End If
    End If
    
End Sub
Private Sub searchShape(ByRef objSheet As Worksheet)

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
    
    strPattern = txtSearch.Text
    
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

                If c.TextFrame2.HasText Then
                    strBuf = c.TextFrame2.TextRange.Text
                    
                    '正規表現の場合
                    If chkRegEx Then
                        Err.Clear
                        On Error Resume Next
                        Set objMatch = mobjRegx.Execute(strBuf)
                        If Err.Number <> 0 Then
                            MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
                            txtSearch.SetFocus
                            Exit Sub
                        End If
                        matchCount = objMatch.count
                    Else
                        If chkCase.Value Then
                            matchCount = InStr(strBuf, strPattern)
                        Else
                            matchCount = InStr(UCase(strBuf), UCase(strPattern))
                        End If
                    End If
                    
                    If matchCount > 0 Then
                    
                        lstResult.AddItem ""
                        lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                        lstResult.List(mlngCount, C_SEARCH_STR) = Left(strBuf, C_SIZE)
                        lstResult.List(mlngCount, C_SEARCH_ADDRESS) = c.Name
                        lstResult.List(mlngCount, C_SEARCH_ID) = C_SEARCH_ID_SHAPE & ":" & c.id
                        lstResult.List(mlngCount, C_SEARCH_SHEET) = objSheet.Name
                        lstResult.List(mlngCount, C_SEARCH_BOOK) = objSheet.Parent.Name

                        mlngCount = mlngCount + 1
                        
                    End If
                Else
                    On Error GoTo 0
                    Err.Clear
                End If
            Case msoGroup
                grouprc objSheet, c, c, colShapes

            Case msoSmartArt
                'check があるときのみ検索
                If chkSmartArt.Value Then
                    SmartArtprc objSheet, c, c, colShapes
                End If

        End Select
        DoEvents
    Next

End Sub
'再帰にてグループ以下のシェイプを検索
Private Sub grouprc(ByRef WS As Worksheet, ByRef objTop As Shape, ByRef objShape As Shape, ByRef colShapes As Collection)

    Dim matchCount As Long
    Dim c As Shape
    Dim strBuf As String
    Dim objMatch As Object
    Dim strPattern As String
    strPattern = txtSearch.Text
    
    For Each c In objShape.GroupItems
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                'シェイプに文字があるかないか判断がつかないためエラー検出にて処理
                If c.TextFrame2.HasText Then
                    strBuf = c.TextFrame2.TextRange.Text
                    
                    '正規表現の場合
                    If chkRegEx Then
                        Err.Clear
                        On Error Resume Next
                        Set objMatch = mobjRegx.Execute(strBuf)
                        If Err.Number <> 0 Then
                            MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
                            txtSearch.SetFocus
                            Exit Sub
                        End If
                        matchCount = objMatch.count
                    Else
                        If chkCase.Value Then
                            matchCount = InStr(strBuf, strPattern)
                        Else
                            matchCount = InStr(UCase(strBuf), UCase(strPattern))
                        End If
                    End If
                    
                    If matchCount > 0 Then
                    
                        lstResult.AddItem ""
                        lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                        lstResult.List(mlngCount, C_SEARCH_STR) = Left(strBuf, C_SIZE)

                        lstResult.List(mlngCount, C_SEARCH_ADDRESS) = c.Name
                        lstResult.List(mlngCount, C_SEARCH_ID) = C_SEARCH_ID_SHAPE & getGroupId(c) & ":" & c.id
                        lstResult.List(mlngCount, C_SEARCH_SHEET) = WS.Name
                        lstResult.List(mlngCount, C_SEARCH_BOOK) = WS.Parent.Name

                        mlngCount = mlngCount + 1
                    
                    End If
                Else
                    On Error GoTo 0
                    Err.Clear
                End If
            Case msoGroup
                '再帰呼出
                grouprc WS, objTop, c, colShapes
            
            Case msoSmartArt
                'check があるときのみ検索
                If chkSmartArt.Value Then
                    SmartArtprc WS, c, c, colShapes
                End If
        End Select
    Next

End Sub
'スマートアートを検索
Private Sub SmartArtprc(ByRef WS As Worksheet, ByRef objTop As Shape, ByRef objShape As Shape, ByRef colShapes As Collection)

    Dim matchCount As Long
    Dim c As SmartArtNode
    Dim strBuf As String
    Dim objMatch As Object
    Dim strPattern As String
    Dim lngIdx As Long
    strPattern = txtSearch.Text
    
    
    For lngIdx = 1 To objShape.SmartArt.AllNodes.count
    
        Set c = objShape.SmartArt.AllNodes(lngIdx)
                
        'シェイプに文字があるかないか判断がつかないためエラー検出にて処理
        If c.TextFrame2.HasText Then
        
            strBuf = c.TextFrame2.TextRange.Text
            
            '正規表現の場合
            If chkRegEx Then
                Err.Clear
                On Error Resume Next
                Set objMatch = mobjRegx.Execute(strBuf)
                If Err.Number <> 0 Then
                    MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
                    txtSearch.SetFocus
                    Exit Sub
                End If
                matchCount = objMatch.count
            Else
                If chkCase.Value Then
                    matchCount = InStr(strBuf, strPattern)
                Else
                    matchCount = InStr(UCase(strBuf), UCase(strPattern))
                End If
            End If
            
            If matchCount > 0 Then
                lstResult.AddItem ""
                lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                lstResult.List(mlngCount, C_SEARCH_STR) = Left(strBuf, C_SIZE)
                lstResult.List(mlngCount, C_SEARCH_ADDRESS) = objShape.Name
                lstResult.List(mlngCount, C_SEARCH_ID) = C_SEARCH_ID_SMARTART & getGroupId(objShape) & "/" & objShape.id & ":" & objShape.id & "," & lngIdx
                lstResult.List(mlngCount, C_SEARCH_SHEET) = WS.Name
                lstResult.List(mlngCount, C_SEARCH_BOOK) = WS.Parent.Name
                mlngCount = mlngCount + 1
            End If
        Else
            On Error GoTo 0
            Err.Clear
        End If
    
    Next

End Sub
'グループ文字列を取得
Private Function getGroupId(ByRef objShape As Object) As String

    Dim strBuf As String
    Dim s As Object
    
    On Error Resume Next
    Err.Clear
    Set s = objShape.ParentGroup
    Do Until Err.Number <> 0
        strBuf = "/" & s.id & strBuf
        Set s = s.ParentGroup
    Loop
    
    getGroupId = strBuf

End Function

Private Sub lstResult_Change()

    If mblnRefresh = False Then
         Exit Sub
    End If

    Dim lngCnt As Long
    Dim strRange As String
    Dim r As Range
    Dim s As String
    
    Dim selSheet As String
    Dim selBook As String
    Dim blnCell As Boolean
'    Dim blnShape As Boolean
    Dim strPath As String
    selSheet = ""
    selBook = ""
    
    blnCell = False
    
    For lngCnt = 0 To lstResult.ListCount - 1
    
        If lstResult.Selected(lngCnt) Then
            If selSheet = "" Then
                selSheet = lstResult.List(lngCnt, C_SEARCH_SHEET)
                selBook = lstResult.List(lngCnt, C_SEARCH_BOOK)
                If Left$(lstResult.List(lngCnt, C_SEARCH_ID), 1) = "$" Then
                    blnCell = True
                Else
                    Dim p() As String
                    p = Split(lstResult.List(lngCnt, C_SEARCH_ID), ":")
'                    blnShape = True
                    strPath = p(0)
                End If
            Else
                If selSheet <> lstResult.List(lngCnt, C_SEARCH_SHEET) Then
                    mblnRefresh = False
                    lstResult.Selected(lngCnt) = False
                    mblnRefresh = True
                Else
                    If blnCell Then
                        '１行目がセルで２行目以降でセル以外
                        If Left$(lstResult.List(lngCnt, C_SEARCH_ID), 1) <> "$" Then
                            mblnRefresh = False
                            lstResult.Selected(lngCnt) = False
                            mblnRefresh = True
                        End If
                    Else
                        '１行目がシェイプ
                        If Left$(lstResult.List(lngCnt, C_SEARCH_ID), 1) = "$" Then
                            mblnRefresh = False
                            lstResult.Selected(lngCnt) = False
                            mblnRefresh = True
                        Else
                            p = Split(lstResult.List(lngCnt, C_SEARCH_ID), ":")
                            If strPath <> p(0) Then
                                mblnRefresh = False
                                lstResult.Selected(lngCnt) = False
                                mblnRefresh = True
                            End If
                        End If
                    
                    End If
                    
                    
                End If
            End If
        
        End If
    Next
    
    If Len(selSheet) = 0 Then
        Exit Sub
    End If
    
    Workbooks(selBook).Activate
    Worksheets(selSheet).Select
    
    If blnCell Then
        For lngCnt = 0 To lstResult.ListCount - 1
    
            If lstResult.Selected(lngCnt) Then
                If r Is Nothing Then
                    Set r = Range("'[" & lstResult.List(lngCnt, C_SEARCH_BOOK) & "]" & lstResult.List(lngCnt, C_SEARCH_SHEET) & "'!" & lstResult.List(lngCnt, C_SEARCH_ID))
                Else
                    Set r = Union(r, Range("'[" & lstResult.List(lngCnt, C_SEARCH_BOOK) & "]" & lstResult.List(lngCnt, C_SEARCH_SHEET) & "'!" & lstResult.List(lngCnt, C_SEARCH_ID)))
                End If
            End If
        Next
        If r Is Nothing Then
        Else
            Application.GoTo setCellPos(r(1)), True
            r.Select
        End If
    Else
    
        Dim strBuf As String
        Dim strId As String
        Dim objShape As Object
        Dim objArt As Object
        Dim blnFlg As Boolean
        blnFlg = False
        For lngCnt = 0 To lstResult.ListCount - 1

            If lstResult.Selected(lngCnt) Then

                strBuf = lstResult.List(lngCnt, C_SEARCH_ID)
                
                Set objShape = getObjFromID(Worksheets(selSheet), Mid$(strBuf, InStrRev(strBuf, ":") + 1))
                
                'SmartArtの場合
                If InStr(strBuf, C_SEARCH_ID_SMARTART) > 0 Then
                
                    Set objArt = getObjFromID2(Worksheets(selSheet), Mid$(strBuf, InStrRev(strBuf, ":") + 1))
                    
                    On Error Resume Next
                    If blnFlg Then
                        objShape.Shapes(1).Select False
                    Else
                        blnFlg = True
                        Application.GoTo setCellPos(objArt.TopLeftCell), True
                        objShape.Shapes(1).Select
                    End If
                    On Error GoTo 0
                Else
                    On Error Resume Next
                    If blnFlg Then
                        objShape.Select False
                    Else
                        blnFlg = True
                        Application.GoTo setCellPos(objShape.TopLeftCell), True
                        objShape.Select
                    End If
                    On Error GoTo 0
                End If

            End If
        Next

        Me.Show

    End If
End Sub
Private Function setCellPos(ByRef r As Range) As Range

    Dim lngRow As Long
    Dim lngCol As Long
    
    Dim lngCol1 As Long
    Dim lngCol2 As Long
    
    lngCol1 = Windows(1).VisibleRange(1).Column
    lngCol2 = Windows(1).VisibleRange(Windows(1).VisibleRange.count).Column
    
    Select Case r.Column
        Case lngCol1 To lngCol2
            lngCol = lngCol1
        Case Else
            lngCol = r.Column
    End Select

    Set setCellPos = r.Worksheet.Cells(r.Row, lngCol)

End Function
Private Sub lstResult_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call lstResult_Change

End Sub

Private Sub lstResult_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If lstResult.ListCount = 0 Then
        Exit Sub
    End If
    
    If Button = 2 Then
    
        With CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
            With .Controls.Add
                .Caption = "クリップボードにコピー"
                .OnAction = "basSearchEx.listCopy"
                .FaceId = 1436
            End With
            .ShowPopup
        End With
    
    End If
End Sub

Private Sub lstResult_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#If VBA7 And Win64 Then
#Else
    Set MW.obj = lstResult
#End If
End Sub

Private Sub schTab_Change()
    Select Case schTab.Value
        Case 0
            Dim c As Object
            For Each c In Controls
                If c.Tag = "v" Then
                    c.visible = False
                End If
            Next
            cboValue.enabled = True
        Case 1
            For Each c In Controls
                If c.Tag = "v" Then
                    c.visible = True
                End If
            Next
            cboValue.enabled = False
            cboValue.ListIndex = 0
    End Select
End Sub





Private Sub schTab_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#If VBA7 And Win64 Then
#Else
    Set MW.obj = Nothing
#End If

End Sub

'Private Sub scrTransparent_Change()
'    Tr.setTransparent scrTransparent.Value
'End Sub

'Private Sub UserForm_Activate()
''    Call FormResize
'
''    Me.Top = GetSetting(C_TITLE, "Search", "Top", (Application.Top + Application.Height - Me.Height) - 20)
''    Me.Left = GetSetting(C_TITLE, "Search", "Left", (Application.Left + Application.Width - Me.Width) - 20)
''    Me.Width = GetSetting(C_TITLE, "Search", "Width", Me.Width)
''    Me.Height = GetSetting(C_TITLE, "Search", "Height", Me.Height)
'
'
'
''    Call UserForm_Resize
'#If VBA7 And Win64 Then
'#Else
'    MW.Activate
'#End If
'End Sub

Private Sub UserForm_Initialize()
    
    Dim strBuf As String
    Dim strSearch() As String
    Dim strReplace() As String
    Dim i As Long
    
    Set RW = New ResizeWindow
    
    strBuf = GetSetting(C_TITLE, "Search", "SearchStr", "")
    strSearch = Split(strBuf, vbTab)
    
    For i = LBound(strSearch) To UBound(strSearch)
        txtSearch.AddItem strSearch(i)
    Next
    If txtSearch.ListCount > 0 Then
        txtSearch.ListIndex = 0
    End If
    
    strBuf = GetSetting(C_TITLE, "Search", "ReplaceStr", "")
    strReplace = Split(strBuf, vbTab)
    
    For i = LBound(strReplace) To UBound(strReplace)
        txtReplace.AddItem strReplace(i)
    Next
    If txtReplace.ListCount > 0 Then
        txtReplace.ListIndex = 0
    End If
    
    mblnRefresh = True
    
    cboPlace.AddItem C_SEARCH_PLACE_SHEET
    cboPlace.AddItem C_SEARCH_PLACE_SELECT
    cboPlace.AddItem C_SEARCH_PLACE_BOOK
    cboPlace.ListIndex = GetSetting(C_TITLE, "Search", "cboPlace", 0)
    
    cboObj.AddItem C_SEARCH_OBJECT_CELL_AND_SHAPE
    cboObj.AddItem C_SEARCH_OBJECT_CELL
    cboObj.AddItem C_SEARCH_OBJECT_SHAPE
    cboObj.ListIndex = GetSetting(C_TITLE, "Search", "cboObj", 0)
    
    cboValue.AddItem C_SEARCH_VALUE_FORMULA
    cboValue.AddItem C_SEARCH_VALUE_VALUE
    cboValue.ListIndex = GetSetting(C_TITLE, "Search", "cboValue", 0)
    
    schTab.Tabs(0).Caption = "検索"
    schTab.Tabs(1).Caption = "置換"
    schTab.Value = 0
    Call schTab_Change
    
    chkRegEx.Value = GetSetting(C_TITLE, "Search", "chkRegEx", False)
    chkCase.Value = GetSetting(C_TITLE, "Search", "chkCase", False)
    chkZenHan.Value = GetSetting(C_TITLE, "Search", "chkZenHan", False)
    chkSmartArt.Value = GetSetting(C_TITLE, "Search", "chkSmartArt", False)

    Me.Top = (Application.Top + Application.Height - Me.Height) - 20
    Me.Left = (Application.Left + Application.width - Me.width) - 20
    
    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    
    RW.FormWidth = Me.width
    RW.FormHeight = Me.Height
    
    mlngListWidth = Me.lstResult.width
    mlngListHeight = Me.lstResult.Height
    mlngTabWidth = Me.schTab.width
    mlngTabHeight = Me.schTab.Height
    mlngLblWidth = Me.lblSearch.width
    mlngLblObjLeft = Me.lblObj.Left
    mlngLblPlaceLeft = Me.lblPlace.Left
    mlngColumnWidth = Val(Split(Me.lstResult.ColumnWidths, ";")(1))

#If VBA7 And Win64 Then
#Else
    Set MW = basMouseWheel.GetInstance
    MW.Install Me
#End If
'
'    Set TR = New Transparent
'    TR.Init Me
'    TR.setTransparent 220

End Sub
Public Sub Start(ByVal lngTab As Long)

    schTab.Value = lngTab

    Me.Show

End Sub

Private Function convEscSeq(ByVal strBuf As String) As String

    Dim strRet As String
    
    strRet = Replace(strBuf, "\\", "\")
    strRet = Replace(strRet, "\n", vbLf)
    strRet = Replace(strRet, "\r", vbCr)
    strRet = Replace(strRet, "\t", vbTab)

    convEscSeq = strRet
    
End Function
Private Function getObjFromID2(ByRef WS As Worksheet, ByVal id As String) As Object

    Dim ret As Object
    Dim s As Shape
    
    Dim lngID As Long
    Dim lngPos As Long
    
    Set ret = Nothing
    
    If InStr(id, ",") > 0 Then
        lngID = CLng(Mid$(id, 1, InStrRev(id, ",") - 1))
    Else
        lngID = CLng(id)
    End If
    
    For Each s In WS.Shapes
        Select Case s.Type
            Case msoAutoShape, msoTextBox, msoSmartArt, msoCallout, msoFreeform
                If s.id = lngID Then
                    Set ret = s
                    Exit For
                End If
            
            Case msoGroup
                Set ret = getObjFromIDSub2(s, lngID)
                If ret Is Nothing Then
                Else
                    Exit For
                End If

        End Select
    Next
    Set getObjFromID2 = ret

End Function
Private Function getObjFromIDSub2(ByRef objShape As Shape, ByVal id As Long) As Object
    
    Dim s As Shape
    Dim ret As Object
    
    For Each s In objShape.GroupItems
        Select Case s.Type
            Case msoAutoShape, msoTextBox, msoSmartArt, msoCallout, msoFreeform
                If s.id = id Then
                    Set ret = s
                    Exit For
                End If
            
            Case msoGroup
                Set ret = getObjFromIDSub(s, id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If

        End Select
    Next

    Set getObjFromIDSub2 = ret
End Function
Private Function getObjFromID(ByRef WS As Worksheet, ByVal id As String) As Object
    Dim ret As Object
    Dim s As Shape
    
    For Each s In WS.Shapes
        Select Case s.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                If s.id = CLng(id) Then
                    Set ret = s
                    Exit For
                End If
            
            Case msoGroup
                Set ret = getObjFromIDSub(s, id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If
            Case msoSmartArt
                Set ret = getSmartArtFromIDSub(s, id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If
        End Select
    Next
    Set getObjFromID = ret

End Function
Private Function getObjFromIDSub(ByRef objShape As Shape, ByVal id As String) As Object
    
    Dim s As Shape
    Dim ret As Object
    
    For Each s In objShape.GroupItems
        Select Case s.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                If s.id = CLng(id) Then
                    Set ret = s
                    Exit For
                End If
            
            Case msoGroup
                Set ret = getObjFromIDSub(s, id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If
                
            Case msoSmartArt
                Set ret = getSmartArtFromIDSub(s, id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If
        End Select
    Next

    Set getObjFromIDSub = ret
End Function

Private Function getSmartArtFromIDSub(ByRef objShape As Shape, ByVal id As String) As Object
    
    Dim ret As Object
    Dim i As Long
    
    Dim lngID As Long
    Dim lngPos As Long
    
    Set ret = Nothing
    
    If InStr(id, ",") > 0 Then
        lngID = CLng(Mid$(id, 1, InStrRev(id, ",") - 1))
        lngPos = CLng(Mid$(id, InStrRev(id, ",") + 1))
        
        If lngID = objShape.id Then
        
            For i = 1 To objShape.SmartArt.AllNodes.count
            
                If i = lngPos Then
                    Set ret = objShape.SmartArt.AllNodes(i)
                    Exit For
                End If
            
            Next
        End If
    End If
    
    Set getSmartArtFromIDSub = ret
    
End Function

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#If VBA7 And Win64 Then
#Else
    Set MW.obj = Nothing
#End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
'    SaveSetting C_TITLE, "Search", "Top", Me.Top
'    SaveSetting C_TITLE, "Search", "Left", Me.Left
'    SaveSetting C_TITLE, "Search", "Width", Me.Width
'    SaveSetting C_TITLE, "Search", "Height", Me.Height
    
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    If RW.FormWidth > Me.width Then
        Me.width = RW.FormWidth
    End If
    If RW.FormHeight > Me.Height Then
        Me.Height = RW.FormHeight
    End If
    
    lstResult.width = mlngListWidth + (Me.width - RW.FormWidth)
    lstResult.Height = mlngListHeight + (Me.Height - RW.FormHeight)
    schTab.width = mlngTabWidth + (Me.width - RW.FormWidth)
    schTab.Height = mlngTabHeight + (Me.Height - RW.FormHeight)
    lblSearch.width = mlngLblWidth + (Me.width - RW.FormWidth)
    lblObj.Left = mlngLblObjLeft + (Me.width - RW.FormWidth)
    lblPlace.Left = mlngLblPlaceLeft + (Me.width - RW.FormWidth)
    
    
    Dim a As Variant
    a = Split(Me.lstResult.ColumnWidths, ";")
    
    a(1) = mlngColumnWidth + (Me.width - RW.FormWidth)
    Me.lstResult.ColumnWidths = Join(a, ";")
    
    
    
    DoEvents

    
End Sub

Private Sub UserForm_Terminate()

#If VBA7 And Win64 Then
#Else
    MW.UnInstall
    Set MW = Nothing
#End If
    
'    TR.Term
'    Set TR = Nothing

End Sub

Private Sub MW_WheelDown(obj As Object)

    On Error GoTo e

    If obj.ListCount = 0 Then Exit Sub
    obj.TopIndex = obj.TopIndex + 3
    
e:
End Sub

Private Sub MW_WheelUp(obj As Object)

    Dim lngPos As Long

    On Error GoTo e

    If obj.ListCount = 0 Then Exit Sub
    lngPos = obj.TopIndex - 3

    If lngPos < 0 Then
        lngPos = 0
    End If

    obj.TopIndex = lngPos
e:
End Sub
Public Sub listCopy()

    Dim strBuf As String
    
    Dim i As Long
    
    If lstResult.ListCount = 0 Then
    Else
        strBuf = ""
        For i = 0 To lstResult.ListCount - 1
        
            strBuf = strBuf & rep(lstResult.List(i, C_SEARCH_NO)) & vbTab
            strBuf = strBuf & rep(lstResult.List(i, C_SEARCH_STR)) & vbTab
            strBuf = strBuf & rep(lstResult.List(i, C_SEARCH_ADDRESS)) & vbTab
            strBuf = strBuf & rep(lstResult.List(i, C_SEARCH_SHEET)) & vbTab
'            strBuf = strBuf & rep(lstResult.List(i, C_SEARCH_ID)) & vbTab
            strBuf = strBuf & rep(lstResult.List(i, C_SEARCH_BOOK)) & vbCrLf
        
        Next
        SetClipText strBuf
    End If
End Sub
Private Function rep(ByVal strBuf As String) As String
    Dim strRet As String

    strRet = Replace(strBuf, vbLf, "")
    strRet = Replace(strRet, vbCr, "")
    strRet = Replace(strRet, vbCrLf, "")
    
    rep = strRet

End Function

