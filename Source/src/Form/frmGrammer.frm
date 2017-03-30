VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrammer 
   Caption         =   "文法校正"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14865
   OleObjectBlob   =   "frmGrammer.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmGrammer"
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

'Private RW As ResizeWindow
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

'Private Const C_SEARCH_VALUE_FORMULA_NO As Long = 2
'Private Const C_SEACH_VALUE_VALUE_NO As Long = 1

Private Sub chkRegEx_Change()

'    chkZenHan.enabled = Not (chkRegEx.Value)

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
'    If Val(Application.Version) >= C_EXCEL_VERSION_2013 Then
    
        If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
        
        Dim WSH As Object
        
        Set WSH = CreateObject("WScript.Shell")
        
        Call WSH.Run(C_REGEXP_URL)
        
        Set WSH = Nothing
    
'    Else
'        frmHelp.Start "regexp"
'    End If
End Sub

Private Sub cmdOK_Click()


    

    Call searchStart
    
'    Dim strBuf As String
'    Dim i As Long
'    Dim lngCount As Long
'    Dim strSearch() As String
    

    
'    SaveSetting C_TITLE, "Search", "cboPlace", cboPlace.ListIndex
'    SaveSetting C_TITLE, "Search", "cboObj", cboObj.ListIndex
'    SaveSetting C_TITLE, "Search", "chkRegEx", chkRegEx.value
'    SaveSetting C_TITLE, "Search", "chkCase", chkCase.value
'    SaveSetting C_TITLE, "Search", "chkZenHan", chkZenHan.value
'    SaveSetting C_TITLE, "Search", "cboValue", cboValue.ListIndex
'    SaveSetting C_TITLE, "Search", "chkSmartArt", chkSmartArt.value
    
    If lstResult.ListCount = 0 Then
        MsgBox "校正対象が見つかりませんでした。", vbInformation + vbOKOnly, C_TITLE
    End If
    
End Sub




Private Sub searchStart()
    
    Dim colSheet As Collection
    Dim objSheet1 As Worksheet
    Dim objSheet2 As Worksheet
    Dim WD As Object
    Dim DC As Object
    
    lstResult.Clear
    mlngCount = 0

    Set colSheet = New Collection

    On Error Resume Next
    err.Clear
    Set WD = CreateObject("Word.Application")
    If err.Number <> 0 Then
        MsgBox "Wordがインストールされていないか、使用できません。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    Set DC = WD.Documents.Add
'    WD.visible = True

    WD.DisplayAlerts = False

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
                Call seachCell(WD, objSheet2)
                
            Case C_SEARCH_OBJECT_SHAPE
                Call searchShape(WD, objSheet2)
                
            Case C_SEARCH_OBJECT_CELL_AND_SHAPE
                Call seachCell(WD, objSheet2)
                Call searchShape(WD, objSheet2)
                
        End Select
    
    Next

    Set colSheet = Nothing
    
    

    WD.DisplayAlerts = True

    DC.Close False
    WD.Quit
    Set WD = Nothing

    
End Sub

Private Sub seachCell(ByRef WD As Object, ByRef objSheet As Worksheet)

    Dim strPattern As String
    Dim strRet As String
    
    strPattern = "*"
 
    Dim objFind As Range
    Dim strFirstAddress As String
    
    Set objFind = objSheet.UsedRange.Find(strPattern, , xlValues, xlPart, xlByRows, xlNext, False, False)
    
    If Not objFind Is Nothing Then
    
        strFirstAddress = objFind.Address

        Do
            If GetGrammer(WD, objFind.Value, strRet) Then

                lstResult.AddItem ""
                lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                
                lstResult.List(mlngCount, C_SEARCH_STR) = Left(strRet, C_SIZE)
                
                
                lstResult.List(mlngCount, C_SEARCH_ADDRESS) = objFind.Address
                lstResult.List(mlngCount, C_SEARCH_ID) = objFind.Address
                
                lstResult.List(mlngCount, C_SEARCH_SHEET) = objSheet.Name
                lstResult.List(mlngCount, C_SEARCH_BOOK) = objSheet.Parent.Name
    
                mlngCount = mlngCount + 1
            End If
            
    
            Set objFind = objSheet.UsedRange.FindNext(objFind)
            
            If objFind Is Nothing Then
                Exit Do
            End If
            
        Loop Until strFirstAddress = objFind.Address
        
    End If

    
End Sub

Private Function GetGrammer(ByRef WD As Object, ByVal Value As String, ByRef strRet As String) As Boolean

    
    Dim a As Object
    

    
    Dim i As Long
    Dim s As String
    Dim cnt As Long
    Dim ctl As Object
    Dim lngCnt  As Long
    Dim lngMax As Long
    
    GetGrammer = False

    WD.ActiveDocument.Range.Text = Value
    DoEvents
    
    lngMax = WD.ActiveDocument.GrammaticalErrors.count
    DoEvents
    
    For lngCnt = 1 To lngMax
    
        On Error Resume Next
        err.Clear
        Set a = WD.ActiveDocument.GrammaticalErrors(lngCnt)
        If err.Number <> 0 Then
            Exit For
        End If
        
        Select Case a.LanguageID
            '日本語のみ処理
            Case 1041
        
                  For i = 1 To Len(a.Text)
                
                      cnt = 0 '初期化
                  
                      a.Characters(i).Select
                  
                      '修正候補をCommandBarControlから取得
                  
                      For Each ctl In WD.CommandBars("Grammar").Controls
                  
                        '[IDが「0」のもの = 修正候補]として取得
                        If ctl.Id = 0 Then
                        
                            If cnt < 1 Then
                                s = ctl.Caption
                            Else
                                s = s & "," & ctl.Caption
                            End If
                            cnt = cnt + 1
                        Else
                            Exit For
                        End If
                      Next
                      If cnt > 0 Then Exit For
                    Next
                    
                    strRet = a.Text & " " & s
                    GetGrammer = True
                    
        End Select
    Next

'    '英語のスペルミスを列挙して修正候補をコメントとして追加するWordマクロ
'    Dim rngSpellingError As Object
'    Dim ssgn As Object
'
'    'スペルミスを列挙
'    For Each a In WD.ActiveDocument.SpellingErrors
'        cnt = 0 '初期化
'        Select Case a.LanguageID
'            '英語のみ処理
'            Case 1033
'            '修正候補取得
'            For Each ssgn In a.GetSpellingSuggestions
'                If cnt < 1 Then
'                    s = ssgn.Name
'                Else
'                    s = s & "," & ssgn.Name
'                End If
'                cnt = cnt + 1
'            Next
'            'エラー箇所に修正候補をコメントとして追加
'
'            strRet = a.Text & " 修正候補:" & s
'            GetGrammer = True
'
'        End Select
'    Next
    
    DoEvents

End Function
Private Sub searchShape(ByRef WD As Object, ByRef objSheet As Worksheet)

    Dim matchCount As Long
    Dim objMatch As Object
    Dim strPattern As String

    Dim objShape As Shape
    Dim objAct As Worksheet
    Dim c As Shape
    
    Dim strBuf As String
    Dim strRet As String

    Dim colShapes As Collection
    Set colShapes = New Collection

    Const C_RESULT_NAME As String = "シェイプ検索Result"
    
'    strPattern = txtSearch.Text
    

    
    For Each c In objSheet.Shapes
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                'シェイプに文字があるかないか判断がつかないためエラー検出にて処理

                If c.TextFrame2.HasText Then
                    strBuf = c.TextFrame2.TextRange.Text
                    
                    If GetGrammer(WD, strBuf, strRet) Then
                    
                        lstResult.AddItem ""
                        lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                        lstResult.List(mlngCount, C_SEARCH_STR) = Left(strRet, C_SIZE)
                        lstResult.List(mlngCount, C_SEARCH_ADDRESS) = c.Name
                        lstResult.List(mlngCount, C_SEARCH_ID) = C_SEARCH_ID_SHAPE & ":" & c.Id
                        lstResult.List(mlngCount, C_SEARCH_SHEET) = objSheet.Name
                        lstResult.List(mlngCount, C_SEARCH_BOOK) = objSheet.Parent.Name

                        mlngCount = mlngCount + 1
                        
                    End If
                Else
                    On Error GoTo 0
                    err.Clear
                End If
            Case msoGroup
                grouprc objSheet, c, c, colShapes, WD

'            Case msoSmartArt
'                'check があるときのみ検索
'                If chkSmartArt.value Then
'                    SmartArtprc objSheet, c, c, colShapes
'                End If

        End Select
    Next

End Sub
'再帰にてグループ以下のシェイプを検索
Private Sub grouprc(ByRef WS As Worksheet, ByRef objTop As Shape, ByRef objShape As Shape, ByRef colShapes As Collection, ByRef WD As Object)

    Dim matchCount As Long
    Dim c As Shape
    Dim strBuf As String
    Dim objMatch As Object
    Dim strPattern As String
    Dim strRet As String
    
    For Each c In objShape.GroupItems
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                'シェイプに文字があるかないか判断がつかないためエラー検出にて処理
                If c.TextFrame2.HasText Then
                    strBuf = c.TextFrame2.TextRange.Text
                    
                    If GetGrammer(WD, strBuf, strRet) Then

                        lstResult.AddItem ""
                        lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                        lstResult.List(mlngCount, C_SEARCH_STR) = Left(strRet, C_SIZE)

                        lstResult.List(mlngCount, C_SEARCH_ADDRESS) = c.Name
                        lstResult.List(mlngCount, C_SEARCH_ID) = C_SEARCH_ID_SHAPE & getGroupId(c) & ":" & c.Id
                        lstResult.List(mlngCount, C_SEARCH_SHEET) = WS.Name
                        lstResult.List(mlngCount, C_SEARCH_BOOK) = WS.Parent.Name

                        mlngCount = mlngCount + 1
                    
                    End If
                Else
                    On Error GoTo 0
                    err.Clear
                End If
            Case msoGroup
                '再帰呼出
                grouprc WS, objTop, c, colShapes, WD
            
'            Case msoSmartArt
'                'check があるときのみ検索
'                If chkSmartArt.value Then
'                    SmartArtprc WS, c, c, colShapes
'                End If
        End Select
    Next

End Sub
''スマートアートを検索
'Private Sub SmartArtprc(ByRef WS As Worksheet, ByRef objTop As Shape, ByRef objShape As Shape, ByRef colShapes As Collection)
'
'    Dim matchCount As Long
'    Dim c As SmartArtNode
'    Dim strBuf As String
'    Dim objMatch As Object
'    Dim strPattern As String
'    Dim lngIdx As Long
'    strPattern = txtSearch.Text
'
'
'    For lngIdx = 1 To objShape.SmartArt.AllNodes.count
'
'        Set c = objShape.SmartArt.AllNodes(lngIdx)
'
'        'シェイプに文字があるかないか判断がつかないためエラー検出にて処理
'        If c.TextFrame2.HasText Then
'
'            strBuf = c.TextFrame2.TextRange.Text
'
'            '正規表現の場合
'            If chkRegEx Then
'                err.Clear
'                On Error Resume Next
'                Set objMatch = mobjRegx.Execute(strBuf)
'                If err.Number <> 0 Then
'                    MsgBox "検索文字列の正規表現が正しくありません。", vbExclamation, C_TITLE
'                    txtSearch.SetFocus
'                    Exit Sub
'                End If
'                matchCount = objMatch.count
'            Else
'                If chkCase.value Then
'                    matchCount = InStr(strBuf, strPattern)
'                Else
'                    matchCount = InStr(UCase(strBuf), UCase(strPattern))
'                End If
'            End If
'
'            If matchCount > 0 Then
'                lstResult.AddItem ""
'                lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
'                lstResult.List(mlngCount, C_SEARCH_STR) = Left(strBuf, C_SIZE)
'                lstResult.List(mlngCount, C_SEARCH_ADDRESS) = objShape.Name
'                lstResult.List(mlngCount, C_SEARCH_ID) = C_SEARCH_ID_SMARTART & getGroupId(objShape) & "/" & objShape.id & ":" & objShape.id & "," & lngIdx
'                lstResult.List(mlngCount, C_SEARCH_SHEET) = WS.Name
'                lstResult.List(mlngCount, C_SEARCH_BOOK) = WS.Parent.Name
'                mlngCount = mlngCount + 1
'            End If
'        Else
'            On Error GoTo 0
'            err.Clear
'        End If
'
'    Next
'
'End Sub
'グループ文字列を取得
Private Function getGroupId(ByRef objShape As Object) As String

    Dim strBuf As String
    Dim s As Object
    
    On Error Resume Next
    err.Clear
    Set s = objShape.ParentGroup
    Do Until err.Number <> 0
        strBuf = "/" & s.Id & strBuf
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
'                If r Is Nothing Then
'                    Set r = Range("'" & lstResult.List(lngCnt, C_SEARCH_SHEET) & "'!" & lstResult.List(lngCnt, C_SEARCH_ID))
'                Else
'                    Set r = Union(r, Range("'" & lstResult.List(lngCnt, C_SEARCH_SHEET) & "'!" & lstResult.List(lngCnt, C_SEARCH_ID)))
'                End If
                If r Is Nothing Then
                    Set r = Range("'[" & lstResult.List(lngCnt, C_SEARCH_BOOK) & "]" & lstResult.List(lngCnt, C_SEARCH_SHEET) & "'!" & lstResult.List(lngCnt, C_SEARCH_ID))
                Else
                    Set r = Union(r, Range("'[" & lstResult.List(lngCnt, C_SEARCH_BOOK) & "]" & lstResult.List(lngCnt, C_SEARCH_SHEET) & "'!" & lstResult.List(lngCnt, C_SEARCH_ID)))
                End If
    
            
            End If
        Next
        If r Is Nothing Then
        Else
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
    
    If lngCol1 <= r.Column And r.Column <= lngCol2 Then
        lngCol = lngCol1
    Else
        lngCol = r.Column
    End If
    
    lngRow = r.Row - 5
    If lngRow < 1 Then
        lngRow = 1
    End If

    Set setCellPos = r.Worksheet.Cells(lngRow, lngCol)

End Function
Private Sub lstResult_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call lstResult_Change

End Sub

Private Sub lstResult_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstResult
End Sub

'Private Sub schTab_Change()
'    Select Case schTab.value
'        Case 0
'            Dim c As Object
'            For Each c In Controls
'                If c.Tag = "v" Then
'                    c.visible = False
'                End If
'            Next
'            cboValue.enabled = True
'        Case 1
'            For Each c In Controls
'                If c.Tag = "v" Then
'                    c.visible = True
'                End If
'            Next
'            cboValue.enabled = False
'            cboValue.ListIndex = 0
'    End Select
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
'    MW.Activate
'End Sub

Private Sub UserForm_Initialize()
    
    Dim strBuf As String
    Dim strSearch() As String
    Dim strReplace() As String
    Dim i As Long
    
'    Set RW = New ResizeWindow
    
'    strBuf = GetSetting(C_TITLE, "Search", "SearchStr", "")
'    strSearch = Split(strBuf, vbVerticalTab)
'
'    For i = LBound(strSearch) To UBound(strSearch)
'        txtSearch.AddItem strSearch(i)
'    Next
'    If txtSearch.ListCount > 0 Then
'        txtSearch.ListIndex = 0
'    End If
    
'    strBuf = GetSetting(C_TITLE, "Search", "ReplaceStr", "")
'    strReplace = Split(strBuf, vbVerticalTab)
'
'    For i = LBound(strReplace) To UBound(strReplace)
'        txtReplace.AddItem strReplace(i)
'    Next
'    If txtReplace.ListCount > 0 Then
'        txtReplace.ListIndex = 0
'    End If
    
    mblnRefresh = True
    
    cboPlace.AddItem C_SEARCH_PLACE_SHEET
    cboPlace.AddItem C_SEARCH_PLACE_SELECT
    cboPlace.AddItem C_SEARCH_PLACE_BOOK
    cboPlace.ListIndex = GetSetting(C_TITLE, "Search", "cboPlace", 0)
    
    cboObj.AddItem C_SEARCH_OBJECT_CELL_AND_SHAPE
    cboObj.AddItem C_SEARCH_OBJECT_CELL
    cboObj.AddItem C_SEARCH_OBJECT_SHAPE
    cboObj.ListIndex = GetSetting(C_TITLE, "Search", "cboObj", 0)
    
'    cboValue.AddItem C_SEARCH_VALUE_FORMULA
'    cboValue.AddItem C_SEARCH_VALUE_VALUE
'    cboValue.ListIndex = GetSetting(C_TITLE, "Search", "cboValue", 0)
    
'    schTab.Tabs(0).Caption = "検索"
'    schTab.Tabs(1).Caption = "置換"
'    schTab.value = 0
'    Call schTab_Change
    
'    chkRegEx.value = GetSetting(C_TITLE, "Search", "chkRegEx", False)
'    chkCase.value = GetSetting(C_TITLE, "Search", "chkCase", False)
'    chkZenHan.value = GetSetting(C_TITLE, "Search", "chkZenHan", False)
'    chkSmartArt.value = GetSetting(C_TITLE, "Search", "chkSmartArt", False)

    Me.Top = (Application.Top + Application.Height - Me.Height) - 20
    Me.Left = (Application.Left + Application.Width - Me.Width) - 20
'
'    With txtSearch
'        .SelStart = 0
'        .SelLength = Len(.Text)
'    End With
'
    
'    RW.FormWidth = Me.Width
'    RW.FormHeight = Me.Height
    
'    mlngListWidth = Me.lstResult.Width
'    mlngListHeight = Me.lstResult.Height
'    mlngTabWidth = Me.schTab.Width
'    mlngTabHeight = Me.schTab.Height
'    mlngLblWidth = Me.lblSearch.Width
'    mlngLblObjLeft = Me.lblObj.Left
'    mlngLblPlaceLeft = Me.lblPlace.Left
'    mlngColumnWidth = Val(Split(Me.lstResult.ColumnWidths, ";")(1))

    Set MW = basMouseWheel.GetInstance
    MW.Install Me
    
End Sub
Public Sub Start(ByVal lngTab As Long)

'    schTab.value = lngTab

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
Private Function getObjFromID2(ByRef WS As Worksheet, ByVal Id As String) As Object

    Dim ret As Object
    Dim s As Shape
    
    Dim lngID As Long
    Dim lngPos As Long
    
    Set ret = Nothing
    
    If InStr(Id, ",") > 0 Then
        lngID = CLng(Mid$(Id, 1, InStrRev(Id, ",") - 1))
    Else
        lngID = CLng(Id)
    End If
    
    For Each s In WS.Shapes
        Select Case s.Type
            Case msoAutoShape, msoTextBox, msoSmartArt, msoCallout, msoFreeform
                If s.Id = lngID Then
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
Private Function getObjFromIDSub2(ByRef objShape As Shape, ByVal Id As Long) As Object
    
    Dim s As Shape
    Dim ret As Object
    
    For Each s In objShape.GroupItems
        Select Case s.Type
            Case msoAutoShape, msoTextBox, msoSmartArt, msoCallout, msoFreeform
                If s.Id = Id Then
                    Set ret = s
                    Exit For
                End If
            
            Case msoGroup
                Set ret = getObjFromIDSub(s, Id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If

        End Select
    Next

    Set getObjFromIDSub2 = ret
End Function
Private Function getObjFromID(ByRef WS As Worksheet, ByVal Id As String) As Object
    Dim ret As Object
    Dim s As Shape
    
    For Each s In WS.Shapes
        Select Case s.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                If s.Id = CLng(Id) Then
                    Set ret = s
                    Exit For
                End If
            
            Case msoGroup
                Set ret = getObjFromIDSub(s, Id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If
            Case msoSmartArt
                Set ret = getSmartArtFromIDSub(s, Id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If
        End Select
    Next
    Set getObjFromID = ret

End Function
Private Function getObjFromIDSub(ByRef objShape As Shape, ByVal Id As String) As Object
    
    Dim s As Shape
    Dim ret As Object
    
    For Each s In objShape.GroupItems
        Select Case s.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                If s.Id = CLng(Id) Then
                    Set ret = s
                    Exit For
                End If
            
            Case msoGroup
                Set ret = getObjFromIDSub(s, Id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If
                
            Case msoSmartArt
                Set ret = getSmartArtFromIDSub(s, Id)
                If ret Is Nothing Then
                Else
                    Exit For
                End If
        End Select
    Next

    Set getObjFromIDSub = ret
End Function

Private Function getSmartArtFromIDSub(ByRef objShape As Shape, ByVal Id As String) As Object
    
    Dim ret As Object
    Dim i As Long
    
    Dim lngID As Long
    Dim lngPos As Long
    
    Set ret = Nothing
    
    If InStr(Id, ",") > 0 Then
        lngID = CLng(Mid$(Id, 1, InStrRev(Id, ",") - 1))
        lngPos = CLng(Mid$(Id, InStrRev(Id, ",") + 1))
        
        If lngID = objShape.Id Then
        
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
    Set MW.obj = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
'    SaveSetting C_TITLE, "Search", "Top", Me.Top
'    SaveSetting C_TITLE, "Search", "Left", Me.Left
'    SaveSetting C_TITLE, "Search", "Width", Me.Width
'    SaveSetting C_TITLE, "Search", "Height", Me.Height
    
End Sub

Private Sub UserForm_Resize()
'    On Error Resume Next
'    If RW.FormWidth > Me.Width Then
'        Me.Width = RW.FormWidth
'    End If
'    If RW.FormHeight > Me.Height Then
'        Me.Height = RW.FormHeight
'    End If
    
'    lstResult.Width = mlngListWidth + (Me.Width - RW.FormWidth)
'    lstResult.Height = mlngListHeight + (Me.Height - RW.FormHeight)
'    schTab.Width = mlngTabWidth + (Me.Width - RW.FormWidth)
'    schTab.Height = mlngTabHeight + (Me.Height - RW.FormHeight)
'    lblSearch.Width = mlngLblWidth + (Me.Width - RW.FormWidth)
'    lblObj.Left = mlngLblObjLeft + (Me.Width - RW.FormWidth)
'    lblPlace.Left = mlngLblPlaceLeft + (Me.Width - RW.FormWidth)
'
'
'    Dim a As Variant
'    a = Split(Me.lstResult.ColumnWidths, ";")
'
'    a(1) = mlngColumnWidth + (Me.Width - RW.FormWidth)
'    Me.lstResult.ColumnWidths = Join(a, ";")
'
'
'
'    DoEvents

    
End Sub

Private Sub UserForm_Terminate()

    MW.UnInstall
    Set MW = Nothing

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
