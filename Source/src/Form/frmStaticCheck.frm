VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStaticCheck 
   Caption         =   "ブックの静的チェック"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   OleObjectBlob   =   "frmStaticCheck.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmStaticCheck"
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
Private Const C_SEARCH_NO As Long = 0
Private Const C_SEARCH_STR As Long = 1
Private Const C_SEARCH_ADDRESS As Long = 2
Private Const C_SEARCH_SHEET As Long = 3
Private Const C_SEARCH_ID As Long = 4
Private Const C_SEARCH_BOOK As Long = 5
Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1
Private Const C_SEARCH_ID_SHAPE As String = "Shape"
Private Const C_DEFAULT_CELL As String = "$A$1"
Private mblnRefresh As Boolean
Private Sub cmdAll_Click()
    Dim i As Long
    For i = 0 To lstContents.ListCount - 1
        lstContents.Selected(i) = True
    Next
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    For i = 0 To lstContents.ListCount - 1
        lstContents.Selected(i) = False
    Next
End Sub

Private Sub lstContents_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#If VBA7 And Win64 Then
#Else
    Set MW.obj = lstContents
#End If
End Sub

'Private Sub lstResult_Click()
'
'
'    Dim strAddress As String
'    Dim strSheet As String
'    Dim strBook As String
'    Dim WB As Workbook
'    Dim WS As Worksheet
'
'    strBook = lstResult.List(lstResult.ListIndex, C_SEARCH_BOOK)
'    strSheet = lstResult.List(lstResult.ListIndex, C_SEARCH_SHEET)
'    strAddress = lstResult.List(lstResult.ListIndex, C_SEARCH_ADDRESS)
'
'    Set WB = Workbooks(strBook)
'
'    If Len(strSheet) <= 0 Then
'        Set WB = Nothing
'        Exit Sub
'    End If
'
'    On Error Resume Next
'
'    Set WS = WB.Sheets(strSheet)
'    If WS.visible = xlSheetVisible Then
'        WS.Select
'        If Len(strAddress) <= 0 Then
'            Set WB = Nothing
'            Set WS = Nothing
'            Exit Sub
'        End If
'
'        WS.Range(strAddress).Select
'    End If
'
'
'End Sub
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
                If Left$(lstResult.List(lngCnt, C_SEARCH_ID), 1) = "$" Or lstResult.List(lngCnt, C_SEARCH_ID) = "" Then
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
    If Worksheets(selSheet).visible <> xlSheetVisible Then
        If MsgBox("非表示のシートです。表示しますか？", vbOKCancel + vbQuestion, C_TITLE) = vbOK Then
            Worksheets(selSheet).visible = xlSheetVisible
        Else
            Exit Sub
        End If
    End If
    Worksheets(selSheet).Select
    
    If blnCell Then
        For lngCnt = 0 To lstResult.ListCount - 1
    
            If lstResult.Selected(lngCnt) Then
                If r Is Nothing Then
'                    If lstResult.List(lngCnt, C_SEARCH_ID) = "" Then
'                    Else
                        Set r = Range("'[" & lstResult.List(lngCnt, C_SEARCH_BOOK) & "]" & lstResult.List(lngCnt, C_SEARCH_SHEET) & "'!" & lstResult.List(lngCnt, C_SEARCH_ID))
'                    End If
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
                
'                'SmartArtの場合
'                If InStr(strBuf, C_SEARCH_ID_SMARTART) > 0 Then
'
'                    Set objArt = getObjFromID2(Worksheets(selSheet), Mid$(strBuf, InStrRev(strBuf, ":") + 1))
'
'                    On Error Resume Next
'                    If blnFlg Then
'                        objShape.Shapes(1).Select False
'                    Else
'                        blnFlg = True
'                        Application.GoTo setCellPos(objArt.TopLeftCell), True
'                        objShape.Shapes(1).Select
'                    End If
'                    On Error GoTo 0
'                Else
                    On Error Resume Next
                    If blnFlg Then
                        objShape.Select False
                    Else
                        blnFlg = True
                        Application.GoTo setCellPos(objShape.TopLeftCell), True
                        objShape.Select
                    End If
                    On Error GoTo 0
'                End If

            End If
        Next

        Me.Show

    End If
End Sub
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
'            Case msoSmartArt
'                Set ret = getSmartArtFromIDSub(s, id)
'                If ret Is Nothing Then
'                Else
'                    Exit For
'                End If
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
                
'            Case msoSmartArt
'                Set ret = getSmartArtFromIDSub(s, id)
'                If ret Is Nothing Then
'                Else
'                    Exit For
'                End If
        End Select
    Next

    Set getObjFromIDSub = ret
End Function
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
Private Sub searchShape(ByRef objSheet As Worksheet, ByVal strCheck As String)

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
    
'    strPattern = txtSearch.Text
    
'    '正規表現の場合
'    If chkRegEx Then
'        Set mobjRegx = CreateObject("VBScript.RegExp")
'        mobjRegx.Pattern = strPattern
'        mobjRegx.IgnoreCase = Not (chkCase.value)
'        mobjRegx.Global = True
'    End If
    
    For Each c In objSheet.Shapes
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                'シェイプに文字があるかないか判断がつかないためエラー検出にて処理

                If c.TextFrame2.HasText Then
                    
                    
                    If c.TextFrame2.TextRange.Font.Name <> "ＭＳ ゴシック" Or c.TextFrame2.TextRange.Font.size <> 9 Then
                        ReportCheck strCheck, c.Name, objSheet.Name, C_SEARCH_ID_SHAPE & ":" & c.id, objSheet.Parent.Name
                    End If
                        
                Else
                    On Error GoTo 0
                    Err.Clear
                End If
            Case msoGroup
                grouprc objSheet, c, c, colShapes, strCheck


        End Select
        DoEvents
    Next

End Sub
'再帰にてグループ以下のシェイプを検索
Private Sub grouprc(ByRef WS As Worksheet, ByRef objTop As Shape, ByRef objShape As Shape, ByRef colShapes As Collection, ByVal strCheck As String)

    Dim matchCount As Long
    Dim c As Shape
    Dim strBuf As String
    Dim objMatch As Object
    Dim strPattern As String
'    strPattern = txtSearch.Text
    
    For Each c In objShape.GroupItems
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                'シェイプに文字があるかないか判断がつかないためエラー検出にて処理
                If c.TextFrame2.HasText Then
                    
                    If c.TextFrame2.TextRange.Font.Name <> "ＭＳ ゴシック" Or c.TextFrame2.TextRange.Font.size <> 9 Then
                        ReportCheck strCheck, c.Name, WS.Name, C_SEARCH_ID_SHAPE & getGroupId(c) & ":" & c.id, WS.Parent.Name
                    End If
                Else
                    On Error GoTo 0
                    Err.Clear
                End If
            Case msoGroup
                '再帰呼出
                grouprc WS, objTop, c, colShapes, strCheck
            
        End Select
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
Private Sub lstResult_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#If VBA7 And Win64 Then
#Else
    Set MW.obj = lstResult
#End If
End Sub

Private Sub UserForm_Initialize()

    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "シート：Sheet1、Sheet2 などの名前が無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "シート：Sheet1、Sheet2 などの名前を修正してください。"
    
    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "シート：使用されていないシートが無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "シート：使用されていないシートがあります。"
    
    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "シート：非表示のシートが無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "シート：非表示のシートがあります。"
    
    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "リンク：他ブックへの参照が無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "リンク：他ブックへの参照があります。"
    
    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "リンク：他ブックへのハイパーリンク切れをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "リンク：他ブックへのハイパーリンク切れがあります。"

    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "式　　：式のエラーが無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "式　　：式のエラーがあります。"

    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "式　　：式が存在し無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "式　　：式が存在します。"

    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "セル　：結合されたセルが無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "セル　：結合されたセルがあります。"

    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "列　　：非表示列が無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "列　　：非表示列があります。"
    
    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "行　　：非表示行が無いことをチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "行　　：非表示行があります。"
    
    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "お作法：カーソルがＡ１に設定されているかチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "お作法：カーソルがＡ１に設定されていません。"
    
    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "お作法：シートの倍率が１００％に設定されているかチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "お作法：シートの倍率が１００％に設定されていません。"
    
    lstContents.AddItem ""
    lstContents.List(lstContents.ListCount - 1, 0) = "お作法：表示スタイルが標準ビューに設定されているかチェックする。"
    lstContents.List(lstContents.ListCount - 1, 1) = "お作法：表示スタイルが標準ビューに設定されていません。"

    Dim i As Long
    For i = 0 To lstContents.ListCount - 1
        lstContents.Selected(i) = CBool(GetSetting(C_TITLE, "StaticCheck", CStr(i), False))
    Next
    
    Me.Top = (Application.Top + Application.Height - Me.Height) - 20
    Me.Left = (Application.Left + Application.width - Me.width) - 20
    mblnRefresh = True
#If VBA7 And Win64 Then
#Else
    Set MW = basMouseWheel.GetInstance
    MW.Install Me
#End If
    
End Sub


Private Sub cmdOk_Click()

    lstResult.Clear
    
    Dim i As Long
    For i = 0 To lstContents.ListCount - 1
        If lstContents.Selected(i) Then
            Select Case i
                Case 0
                    Call checkSheet1(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 1
                    Call checkSheetNoUse(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 2
                    Call checkSheetNoVisible(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 3
                    Call checkSheetHyperlink(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 4
                    Call checkBreakHyperlink(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 5
                    Call checkSheetError(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 6
                    Call checkSheetFormura(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 7
                    Call checkSheetMerge(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 8
                    Call checkSheetCol(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 9
                    Call checkSheetRow(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 10
                    Call checkSheetA1(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 11
                    Call checkSheetZoom(lstContents.List(i, 0), lstContents.List(i, 1))
                Case 12
                    Call checkSheetNormal(lstContents.List(i, 0), lstContents.List(i, 1))
            End Select
            
        End If
    Next
    
    
    Dim lngAns As Long
    
    txtChk.Value = lstResult.ListCount
    lngAns = txtTotal.Value - (Val(txtTen.Value) * Val(txtChk.Value))
    
    If lngAns < 0 Then
        txtANS.ForeColor = vbRed
    Else
        txtANS.ForeColor = vbBlack
    End If
    txtANS.Text = lngAns
    
    If lstResult.ListCount = 0 Then
        lblStatus.Caption = "エラーはありませんでした。"
    Else
        lblStatus.Caption = "エラーが" & lstResult.ListCount & "件あります。"
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub checkSheet1(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim re As Object
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    Set re = CreateObject("VBScript.RegExp")
    
    For Each WS In WB.Sheets
        With re
            
            .Pattern = "^Sheet[0-9]+$"
            .IgnoreCase = False
            .Global = True
            
            If .Test(WS.Name) Then
                ReportCheck strCheck, "-", WS.Name, C_DEFAULT_CELL, WB.Name
            End If
            
        End With
    Next
    
    Set re = Nothing
    
End Sub
Private Sub checkSheetNoUse(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
            
        If Application.WorksheetFunction.CountA(WS.UsedRange) = 0 And WS.Shapes.count = 0 Then
            ReportCheck strCheck, "-", WS.Name, C_DEFAULT_CELL, WB.Name
        End If
        
    Next
    
    
End Sub
Private Sub checkSheetNoVisible(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
            
        If WS.visible = xlSheetHidden Or WS.visible = xlSheetVeryHidden Then
            ReportCheck strCheck, "-", WS.Name, C_DEFAULT_CELL, WB.Name
        End If
        
    Next
    
    
End Sub

Private Sub checkSheetA1(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    
    Set WB = ActiveWorkbook
    Set BS = WB.ActiveSheet
    
    For Each WS In WB.Sheets
            
        If WS.visible = xlSheetVisible Then
            WS.Select
            If WB.Windows(1).Selection.Address <> "$A$1" Then
                ReportCheck strCheck, "-", WS.Name, C_DEFAULT_CELL, WB.Name
            End If
        End If
    Next
    BS.Select
    
End Sub
Private Sub checkSheetZoom(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    
    Set WB = ActiveWorkbook
    Set BS = WB.ActiveSheet
    
    For Each WS In WB.Sheets
            
        If WS.visible = xlSheetVisible Then
            WS.Select
            If WB.Windows(1).Zoom <> 100 Then
                ReportCheck strCheck, "-", WS.Name, C_DEFAULT_CELL, WB.Name
            End If
        End If
    Next
    BS.Select
    
End Sub
Private Sub checkSheetNormal(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    
    Set WB = ActiveWorkbook
    Set BS = WB.ActiveSheet
    
    For Each WS In WB.Sheets
            
        If WS.visible = xlSheetVisible Then
            WS.Select
            If WB.Windows(1).View <> xlNormalView Then
                ReportCheck strCheck, "-", WS.Name, C_DEFAULT_CELL, WB.Name
            End If
        End If
    Next
    BS.Select
    
End Sub
Private Sub checkSheetHyperlink(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    Dim HL As Hyperlink
    Dim r As Range
    Dim i As Long
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
    
        i = 0
        StartBar strCheckNm, WS.Hyperlinks.count + WS.UsedRange.count
        For Each HL In WS.Hyperlinks
            
            If InStr(HL.Address, "\") > 0 Then
                Select Case HL.Type
                    Case msoHyperlinkRange
                        ReportCheck strCheck, "-", WS.Name, HL.Range.Address, WB.Name
                    Case msoHyperlinkShape
                        ReportCheck strCheck, "-", WS.Name, HL.Shape.id, WB.Name
                End Select
            End If
            i = i + 1
            ReportBar i
        Next
        
        For Each r In WS.UsedRange
            If r.HasFormula And InStr(r.Formula, "=[") = 1 Then
                ReportCheck strCheck, "-", WS.Name, r.Address, WB.Name
            End If
            i = i + 1
            ReportBar i
        Next
        
        StopBar
        
        
'        'ブックのリンクがあったら解除
'        Dim LK As Variant
'        Dim a As Variant
'
'
'        a = WB.LinkSources(Type:=xlLinkTypeExcelLinks)
'
'        If Not IsEmpty(a) Then
'            For Each LK In WB.LinkSources(Type:=xlLinkTypeExcelLinks)
'                WB.BreakLink Name:=LK, Type:=xlLinkTypeExcelLinks
'            Next
'        End If

        

'
'        '画面レイアウトのリンクを削除
'        Dim ss As Picture
'        For Each ss In WS.Pictures
'            If ss.HasFormula And InStr(ss.Formula, "\") > 0 Then
'                ReportCheck strCheck, "-", WS.Name, r.Address, WB.Name
'            End If
'        Next
        
    Next

    
End Sub

Private Sub checkBreakHyperlink(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    Dim HL As Hyperlink
    Dim r As Range
    
    Set WB = ActiveWorkbook
    
    Dim i As Long
    For Each WS In WB.Sheets
        i = 0
        StartBar strCheckNm, WS.Hyperlinks.count
        For Each HL In WS.Hyperlinks
            If rlxIsExcelFile(HL.Address) Then
                If rlxIsFileExists(HL.Address) Then
                Else
                    Select Case HL.Type
                        Case msoHyperlinkRange
                            ReportCheck strCheck, "-", WS.Name, HL.Range.Address, WB.Name
                        Case msoHyperlinkShape
                            ReportCheck strCheck, "-", WS.Name, HL.Shape.id, WB.Name
                    End Select
                End If
            End If
            i = i + 1
            ReportBar i
        Next
        StopBar
    Next
        
End Sub

Private Sub checkSheetError(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    Dim HL As Hyperlink
    Dim r As Range
    Dim s As Range
    
    Set WB = ActiveWorkbook
    Set BS = WB.ActiveSheet
    
    For Each WS In WB.Sheets
    
        On Error Resume Next
        Err.Clear
        Set r = WS.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
        If Err.Number = 0 Then
            For Each s In r
                ReportCheck strCheck, s.Address, WS.Name, s.Address, WB.Name
            Next
        End If
        
        Err.Clear
        Set r = WS.UsedRange.SpecialCells(xlCellTypeConstants, xlErrors)
        If Err.Number = 0 Then
            For Each s In r
                ReportCheck strCheck, s.Address, WS.Name, s.Address, WB.Name
            Next
        End If
        
    Next
    
End Sub
Private Sub checkSheetFormura(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    Dim HL As Hyperlink
    Dim r As Range
    Dim s As Range
    
    Set WB = ActiveWorkbook
    Set BS = WB.ActiveSheet
    
    For Each WS In WB.Sheets
    
        On Error Resume Next
        Err.Clear
        Set r = WS.UsedRange.SpecialCells(xlCellTypeFormulas, xlLogical Or xlNumbers Or xlTextValues)
        If Err.Number = 0 Then
            For Each s In r
                ReportCheck strCheck, s.Address, WS.Name, s.Address, WB.Name
            Next
        End If
    Next
    
End Sub
Private Sub checkSheetMerge(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim r As Range
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
    
        Dim i As Long
        StartBar strCheckNm, WS.UsedRange.count
        For Each r In WS.UsedRange
        
            If r.MergeCells Then
                If r.MergeArea(1).Address = r(1).Address Then
                    ReportCheck strCheck, r.Address, WS.Name, r.Address, WB.Name
                End If
            End If
            i = i + 1
            ReportBar i
        Next
        StopBar
    Next
    
End Sub
Private Sub checkSheetCol(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim i As Long
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
    
'        Dim i As Long
        StartBar strCheckNm, WS.UsedRange.count
        For i = WS.UsedRange(1).Column To WS.UsedRange(WS.UsedRange.count).Column
            If WS.Columns(i).Hidden Then
                ReportCheck strCheck, WS.Columns(i).Address, WS.Name, WS.Columns(i).Address, WB.Name
            End If
            i = i + 1
            ReportBar i
        Next
        StopBar
    Next
    
End Sub
Private Sub checkSheetRow(ByVal strCheckNm As String, ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim i As Long
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
    
'        Dim i As Long
        StartBar strCheckNm, WS.UsedRange.count
        For i = WS.UsedRange(1).Row To WS.UsedRange(WS.UsedRange.count).Row
            If WS.Rows(i).Hidden Then
                ReportCheck strCheck, WS.Rows(i).Address, WS.Name, WS.Rows(i).Address, WB.Name
            End If
            i = i + 1
            ReportBar i
        Next
        StopBar
    Next
    
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#If VBA7 And Win64 Then
#Else
    Set MW.obj = Nothing
#End If
End Sub

Private Sub UserForm_Terminate()

#If VBA7 And Win64 Then
#Else
    MW.UnInstall
    Set MW = Nothing
#End If

    Dim i As Long
    For i = 0 To lstContents.ListCount - 1
         Call SaveSetting(C_TITLE, "StaticCheck", CStr(i), lstContents.Selected(i))
    Next
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
Sub ReportCheck(ByVal strCheck As String, ByVal strAddress As String, ByVal strSheet As String, ByVal strId As String, ByVal strBook As String)

    lstResult.AddItem ""
    lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
    lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
    lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = strAddress
    lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = strSheet
    lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = strId
    lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = strBook

End Sub

Private Sub StartBar(ByVal strMsg As String, ByVal lngMax As Long)

    lblBar.visible = True
    lblBar.width = 0
    lblBar.Caption = strMsg
    
    lblStatus.Caption = strMsg
'    lblStatus.TextAlign = fmTextAlignCenter
    lblStatus.Tag = strMsg
    lblBar.Tag = lngMax
    
    'ReportBar 1

End Sub
Private Sub ReportBar(ByVal lngPos As Long)

    Dim dblPercent As Double
    
    dblPercent = (lngPos / Val(lblBar.Tag))

    lblBar.width = lblStatus.width * dblPercent
    lblBar.Caption = lblStatus.Tag & " 処理中です..." & Fix(dblPercent * 100) & "%"
    lblStatus.Caption = lblBar.Caption
    DoEvents

End Sub

Private Sub StopBar()
    
    lblBar.visible = False
    lblBar.width = 0
    lblBar.Caption = ""
    
    lblStatus.Caption = ""
'    lblStatus.TextAlign = fmTextAlignLeft

End Sub
