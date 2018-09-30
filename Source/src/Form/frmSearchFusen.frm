VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchFusen 
   Caption         =   "付箋の検索"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10875
   OleObjectBlob   =   "frmSearchFusen.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmSearchFusen"
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
Private Const C_SEARCH_VISIBLE As Long = 2
Private Const C_SEARCH_SHEET As Long = 3
Private Const C_SEARCH_ADDRESS As Long = 4
Private Const C_SEARCH_ID As Long = 5
'Private Const C_SEARCH_ID_SHAPE As Long = 5

Private mlngCount As Long
Private mblnRefresh As Boolean
    Private Const C_SEARCH_ID_CELL As String = "Cell:"
Private Const C_SEARCH_ID_SHAPE As String = "Shape"
Private Const C_SEARCH_ID_SMARTART As String = "SmartArt"
Private Const C_SIZE As Long = 256
Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1
Private mstrhWnd As String

Private Sub searchShape(ByRef strPattern As String, ByRef objSheet As Worksheet)

    Dim matchCount As Long
    Dim objShape As Shape
    Dim objAct As Worksheet
    Dim c As Shape
    
    Dim strBuf As String

    For Each c In objSheet.Shapes
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform

                strBuf = c.AlternativeText
                matchCount = InStr(UCase(strBuf), UCase(strPattern))
                
                If matchCount > 0 Then
                
                    lstResult.AddItem ""
                    lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                    lstResult.List(mlngCount, C_SEARCH_STR) = Left(c.TextFrame2.TextRange.Text, C_SIZE)
                    If c.visible Then
                        lstResult.List(mlngCount, C_SEARCH_VISIBLE) = "表示"
                    Else
                        lstResult.List(mlngCount, C_SEARCH_VISIBLE) = "非表示"
                    End If
                    lstResult.List(mlngCount, C_SEARCH_SHEET) = objSheet.Name
                    lstResult.List(mlngCount, C_SEARCH_ADDRESS) = c.Name
                    lstResult.List(mlngCount, C_SEARCH_ID) = C_SEARCH_ID_SHAPE & ":" & c.id

                    mlngCount = mlngCount + 1
                    
                End If
            
            Case msoGroup
                grouprc objSheet, c, c, strPattern

        End Select
    Next

End Sub
'再帰にてグループ以下のシェイプを検索
Private Sub grouprc(ByRef WS As Worksheet, ByRef objTop As Shape, ByRef objShape As Shape, ByRef strPattern As String)

    Dim matchCount As Long
    Dim c As Shape
    Dim strBuf As String
    
    For Each c In objShape.GroupItems
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                'シェイプに文字があるかないか判断がつかないためエラー検出にて処理
                
                strBuf = c.AlternativeText
                matchCount = InStr(UCase(strBuf), UCase(strPattern))
                
                If matchCount > 0 Then
                
                    lstResult.AddItem ""
                    lstResult.List(mlngCount, C_SEARCH_NO) = mlngCount + 1
                    lstResult.List(mlngCount, C_SEARCH_STR) = Left(c.TextFrame2.TextRange.Text, C_SIZE)
                    If c.visible Then
                        lstResult.List(mlngCount, C_SEARCH_VISIBLE) = "表示"
                    Else
                        lstResult.List(mlngCount, C_SEARCH_VISIBLE) = "非表示"
                    End If
                    lstResult.List(mlngCount, C_SEARCH_SHEET) = WS.Name
                    lstResult.List(mlngCount, C_SEARCH_ADDRESS) = c.Name
                    lstResult.List(mlngCount, C_SEARCH_ID) = C_SEARCH_ID_SHAPE & getGroupId(c) & ":" & c.id

                    mlngCount = mlngCount + 1
                    
                End If
                
            Case msoGroup
                '再帰呼出
                grouprc WS, objTop, c, strPattern
            
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDsp_Click()
    
    On Error Resume Next
    
    Dim strBuf As String
    Dim objShape As Object
    Dim selSheet As String
    
    Dim lngCnt As Long
    
    For lngCnt = 0 To lstResult.ListCount - 1

        If lstResult.Selected(lngCnt) Then

            strBuf = lstResult.List(lngCnt, C_SEARCH_ID)
            selSheet = lstResult.List(lngCnt, C_SEARCH_SHEET)
            
            Set objShape = getObjFromID(Worksheets(selSheet), Mid$(strBuf, InStrRev(strBuf, ":") + 1))
            
            If objShape Is Nothing Then
            Else
                objShape.visible = True
                lstResult.List(lngCnt, C_SEARCH_VISIBLE) = "表示"
            End If
        End If
    Next

End Sub

Private Sub cmdNoDsp_Click()

    On Error Resume Next
    
    Dim strBuf As String
    Dim objShape As Object
    Dim selSheet As String
    
    Dim lngCnt As Long
    
    For lngCnt = 0 To lstResult.ListCount - 1

        If lstResult.Selected(lngCnt) Then

            strBuf = lstResult.List(lngCnt, C_SEARCH_ID)
            selSheet = lstResult.List(lngCnt, C_SEARCH_SHEET)
            
            Set objShape = getObjFromID(Worksheets(selSheet), Mid$(strBuf, InStrRev(strBuf, ":") + 1))
            
            If objShape Is Nothing Then
            Else
                objShape.visible = False
                lstResult.List(lngCnt, C_SEARCH_VISIBLE) = "非表示"
            End If
        End If
    Next

End Sub

Private Sub cmdOk_Click()
    
    On Error Resume Next
    
    Dim strBuf As String
    Dim objShape As Object
    Dim selSheet As String
    
    Dim lngCnt As Long
    
    For lngCnt = 0 To lstResult.ListCount - 1

        If lstResult.Selected(lngCnt) Then

            strBuf = lstResult.List(lngCnt, C_SEARCH_ID)
            selSheet = lstResult.List(lngCnt, C_SEARCH_SHEET)
            
            Set objShape = getObjFromID(Worksheets(selSheet), Mid$(strBuf, InStrRev(strBuf, ":") + 1))
            
            If objShape Is Nothing Then
            Else
                objShape.Delete
            End If
        End If
    Next
    
    dispList
    
End Sub

Private Sub lstResult_Change()

    If mblnRefresh = False Then
         Exit Sub
    End If

    Dim lngCnt As Long
    Dim strRange As String
    Dim r As Range
    Dim s As String
    
    Dim selSheet As String
    Dim blnCell As Boolean
'    Dim blnShape As Boolean
    Dim strPath As String
    selSheet = ""
    
    blnCell = False
    
    For lngCnt = 0 To lstResult.ListCount - 1
    
        If lstResult.Selected(lngCnt) Then
            If selSheet = "" Then
                txtPreview.Text = lstResult.List(lngCnt, C_SEARCH_STR)
                selSheet = lstResult.List(lngCnt, C_SEARCH_SHEET)
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
    
    Worksheets(selSheet).Select
    
    
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
                
        End Select
    Next

    Set getObjFromIDSub = ret
End Function
'Private Function setCellPos(ByRef r As Range) As Range
'
'    Dim lngRow As Long
'    Dim lngCol As Long
'
'    Dim lngCol1 As Long
'    Dim lngCol2 As Long
'
'    lngCol1 = Windows(1).VisibleRange(1).Column
'    lngCol2 = Windows(1).VisibleRange(Windows(1).VisibleRange.count).Column
'
'    If lngCol1 <= r.Column And r.Column <= lngCol2 Then
'        lngCol = lngCol1
'    Else
'        lngCol = r.Column
'    End If
'
'    lngRow = r.row - 5
'    If lngRow < 1 Then
'        lngRow = 1
'    End If
'
'    Set setCellPos = r.Worksheet.Cells(lngRow, lngCol)
'
'End Function
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

Private Sub lstResult_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstResult
End Sub

Private Sub UserForm_Initialize()

    dispList

    mblnRefresh = True
    
    Me.Top = (Application.Top + Application.Height - Me.Height) - 20
    Me.Left = (Application.Left + Application.width - Me.width) - 20
    
    mstrhWnd = CStr(FindWindow("ThunderDFrame", Me.Caption))
    Set MW = basMouseWheel.Install(mstrhWnd)
    
End Sub

Sub dispList()

    Dim objSheet As Worksheet
    
    Dim strText As String
    Dim strTag As String
    Dim varPrint As Variant
    
    Dim strWidth  As String
    Dim strHeight  As String
    
    Dim strFormat As String
    Dim strUserDate  As String
    Dim strFusenDate As String
    
    Dim strFont  As String
    Dim strSize  As String
    
    Dim strHorizontalAnchor  As String
    Dim strVerticalAnchor  As String
    
    Dim varAutoSize  As Variant
    Dim varOverFlow As Variant
    Dim varWordWrap As Variant
    
    lstResult.Clear
    txtPreview.Text = ""
    
    Call getSettingFusen(strText, strTag, varPrint, strWidth, strHeight, strFormat, strUserDate, strFusenDate, strFont, strSize, strHorizontalAnchor, strVerticalAnchor, varAutoSize, varOverFlow, varWordWrap)

    mlngCount = 0
    For Each objSheet In ActiveWorkbook.Worksheets

        If objSheet.visible = xlSheetVisible Then
            Call searchShape(strTag, objSheet)
        End If
        
    Next
End Sub
'Private Sub UserForm_Activate()
'    MW.Activate
'End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = Nothing
End Sub

Private Sub UserForm_Terminate()
    Set MW = basMouseWheel.UnInstall(mstrhWnd)
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
