VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
Private Const C_TITLE As String = "Grep置換"
Private Const C_SEARCH_NO As Long = 1
Private Const C_SEARCH_BOOK As Long = 2
Private Const C_SEARCH_SHEET As Long = 3
Private Const C_SEARCH_ADDRESS As Long = 4
Private Const C_SEARCH_STR As Long = 5
Private Const C_START_ROW As Long = 11

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim WB As Workbook
    Dim WS As Worksheet
    Dim strBook As String
    Dim strSheet As String
    Dim strId As String

    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If

    strBook = Cells(ActiveCell.row, C_SEARCH_BOOK).value
    If Len(strBook) = 0 Then
        Exit Sub
    End If
    strSheet = Cells(ActiveCell.row, C_SEARCH_SHEET).value
    If Len(strSheet) = 0 Then
        Exit Sub
    End If
    strId = Cells(ActiveCell.row, C_SEARCH_ADDRESS).value
    If Len(strId) = 0 Then
        Exit Sub
    End If

    'パスワードの取得
    Dim lngPos As Long
    Dim varPassword As Variant
    Dim varPass As Variant
    Dim strBuf As String
    
    lngPos = InStr(Cells(8, C_SEARCH_NO).value, "：")
    If lngPos > 0 Then
        strBuf = Mid(Cells(8, C_SEARCH_NO).value, lngPos + 1)
        varPassword = Split(strBuf, ",")
    Else
        varPassword = Array("")
    End If

    On Error Resume Next
    
    If strBuf = "" Then
        Set WB = Workbooks.Open(FileName:=strBook)
    Else
        For Each varPass In varPassword
            Err.Clear
            Set WB = Workbooks.Open(FileName:=strBook, Password:=varPass)
            If Err.Number = 0 Then
                Exit For
            End If
        Next
    End If
    If Err.Number <> 0 Then
        MsgBox "ブックを開けませんでした。", vbExclamation, "Grep"
        Exit Sub
    End If

    AppActivate Application.Caption

    Set WS = WB.Worksheets(strSheet)
    WS.Select
    
    WS.Range(strId).Select
    
    Dim s As Shape
    Set s = getObjFromID(WS, Mid$(strId, InStrRev(strId, ":") + 1))
    Application.Goto Reference:=setCellPos(s.TopLeftCell), Scroll:=True
    s.Select
    
End Sub
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

    Set setCellPos = r.Worksheet.Cells(r.row, lngCol)

End Function

Sub ReplaceStr()

    Dim WB As Workbook
    Dim WS As Worksheet
    Dim strBook As String
    Dim strSheet As String
    Dim strBookBk As String
    Dim strSheetBk As String
    Dim strId As String
    Dim lngRow As Long
    Dim s As Shape
    Dim FS As Object

    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    If Cells(6, 1).value = "検索対象：値" Then
        MsgBox "検索対象が値の場合、Grep置換はできません。", vbOKOnly + vbCritical, C_TITLE
        Exit Sub
    End If
    
    If MsgBox("Grep置換はβ版です。置換前のファイルは" & vbCrLf & "「ファイル名」+ YYYYMMDDHHMMSS の形式で" & vbCrLf & "保存されます。実行しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    On Error Resume Next
    Set FS = CreateObject("Scripting.FileSystemObject")

    lngRow = C_START_ROW
    
    'パスワードの取得
    Dim lngPos As Long
    Dim varPassword As Variant
    Dim varPass As Variant
    Dim strBuf As String
    
    lngPos = InStr(Cells(8, C_SEARCH_NO).value, "：")
    If lngPos > 0 Then
        strBuf = Mid(Cells(8, C_SEARCH_NO).value, lngPos + 1)
        varPassword = Split(strBuf, ",")
    Else
        varPassword = Array("")
    End If
    
    Do Until Cells(lngRow, C_SEARCH_NO).value = ""
        
        strBook = Cells(lngRow, C_SEARCH_BOOK).value
        
        'バックアップする
        FS.CopyFile strBook, strBook & "." & Format$(Now(), "yyyymmddhhnnss")
        
        If strBuf = "" Then
            Set WB = Workbooks.Open(FileName:=strBook, Local:=True)
        Else
            For Each varPass In varPassword
                Err.Clear
                Set WB = Workbooks.Open(FileName:=strBook, Password:=varPass, Local:=True)
                If Err.Number = 0 Then
                    Exit For
                End If
            Next
        End If
        If Err.Number <> 0 Then
            MsgBox "ブックが開けません(" & FS.getFileName(strBook) & ")です。" & vbCrLf & "Grep置換をスキップします。", vbOKOnly + vbCritical, C_TITLE
            lngRow = lngRow + 1
            GoTo pass
        End If
        
        AppActivate Application.Caption
            
        Do Until strBook <> Cells(lngRow, C_SEARCH_BOOK).value Or Cells(lngRow, C_SEARCH_NO).value = ""
        
            strSheet = Cells(lngRow, C_SEARCH_SHEET).value
            
            Set WS = WB.Worksheets(strSheet)
            WS.Select
            
            Do Until strBook <> Cells(lngRow, C_SEARCH_BOOK).value Or strSheet <> Cells(lngRow, C_SEARCH_SHEET).value Or Cells(lngRow, C_SEARCH_NO).value = ""
        
                strId = Cells(lngRow, C_SEARCH_ADDRESS).value
                
                Dim strSrch As String
                strSrch = Cells(lngRow, C_SEARCH_STR).value
                
                If InStr(strId, "$") > 0 Then
                    WS.Range(strId).value = strSrch
                Else
                    Set s = getObjFromID(WS, Mid$(strId, InStrRev(strId, ":") + 1))
                    s.TextFrame.Characters.Text = strSrch
                End If
                
                lngRow = lngRow + 1
            Loop
        Loop
        
        
        If WB.ReadOnly Then
            MsgBox "読み取り専用のブック(" & WB.Name & ")です。" & vbCrLf & "Grep置換をスキップします。", vbOKOnly + vbCritical, C_TITLE
            WB.Saved = True
        Else
            On Error GoTo e
            Application.DisplayAlerts = False
            WB.Save
            Application.DisplayAlerts = True
            On Error Resume Next
        End If
        
        WB.Close
pass:
        Set WB = Nothing
                
    Loop
    
    Set FS = Nothing
    
    MsgBox "Grep置換完了しました。", vbOKOnly + vbInformation, C_TITLE
    Exit Sub
e:
    If MsgBox("ブック保存時にエラーになりました。(" & Err.Description & ")" & vbCrLf & "再試行しますか？", vbYesNo + vbQuestion, C_TITLE) = vbYes Then
        Resume
    Else
        MsgBox "Grep置換を中断しました。", vbOKOnly + vbCritical, C_TITLE
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
'                Set ret = getSmartArtFromIDSub(s, ID)
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
'                Set ret = getSmartArtFromIDSub(s, ID)
'                If ret Is Nothing Then
'                Else
'                    Exit For
'                End If
        End Select
    Next

    Set getObjFromIDSub = ret
End Function

