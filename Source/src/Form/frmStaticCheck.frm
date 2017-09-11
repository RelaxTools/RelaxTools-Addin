VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStaticCheck 
   Caption         =   "ブックの静的チェック"
   ClientHeight    =   7395
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

Private Sub lstResult_Click()


    Dim strAddress As String
    Dim strSheet As String
    Dim strBook As String
    Dim WB As Workbook
    Dim WS As Worksheet
    
    strBook = lstResult.List(lstResult.ListIndex, C_SEARCH_BOOK)
    strSheet = lstResult.List(lstResult.ListIndex, C_SEARCH_SHEET)
    strAddress = lstResult.List(lstResult.ListIndex, C_SEARCH_ADDRESS)
    
    Set WB = Workbooks(strBook)
    
    If Len(strSheet) <= 0 Then
        Set WB = Nothing
        Exit Sub
    End If
    
    On Error Resume Next
    
    Set WS = WB.Sheets(strSheet)
    If WS.visible = xlSheetVisible Then
        WS.Select
        If Len(strAddress) <= 0 Then
            Set WB = Nothing
            Set WS = Nothing
            Exit Sub
        End If
        
        WS.Range(strAddress).Select
    End If
   
    
End Sub

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
                    Call checkSheet1(lstContents.List(i, 1))
                Case 1
                    Call checkSheetNoUse(lstContents.List(i, 1))
                Case 2
                    Call checkSheetNoVisible(lstContents.List(i, 1))
                Case 3
                    Call checkSheetHyperlink(lstContents.List(i, 1))
                Case 4
                    Call checkSheetError(lstContents.List(i, 1))
                Case 5
                    Call checkSheetFormura(lstContents.List(i, 1))
                Case 6
                    Call checkSheetMerge(lstContents.List(i, 1))
                Case 7
                    Call checkSheetCol(lstContents.List(i, 1))
                Case 8
                    Call checkSheetRow(lstContents.List(i, 1))
                Case 9
                    Call checkSheetA1(lstContents.List(i, 1))
                Case 10
                    Call checkSheetZoom(lstContents.List(i, 1))
                Case 11
                    Call checkSheetNormal(lstContents.List(i, 1))
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
    
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub checkSheet1(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim RE As Object
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    Set RE = CreateObject("VBScript.RegExp")
    
    For Each WS In WB.Sheets
        With RE
            
            .Pattern = "^Sheet[0-9]+$"
            .IgnoreCase = False
            .Global = True
            
            If .Test(WS.Name) Then
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            End If
            
        End With
    Next
    
    Set RE = Nothing
    
End Sub
Private Sub checkSheetNoUse(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
            
        If Application.WorksheetFunction.CountA(WS.UsedRange) = 0 And WS.Shapes.count = 0 Then
            lstResult.AddItem ""
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = ""
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
        End If
        
    Next
    
    
End Sub
Private Sub checkSheetNoVisible(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
            
        If WS.visible = xlSheetHidden Or WS.visible = xlSheetVeryHidden Then
        
            lstResult.AddItem ""
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = ""
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
            lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
        End If
        
    Next
    
    
End Sub

Private Sub checkSheetA1(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    
    Set WB = ActiveWorkbook
    Set BS = WB.ActiveSheet
    
    For Each WS In WB.Sheets
            
        If WS.visible = xlSheetVisible Then
            WS.Select
            If WB.Windows(1).Selection.Address <> "$A$1" Then
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            End If
        End If
    Next
    BS.Select
    
End Sub
Private Sub checkSheetZoom(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    
    Set WB = ActiveWorkbook
    Set BS = WB.ActiveSheet
    
    For Each WS In WB.Sheets
            
        If WS.visible = xlSheetVisible Then
            WS.Select
            If WB.Windows(1).Zoom <> 100 Then
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            End If
        End If
    Next
    BS.Select
    
End Sub
Private Sub checkSheetNormal(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    
    Set WB = ActiveWorkbook
    Set BS = WB.ActiveSheet
    
    For Each WS In WB.Sheets
            
        If WS.visible = xlSheetVisible Then
            WS.Select
            If WB.Windows(1).View <> xlNormalView Then
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            End If
        End If
    Next
    BS.Select
    
End Sub
Private Sub checkSheetHyperlink(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim BS As Worksheet
    Dim HL As Hyperlink
    Dim r As Range
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
    
        For Each HL In WS.Hyperlinks
            
            If InStr(HL.Name, "\") > 0 Then
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = HL.Range.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            End If
        Next
        
        For Each r In WS.UsedRange
            If r.HasFormula And InStr(r.Formula, "\") > 0 Then
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = r.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            End If
        Next
        
    Next

    
End Sub
Private Sub checkSheetError(ByVal strCheck As String)
    
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
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = s.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            Next
        End If
        
        Err.Clear
        Set r = WS.UsedRange.SpecialCells(xlCellTypeConstants, xlErrors)
        If Err.Number = 0 Then
            For Each s In r
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = s.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            Next
        End If
        
    Next
    
End Sub
Private Sub checkSheetFormura(ByVal strCheck As String)
    
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
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = s.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            Next
        End If
        

        
    Next
    
End Sub
Private Sub checkSheetMerge(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim r As Range
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
    
        For Each r In WS.UsedRange
        
            If r.MergeCells Then
                If r.MergeArea(1).Address = r(1).Address Then
                    lstResult.AddItem ""
                    lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                    lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                    lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = r.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    
                    lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                    lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                    lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
                End If
            End If
        
        Next
    Next
    
End Sub
Private Sub checkSheetCol(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim i As Long
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
    
        For i = WS.UsedRange(1).Column To WS.UsedRange(WS.UsedRange.count).Column
            If WS.Columns(i).Hidden Then
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = WS.Columns(i).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            End If
        Next
    Next
    
End Sub
Private Sub checkSheetRow(ByVal strCheck As String)
    
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim i As Long
    
    Set WB = ActiveWorkbook
    
    For Each WS In WB.Sheets
    
        For i = WS.UsedRange(1).Row To WS.UsedRange(WS.UsedRange.count).Row
            If WS.Rows(i).Hidden Then
                lstResult.AddItem ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_NO) = lstResult.ListCount
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_STR) = strCheck
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ADDRESS) = WS.Rows(i).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_SHEET) = WS.Name
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_ID) = ""
                lstResult.List(lstResult.ListCount - 1, C_SEARCH_BOOK) = WB.Name
            End If
        Next
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
