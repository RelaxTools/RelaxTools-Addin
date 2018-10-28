VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPageList 
   Caption         =   "ページ数の取得"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   OleObjectBlob   =   "frmPageList.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmPageList"
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
'Private mobjRegx As Object
'Private mlngCount As Long
Private mblnCancel As Boolean

Private mblnRefresh As Boolean

Private Const C_START_ROW As Long = 2
Private Const C_SEARCH_NO As Long = 1
Private Const C_SEARCH_BOOK As Long = 2
Private Const C_SEARCH_PAGE As Long = 3
'Private Const C_SEARCH_ADDRESS As Long = 4
'Private Const C_SEARCH_STR As Long = 5
'Private Const C_SEARCH_ID As Long = 6

'Private Const C_SEARCH_OBJECT_CELL = "セルのみ"
'Private Const C_SEARCH_OBJECT_SHAPE = "シェイプのみ"
'Private Const C_SEARCH_OBJECT_CELL_AND_SHAPE = "セル＆シェイプ"
'Private Const C_SEARCH_VALUE_VALUE = "値"
'Private Const C_SEARCH_VALUE_FORMULA = "式"



Private Const C_WORD_FILE As String = ".DOC"
Private Const C_EXCEL_FILE As String = ".XLS"
Private Const C_PPT_FILE As String = ".PPT"

Private mMm As MacroManager




Private Sub cmdCancel_Click()
    If cmdCancel.Caption = "閉じる" Then
        Unload Me
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



Private Sub cmdOk_Click()

    Dim XL As Excel.Application
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Dim colBook As Collection
    Dim varBook As Variant
    Dim objFs As Object
    Dim lngBookCount As Long
    Dim lngCount As Long
    Dim lngBookMax As Long
    
    Dim ResultWS As Worksheet
    Dim ResultWB As Workbook
    
    Dim strPath As String
    Dim strPtn As String
    Dim strPatterns() As String
    
    Dim lngPage As Long
   
    If Len(Trim(cboFolder.Text)) = 0 Then
        MsgBox "フォルダを指定してください...", vbExclamation, C_TITLE
        cboFolder.SetFocus
        Exit Sub
    End If
    
    strPath = cboFolder.Text
    
    strPtn = ""
    
    If chkExcel.Value Then
        strPtn = "*.xls;*.xlsx"
    End If
    
    If chkWord.Value Then
        If strPtn = "" Then
            strPtn = strPtn & "*.doc;*.docx"
        Else
            strPtn = strPtn & ";*.doc;*.docx"
        End If
    End If
    
    If chkPoint.Value Then
        If strPtn = "" Then
            strPtn = strPtn & "*.ppt;*.pptx"
        Else
            strPtn = strPtn & ";*.ppt;*.pptx"
        End If
    End If
    
    strPatterns = Split(strPtn, ";")
    

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
    
    On Error GoTo e
    
    ThisWorkbook.Worksheets("ページ数カウント結果").Copy
    Set ResultWB = Application.Workbooks(Application.Workbooks.Count)
    
    'Application.ScreenUpdating の代わり
    ResultWB.Windows(1).visible = False
    
    Set ResultWS = ResultWB.Worksheets(1)
    
    ResultWS.Cells(1, C_SEARCH_NO).Value = "No."
    ResultWS.Cells(1, C_SEARCH_BOOK).Value = "ファイル名"
    ResultWS.Cells(1, C_SEARCH_BOOK).ColumnWidth = 60
    ResultWS.Cells(1, C_SEARCH_PAGE).Value = "ページ数"

    lngCount = C_START_ROW

    AppActivate Me.Caption
    
    cmdCancel.Caption = "キャンセル"
    
    If chkExcel.Value Then
        Set XL = New Excel.Application
    End If
    If chkWord.Value Then
        Dim WD As Object
        Dim dc As Object
        Set WD = CreateObject("Word.Application")
    End If
    If chkPoint.Value Then
        Dim PP As Object
        Dim pt As Object
        Set PP = CreateObject("PowerPoint.Application")
    End If

    lngBookCount = 0
    lngBookMax = colBook.Count
    mMm.StartGauge lngBookMax
    
    For Each varBook In colBook
    
        If mblnCancel Then
            Exit For
        End If
    
        Err.Clear
        
        ResultWS.Cells(lngCount, C_SEARCH_NO).Value = lngBookCount + 1 'lngCount - C_START_ROW + 1
        ResultWS.Cells(lngCount, C_SEARCH_BOOK).Value = varBook
    
        ResultWS.Hyperlinks.Add _
            Anchor:=ResultWS.Cells(lngCount, C_SEARCH_BOOK), _
            Address:="", _
            SubAddress:="", _
            TextToDisplay:=varBook
        
        Select Case True
            Case InStr(UCase(varBook), C_EXCEL_FILE) > 0
            
                Set WB = XL.Workbooks.Open(filename:=varBook, ReadOnly:=True, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True)
                
                Dim w As Long
                lngPage = 0
                w = lngCount
                
                For Each WS In WB.Worksheets
                    If WS.visible = xlSheetVisible Then
                    
                        Dim p As Long
                        
                        'p = (WS.VPageBreaks.count + 1) * (WS.HPageBreaks.count + 1)
                        WB.Windows(1).View = xlPageBreakPreview
                        p = WS.PageSetup.Pages.Count
                        
                        If chkExcelSheet.Value Then
                            lngCount = lngCount + 1
                            ResultWS.Cells(lngCount, C_SEARCH_PAGE).Value = p
                            ResultWS.Cells(lngCount, C_SEARCH_BOOK).Value = "  " & WS.Name
                        End If
                        
                        lngPage = lngPage + p
                        
                    End If
                Next
               
                
                ResultWS.Cells(w, C_SEARCH_PAGE).Value = lngPage
                WB.Close SaveChanges:=False
                Set WB = Nothing
        
            Case InStr(UCase(varBook), C_WORD_FILE) > 0
            
                Set dc = WD.Documents.Open(filename:=varBook, ReadOnly:=True)
                
                dc.Repaginate
                
                ResultWS.Cells(lngCount, C_SEARCH_PAGE).Value = dc.BuiltinDocumentProperties(14)
                
                dc.Close SaveChanges:=False
                Set dc = Nothing
                
            Case InStr(UCase(varBook), C_PPT_FILE) > 0
            
                Set pt = PP.Presentations.Open(filename:=varBook, ReadOnly:=True, withwindow:=False)
                    
                ResultWS.Cells(lngCount, C_SEARCH_PAGE).Value = pt.Slides.Count
                
                pt.Close
                Set pt = Nothing
                
        End Select
        
        lngBookCount = lngBookCount + 1
        lngCount = lngCount + 1
        mMm.DisplayGauge lngBookCount
    Next
e:
    If chkPoint.Value Then
        PP.Quit
        Set PP = Nothing
    End If
    If chkWord.Value Then
        WD.Quit
        Set WD = Nothing
    End If
    If chkExcel.Value Then
        XL.Quit
        Set XL = Nothing
    End If
    
    ResultWB.Windows(1).visible = True
    DoEvents
    
    Dim r As Range
    Set r = ResultWS.Cells(C_START_ROW, 1).CurrentRegion
    
    r.VerticalAlignment = xlTop
    r.Select
    
    Dim strBuf As String
    Dim i As Long
   
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
    SaveSetting C_TITLE, "ExcelPages", "FolderStr", strBuf
    SaveSetting C_TITLE, "ExcelPages", "chkSubFolder", chkSubFolder.Value
    SaveSetting C_TITLE, "ExcelPages", "chkExcelSheet", chkExcelSheet.Value
    
    Set mMm = Nothing
    
    Unload Me
    
    AppActivate ResultWS.Application.Caption
    execSelectionRowDrawGrid
    
    Set ResultWS = Nothing
    Set ResultWB = Nothing

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

Private Sub UserForm_Initialize()
    
    Dim strBuf As String
    Dim strSearch() As String
    Dim strFolder() As String
    Dim i As Long
    
    mblnRefresh = True
    
    strBuf = GetSetting(C_TITLE, "ExcelPages", "FolderStr", "")
    strFolder = Split(strBuf, vbTab)
    
    For i = LBound(strFolder) To UBound(strFolder)
        cboFolder.AddItem strFolder(i)
    Next
    If cboFolder.ListCount > 0 Then
        cboFolder.ListIndex = 0
    End If

    lblGauge.visible = False
    
    chkExcel.Value = True
    chkWord.Value = True
    chkPoint.Value = True
    
    chkSubFolder.Value = GetSetting(C_TITLE, "ExcelPages", "chkSubFolder", False)
    chkExcelSheet.Value = GetSetting(C_TITLE, "ExcelPages", "chkExcelSheet", False)
    
End Sub

