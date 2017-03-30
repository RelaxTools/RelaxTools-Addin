VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTreeList 
   Caption         =   "フォルダツリー構造取得"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   OleObjectBlob   =   "frmTreeList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTreeList"
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
Private mMm As MacroManager
Private XL As Excel.Application

Private mdblLineWidth As Double

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFolder_Click()

    Dim strFile As String

    'フォルダ名取得
    strFile = rlxSelectFolder()
    
    If Trim(strFile) <> "" Then
        txtFolder.Text = strFile
    End If
    
    
End Sub

Private Sub cmdRun_Click()
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFolder As String
    Dim objFs As Object
    Dim strLine As String
    Dim lngFCnt As Long
    Dim lngFolderCnt As Long
    
    If ActiveCell Is Nothing Then
        MsgBox "アクティブなセルがみつかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    'フォルダ名取得
    strFolder = txtFolder.Text
    If strFolder = "" Then
        MsgBox "ツリー構造を取得するフォルダを入力してください。", vbExclamation, "ツリー構造取得"
        txtFolder.SetFocus
        Exit Sub
    End If
    
    
    If Val(txtLineWidth.Text) < 0.5 Then
        MsgBox "0.5以上を入力してください。", vbExclamation, "ツリー構造取得"
        txtLineWidth.SetFocus
        Exit Sub
    End If
    
    mdblLineWidth = Val(txtLineWidth.Text)
    
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    
    lngRow = ActiveCell.Row
    lngCol = ActiveCell.Column
    
    strLine = ""
    
    Set mMm = New MacroManager
    Set mMm.Form = Me
    
    mMm.Disable
    mMm.DispGuidance "ファイルの数をカウントしています..."
    
    rlxGetFilesCount objFs, strFolder, lngFCnt, chkFileName.Value, True, True
    
    mMm.StartGauge lngFCnt
    
    Dim strPath As String
    If Mid$(strFolder, 2, 1) = ":" Then
        'ドライブ名をUNCに変換
        strPath = rlxDriveToUNC(strFolder)
    Else
        strPath = strFolder
    End If
    
    'フォルダ見出し（開始時）
    Cells(lngRow, lngCol).Value = strPath
    
    'フォルダ指定の場合
    If chkFolder.Value Then
        ActiveSheet.Hyperlinks.Add _
            Anchor:=Cells(lngRow, lngCol), _
            Address:=strPath, _
            TextToDisplay:=strPath
    End If
    
    On Error Resume Next
    
    If optViewSheet.Value Then
        Set XL = New Excel.Application
    End If
    
    lngFolderCnt = 0
    FileDisp objFs, strFolder, lngRow, lngCol, lngCol, strLine, lngFolderCnt
    
    Set mMm = Nothing
    Set objFs = Nothing
    
    If optViewSheet.Value Then
        XL.Quit
        Set XL = Nothing
    End If
    
    Select Case err.Number
    Case 75, 76
        MsgBox "フォルダが存在しません。", vbExclamation, "ツリー構造取得"
        txtFolder.SetFocus
        Exit Sub
    End Select
       
    Unload Me
    MsgBox "処理が完了しました。", vbInformation, C_TITLE

End Sub
Private Sub FileDisp(objFs, ByVal strPath, lngRow, ByVal lngCol, ByVal lngHCol As Long, ByVal strLineParent As String, ByRef lngFolderCnt As Long)

    Dim objfld As Object
    Dim objfl As Object
    Dim objKey As Variant
    Dim objSub As Object
    
    Dim i As Long
    Dim lngFolderCount As Long
    Dim lngCol2 As Long
    
    Dim strLine As String
    Dim colFolders As Object
    Dim colFiles As Object
    Dim v As Variant
    
    '罫線の列幅を２とする。
    Columns(lngCol).ColumnWidth = mdblLineWidth
    Columns(lngCol + 1).ColumnWidth = mdblLineWidth
    
    Set objfld = objFs.GetFolder(strPath)
    
    lngCol2 = lngCol + 2
    lngRow = lngRow + 1
    
    lngFolderCount = objfld.SubFolders.count
    
    Select Case lngFolderCount > 0
        Case 0
            strLine = strLineParent & "　　"
        Case Else
            strLine = strLineParent & "│　"
    End Select
    
    If chkFileName.Value Then
        
        Set colFiles = CreateObject("Scripting.Dictionary")
        
        For Each objfl In objfld.files
            colFiles.Add objfl.Name, objfl
        Next
        
        rlxSortDictionary colFiles
        
        'ファイルの一覧を作成する。
        For Each objKey In colFiles.Keys
            DoEvents
            If mblnCancel Then
                Exit Sub
            End If
            '罫線
            SetTree strLine, lngRow, lngHCol
            
            'ファイル名
            Cells(lngRow, lngCol2).NumberFormatLocal = "@"
            Cells(lngRow, lngCol2).Value = colFiles.Item(objKey).Name
            
            'ハイパーリンク
            'Office プログラム内のハイパーリンクのファイル名でポンド文字を使用できません。(KB202261)
            'http://support.microsoft.com/kb/202261/ja
            'ファイル指定の場合
            If chkFile.Value Then
                ActiveSheet.Hyperlinks.Add _
                    Anchor:=Cells(lngRow, lngCol2), _
                    Address:=rlxAddFileSeparator(strPath) & colFiles.Item(objKey).Name, _
                    TextToDisplay:=colFiles.Item(objKey).Name
            End If
    
            If optViewSheet.Value Then
                If rlxIsExcelFile(colFiles.Item(objKey).Name) Then
                    For Each v In getSheets(colFiles.Item(objKey).Path)
                        lngRow = lngRow + 1
                        Cells(lngRow, lngCol2 + 1).Value = v
                        SetTree strLine, lngRow, lngHCol
                    Next
                End If
            End If
            
            lngRow = lngRow + 1
            lngFolderCnt = lngFolderCnt + 1
            mMm.DisplayGauge lngFolderCnt
        Next
        Set colFiles = Nothing
    End If
    
    '罫線
    SetTree strLine, lngRow, lngHCol
    lngRow = lngRow + 1
    
    'サブフォルダ検索あり
    i = 1
    
    Set colFolders = CreateObject("Scripting.Dictionary")
    
    For Each objSub In objfld.SubFolders
        colFolders.Add objSub.Name, objSub
    Next
    
    rlxSortDictionary colFolders
        
    For Each objKey In colFolders
        DoEvents
        If mblnCancel Then
            Exit Sub
        End If
        '罫線
        Select Case lngFolderCount
            Case i
                SetTree strLineParent & "└─", lngRow, lngHCol
                strLine = strLineParent & "　　"
        
            Case Else
                SetTree strLineParent & "├─", lngRow, lngHCol
                strLine = strLineParent & "│　"
        End Select
        
        'フォルダ見出し
        Cells(lngRow, lngCol2).NumberFormatLocal = "@"
        Cells(lngRow, lngCol2).Value = rlxGetFullpathFromFileName(colFolders.Item(objKey).Path)
        
        'フォルダ指定の場合
        If chkFolder.Value Then
            ActiveSheet.Hyperlinks.Add _
                Anchor:=Cells(lngRow, lngCol2), _
                Address:=colFolders.Item(objKey).Path, _
                TextToDisplay:=rlxGetFullpathFromFileName(colFolders.Item(objKey).Path)
        End If
                
        '自分自身を呼び出す（再帰）
        FileDisp objFs, colFolders.Item(objKey).Path, lngRow, lngCol2, lngHCol, strLine, lngFolderCnt
        
        i = i + 1
        lngFolderCnt = lngFolderCnt + 1
        mMm.DisplayGauge lngFolderCnt
        
    Next
    Set colFolders = Nothing
    
End Sub 'Tree描画
Private Sub SetTree(ByVal strLine As String, ByVal lngRow As Long, ByVal lngCol As Long)

    Dim lngLen As Long
    Dim i As Long

    
    lngLen = Len(strLine)
    
    For i = 1 To lngLen

        Cells(lngRow, lngCol + i - 1).Value = Mid$(strLine, i, 1)
    Next


End Sub

Private Sub spnWidth_SpinDown()
    txtLineWidth.Text = spinDown(txtLineWidth.Text)
End Sub

Private Sub spnWidth_SpinUp()
    txtLineWidth.Text = spinUp(txtLineWidth.Text)
End Sub
Private Function spinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 1
    spinUp = lngValue

End Function

Private Function spinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 1
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDown = lngValue

End Function

Private Sub UserForm_Initialize()
    mblnCancel = False
    lblGauge.visible = False
End Sub

Private Sub UserForm_Terminate()
    mblnCancel = True
End Sub

Private Function getSheets(ByVal strBook As String) As Collection

    Dim WB As Workbook
    Dim WS As Object
    Set getSheets = New Collection
    
    On Error GoTo e
        
    Set WB = XL.Workbooks.Open(FileName:=strBook, ReadOnly:=True)
    
    For Each WS In WB.Sheets
        If WS.visible = xlSheetVisible Then
            getSheets.Add WS.Name, WS.Name
        End If
    Next
    
    WB.Close SaveChanges:=False
e:
    Set WB = Nothing

End Function
