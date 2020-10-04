VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMergeFile 
   Caption         =   "指定フォルダ一括マージ"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   OleObjectBlob   =   "frmMergeFile.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMergeFile"
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
Dim mblnCancel As Boolean
Dim mMm As MacroManager


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
    
    Dim filename As String
    Dim objFs As Object
    
    
    'フォルダ名取得
    filename = txtFolder.Text
    If filename = "" Then
        MsgBox "ファイル一覧を取得するフォルダを入力してください。", vbExclamation, "ファイル一覧取得"
        txtFolder.SetFocus
        Exit Sub
    End If
    
    
    On Error GoTo e
    
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    
    Dim lngFCnt As Long
    
    Set mMm = New MacroManager
    Set mMm.Form = Me
    
    mMm.Disable
    mMm.DispGuidance "ファイルの数をカウントしています..."
    
    rlxGetFilesCount objFs, filename, lngFCnt, True, False, False
    
    mMm.StartGauge lngFCnt
    
    
    Dim lngCount As Long
    Dim objfld As Object
    Dim objfl As Object
    
    Dim WB As Workbook
    Dim motoWB As Workbook
    
    Dim WS As Worksheet
    
    Set objfld = objFs.GetFolder(filename)
    
    Dim blnFirst As Boolean
    
    blnFirst = True
        
    Application.DisplayAlerts = False
    
    'ファイル名取得
    For Each objfl In objfld.files
    
        DoEvents
        If mblnCancel Then
            Exit Sub
        End If
        
        'エクセルブック以外は対象としない
        If InStr(UCase(objfl.Name), ".XLS") = 0 Then
            GoTo pass
        End If
        
        Set WB = Workbooks.Open(filename:=objfl.Path, ReadOnly:=True, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True)
        
        'ブックのリンクがあったら解除
        Dim LK As Variant
        Dim a As Variant

        If chkLink.Value Then
            a = WB.LinkSources(Type:=xlLinkTypeExcelLinks)
    
            If Not IsEmpty(a) Then
                For Each LK In WB.LinkSources(Type:=xlLinkTypeExcelLinks)
                    WB.BreakLink Name:=LK, Type:=xlLinkTypeExcelLinks
                Next
            End If
        End If
        
        For Each WS In WB.Worksheets
        
            '画面レイアウトのリンクを削除
            
            ''チェックされてたらループを抜ける
            If chkLink.Value Then
                Dim ss As Picture
                For Each ss In WS.Pictures
                    ss.Formula = ""
                Next
            End If
        
            If blnFirst Then
                
                WS.Copy
                
                '追加された最後のブックがマージ用ブック
                Set motoWB = Application.Workbooks(Application.Workbooks.Count)
                
                blnFirst = False
            Else
            
                'シートの最後にコピー
                WS.Copy , motoWB.Sheets(motoWB.Worksheets.Count)
            End If
            
            

            
        Next
        
        WB.Saved = True
        WB.Close
pass:
        lngCount = lngCount + 1
        mMm.DisplayGauge lngCount
    Next
    Application.DisplayAlerts = True
    
    
    
    Unload Me
    
    MsgBox "マージしました。", vbOKOnly + vbInformation, C_TITLE
    motoWB.Worksheets(1).Select
    
    Call SaveSetting(C_TITLE, "MergeFiles", "chkLink", chkLink.Value)
    Call SaveSetting(C_TITLE, "MergeFiles", "FolderStr", txtFolder.Text)
    
    Exit Sub
    
    
e:
    Application.DisplayAlerts = True
    MsgBox Err.Description, vbOKOnly + vbCritical, C_TITLE
    Unload Me
    Set objFs = Nothing
    
End Sub

'--------------------------------------------------------------
'　ワークブックのマージ
'--------------------------------------------------------------
Sub mergeWorkBook()

    Dim strWorkPath As String
    Dim WS As Worksheet
    Dim W2 As Worksheet
    Dim motoWB As Workbook
    Dim WB As Workbook
    
    Dim blnFirst As Boolean
    
    On Error GoTo ErrHandle
    
    
    'ワークブックが２未満の場合、処理不要
    If Workbooks.Count < 2 Then
        Exit Sub
    End If
    
    blnFirst = True
    
    For Each WB In Workbooks

        For Each WS In WB.Worksheets
            If blnFirst Then
                WS.Copy
                Set motoWB = Application.Workbooks(Application.Workbooks.Count)
                blnFirst = False
            Else
                WS.Copy , motoWB.Worksheets(motoWB.Worksheets.Count)
            End If
        Next
        
    Next
    Exit Sub
ErrHandle:
    MsgBox "エラーが発生しました。", vbOKOnly, C_TITLE

End Sub


Private Sub UserForm_Initialize()
    lblGauge.visible = False
    mblnCancel = False
    
    chkLink.Value = GetSetting(C_TITLE, "MergeFiles", "chkLink", False)
    txtFolder.Text = GetSetting(C_TITLE, "MergeFiles", "FolderStr", "")
    
End Sub

Private Sub UserForm_Terminate()
    mblnCancel = True
End Sub
'--------------------------------------------------------------
'　ファイル数カウント
'--------------------------------------------------------------
Public Sub rlxGetFilesCount(ByRef objFs As Object, ByVal strPath As String, ByRef lngFCnt As Long, ByVal blnFile As Boolean, ByVal blnFolder As Boolean, ByVal blnSubFolder As Boolean)

    Dim objfld As Object
    Dim objSub As Object

    Set objfld = objFs.GetFolder(strPath)
    
    If blnFile Then
        lngFCnt = lngFCnt + objfld.files.Count
    End If
    
    If blnFolder Then
        lngFCnt = lngFCnt + objfld.SubFolders.Count
    End If
    
        'フォルダ取得あり
    If blnSubFolder Then
        For Each objSub In objfld.SubFolders
            DoEvents
            rlxGetFilesCount objFs, objSub.Path, lngFCnt, blnFile, blnFolder, blnSubFolder
        Next
        
    End If
End Sub
'--------------------------------------------------------------
'　フォルダ選択
'--------------------------------------------------------------
Public Function rlxSelectFolder() As String
 
    Dim objShell As Object
    Dim objPath As Object
    Dim WS As Object
    Dim strFolder As String
    
    Set objShell = CreateObject("Shell.Application")
    Set objPath = objShell.BrowseForFolder(&O0, "フォルダを選んでください", &H1 + &H10, "")
    If Not objPath Is Nothing Then
    
        'なぜか「デスクトップ」のパスが取得できない
        If objPath = "デスクトップ" Then
            Set WS = CreateObject("WScript.Shell")
            rlxSelectFolder = WS.SpecialFolders("Desktop")
        Else
            rlxSelectFolder = objPath.Items.Item.Path
        
        End If
    Else
        rlxSelectFolder = ""
    End If
    
End Function
