VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReplace 
   Caption         =   "置換フォルダ選択"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   OleObjectBlob   =   "frmReplace.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmReplace"
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

Private mResult As VbMsgBoxResult

Public Function Start(ByRef strFolder As String, ByRef blnSubFolder As Boolean) As VbMsgBoxResult

    'デフォルト値のセット
    txtFolder.Text = ""
    chkSubFolder.Value = False
    mResult = vbCancel
    
    'フォームの表示
    Me.show vbModal

    '結果の設定
    strFolder = txtFolder.Text
    blnSubFolder = chkSubFolder.Value
    Start = mResult
    

End Function
Private Sub cmdCancel_Click()
    
    mResult = vbCancel
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
    
'    Dim lngRow As Long
'    Dim lngCol As Long
    Dim filename As String
'    Dim objFs As Object
    
'    If ActiveCell Is Nothing Then
'        MsgBox "アクティブなセルがみつかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
    
    'フォルダ名取得
    filename = txtFolder.Text
    If filename = "" Then
        MsgBox "フォルダを入力してください。", vbExclamation, C_TITLE
        txtFolder.SetFocus
        Exit Sub
    End If
    
'    Set objFs = CreateObject("Scripting.FileSystemObject")
'
'    lngRow = ActiveCell.row
'    lngCol = ActiveCell.Column
'
'    On Error Resume Next
'    FileDisp objFs, FileName, lngRow, lngCol
'    Select Case Err.Number
'    Case 75, 76
'        MsgBox "フォルダが存在しません。", vbExclamation, C_TITLE
'        txtFolder.SetFocus
'        Exit Sub
'    End Select
    
    mResult = vbOK
    
    Unload Me
    
End Sub
'Private Sub FileDisp(objFs, strPath, lngRow, lngCol)
'
'    Dim objFld As Object
'    Dim objFl As Object
'    Dim objSub As Object
'
'    Dim lngCol2 As Long
'
'    Set objFld = objFs.GetFolder(strPath)
'
'    For Each objFl In objFld.Files
'        lngCol2 = lngCol
'        If chkFile.Value Then
'            Cells(lngRow, lngCol2) = objFl.Name
'            lngCol2 = lngCol2 + 1
'        End If
'        If chkFolder.Value Then
'            Cells(lngRow, lngCol2) = objFl.ParentFolder.Path
'            lngCol2 = lngCol2 + 1
'        End If
'        If chkFileSize.Value Then
'            Cells(lngRow, lngCol2) = Format(objFl.Size, "#,##0")
'            lngCol2 = lngCol2 + 1
'        End If
'        If chkDate.Value Then
'            Cells(lngRow, lngCol2) = Format(objFl.DateLastModified, "yyyy/mm/dd hh:mm:ss")
'            lngCol2 = lngCol2 + 1
'        End If
'        lngRow = lngRow + 1
'    Next
'
'    'サブフォルダ検索あり
'    If chkSubFolder.Value Then
'        For Each objSub In objFld.SubFolders
'            FileDisp objFs, objSub.Path, lngRow, lngCol
'        Next
'    End If
'
'End Sub

