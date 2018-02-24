VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimeLeapSetting 
   Caption         =   "TimeLeap設定"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   OleObjectBlob   =   "frmTimeLeapSetting.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTimeLeapSetting"
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





Private Sub cmdSave_Click()

    Dim strList As String
    Dim i As Long
    
    If Len(txtFolder.Text) <> 0 Then
        If Not rlxIsFolderExists(txtFolder.Text) Then
            MsgBox "指定されたフォルダは存在しません。", vbOKOnly + vbExclamation, C_TITLE
            txtFolder.SetFocus
            Exit Sub
        End If
    End If
    
    Select Case Val(txtGen.Text)
        Case 1 To 99
        Case Else
            MsgBox "世代数には1～99を入力してください。", vbOKOnly + vbExclamation, C_TITLE
            txtGen.SetFocus
            Exit Sub
    End Select
    
    strList = txtDiff.Text

    SaveSetting C_TITLE, "TimeLeap", "Diff", strList
    SaveSetting C_TITLE, "TimeLeap", "Folder", txtFolder.Text
    SaveSetting C_TITLE, "TimeLeap", "Gen", txtGen.Text
    
    Unload Me
    
End Sub



Private Sub UserForm_Initialize()
        
    Dim strBuf As String
    
    strBuf = "設定例:" & vbCrLf & vbCrLf
    strBuf = strBuf & "　$(SRC)  ・・・比較元ブックに内部で置換されます。" & vbCrLf
    strBuf = strBuf & "　$(DEST) ・・・比較先ブックに内部で置換されます。" & vbCrLf & vbCrLf
    strBuf = strBuf & "◇WinMerge の場合" & vbCrLf
    strBuf = strBuf & " ""(WinMergeのインストールパス)\WinMergeU.exe"" $(SRC) $(DEST)" & vbCrLf
    strBuf = strBuf & "◇ExcelMerge の場合" & vbCrLf
    strBuf = strBuf & " ""(ExcelMergeのインストールパス)\ExcelMerge.GUI.exe"" diff -s $(SRC) -d $(DEST)" & vbCrLf & vbCrLf
        
    txtExample.Text = strBuf
        
    txtGen.Text = GetSetting(C_TITLE, "TimeLeap", "Gen", "99")
    txtFolder.Text = GetSetting(C_TITLE, "TimeLeap", "Folder", GetTimeLeapFolder())
    txtDiff.Text = GetSetting(C_TITLE, "TimeLeap", "Diff", "")

End Sub



