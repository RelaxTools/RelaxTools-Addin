VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateFolder 
   Caption         =   "フォルダ作成"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7155
   OleObjectBlob   =   "frmCreateFolder.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCreateFolder"
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
    Dim strFolder As String
    
    strFolder = rlxSelectFolder()
    If strFolder = "" Then
    Else
        txtFolder.Text = strFolder
    End If
    
End Sub

Private Sub cmdRun_Click()

    Dim strFolder As String
    
    strFolder = txtFolder.Text

    If Not rlxIsFolderExists(strFolder) Then
        MsgBox "指定されたフォルダは存在しません。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    
    Dim obj As SelectionCreateFolder
    
    Set obj = New SelectionCreateFolder
    
    obj.Folder = strFolder
    obj.Run
    
    Set obj = Nothing
    
    Unload Me

End Sub
