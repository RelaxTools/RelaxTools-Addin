VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMessageDigest 
   Caption         =   "メッセージダイジェスト生成"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   OleObjectBlob   =   "frmMessageDigest.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMessageDigest"
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


    strFile = Application.GetOpenFilename("ファイル(*.*),(*.*)", , "ファイル読込", , False)
    If strFile = "False" Then
        'ファイル名が指定されなかった場合
        Exit Sub
    End If
    
    txtFolder.Text = strFile
    
End Sub

Private Sub cmdRun_Click()

    Dim fp As Integer
    Dim strFile As String
    Dim strBuf As String
    Dim bytBuf() As Byte
    
    Dim varRow As Variant
    
    Dim r As Range
    
    strFile = txtFolder.Text
    
    'ファイルの存在チェック
    If rlxIsFileExists(strFile) Then
    Else
        MsgBox "ファイルが存在しません。", vbExclamation, C_TITLE
        Exit Sub
    End If
    
    fp = FreeFile()
    Open strFile For Binary As fp
    
    If LOF(fp) <> 0 Then
        ReDim bytBuf(0 To LOF(fp) - 1)
        Get fp, , bytBuf
    End If
    
    Close fp
    
    Dim md As New CryptoServiceProvider
    
    md.HashType = HashTypeMD5
    txtMD5.Text = md.ComputeHash(bytBuf)
    
    md.HashType = HashTypeSHA1
    txtSHA1.Text = md.ComputeHash(bytBuf)
    
    md.HashType = HashTypeSHA256
    txtSHA256.Text = md.ComputeHash(bytBuf)
    
    md.HashType = HashTypeSHA384
    txtSHA384.Text = md.ComputeHash(bytBuf)
    
    md.HashType = HashTypeSHA512
    txtSHA512.Text = md.ComputeHash(bytBuf)
    
    Set md = Nothing

End Sub




Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()

    txtMD5.Text = ""
    txtSHA1.Text = ""
    txtSHA256.Text = ""
    txtSHA384.Text = ""
    txtSHA512.Text = ""

End Sub
