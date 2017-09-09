VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "バグ報告"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmReport"
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

Private Sub cmdClip_Click()

    SetClipText txtEdit.Text

End Sub

Private Sub UserForm_Initialize()

    Dim strBuf As String
    

    strBuf = strBuf & "RelaxTools, Excel, Windowsの情報、エラー内容をお知らせください。" & vbCrLf
    strBuf = strBuf & "-----------------------------------------------------------" & vbCrLf
    strBuf = strBuf & getVersionInfo & vbCrLf
    strBuf = strBuf & "-----------------------------------------------------------" & vbCrLf
    strBuf = strBuf & "・実行した機能とエラーの内容" & vbCrLf
    strBuf = strBuf & "・再現方法" & vbCrLf
    strBuf = strBuf & "・再現するための情報（実行したブックの内容、読み取り専用、" & vbCrLf
    strBuf = strBuf & "　ブックの共有､パスワードのかかったブック等） " & vbCrLf
    strBuf = strBuf & "" & vbCrLf
    
    strBuf = strBuf & "◆GitHub Issue(GitHubのアカウントが必要です)：" & vbCrLf
    strBuf = strBuf & "https://github.com/RelaxTools/RelaxTools-Addin/issues" & vbCrLf & vbCrLf
    
    strBuf = strBuf & "◆掲示板：" & vbCrLf
    strBuf = strBuf & "http://software.opensquare.net/relaxtools/bbs/wforum.cgi" & vbCrLf & vbCrLf
    
    strBuf = strBuf & "◆メール(relaxtools@opensquare.net)でも受け付けます。" & vbCrLf

    txtEdit.Text = strBuf
    txtEdit.SelStart = 0

End Sub
