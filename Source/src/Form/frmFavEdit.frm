VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFavEdit 
   Caption         =   "お気に入り編集"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7152
   OleObjectBlob   =   "frmFavEdit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFavEdit"
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

Private Sub cmdCancel_Click()

    mResult = vbCancel
    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    If rlxIsFolderExists(txtFile.Text) Or rlxIsFileExists(txtFile.Text) Then
    Else
        MsgBox "ファイルまたはフォルダが存在しません。", vbExclamation + vbOKOnly, C_TITLE
        Exit Sub
    End If

    mResult = vbOK
    Unload Me
    
End Sub

Function Start(ByVal lngMode As Long, ByRef strFile As String) As VbMsgBoxResult

    Select Case lngMode
        Case C_FAVORITE_ADD
            txtFile.Text = ""
        Case C_FAVORITE_MOD
            txtFile.Text = strFile
    End Select
   
    Me.Show
    
    If mResult = vbOK Then
        strFile = txtFile.Text
    End If
    Start = mResult

End Function


Private Sub UserForm_Activate()
'    Call AllwaysOnTop
End Sub

