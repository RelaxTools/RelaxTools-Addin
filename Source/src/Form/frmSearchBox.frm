VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchBox 
   Caption         =   "応答メッセージ"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3480
   OleObjectBlob   =   "frmSearchBox.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSearchBox"
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
Private Const C_RET_REPLACE As Long = 1
Private Const C_RET_ALL As Long = 2
Private Const C_RET_NEXT As Long = 4
Private Const C_RET_CANCEL As Long = 8

Private mlngRet As Long


Public Function Start() As Long

    mlngRet = C_RET_CANCEL

    Me.Show vbModal

    Start = mlngRet


End Function

Private Sub cmdAll_Click()
    mlngRet = C_RET_ALL
    Unload Me

End Sub

Private Sub cmdCancel_Click()
    mlngRet = C_RET_CANCEL
    Unload Me

End Sub

Private Sub cmdNext_Click()
    mlngRet = C_RET_NEXT
    Unload Me

End Sub

Private Sub cmdReplace_Click()
    mlngRet = C_RET_REPLACE
    Unload Me
End Sub


