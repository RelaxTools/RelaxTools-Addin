VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSushiSetting 
   Caption         =   "スシを流す設定"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3525
   OleObjectBlob   =   "frmSushiSetting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSushiSetting"
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
Dim mblnSpin As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim c As Object
    Dim strBuf As String
    
    SaveSetting C_TITLE, "Sushi", "Speed", scrSpeed.Value
    SaveSetting C_TITLE, "Sushi", "Interval", scrInterval.Value
    
'    strBuf = ""
'    For Each c In Controls
'        If TypeName(c) = "CheckBox" Then
'            If c.Value Then
'                strBuf = strBuf & c.tag
'            End If
'        End If
'    Next
'
'    If strBuf = "" Then
'        MsgBox "せめて１つぐらいは指定してね", vbExclamation + vbOKOnly, C_TITLE
'        Exit Sub
'    End If
'
'    SaveSetting C_TITLE, "Sushi", "Show", strBuf
    Unload Me
    
End Sub

Private Sub scrInterval_Change()
    txtInterval.Value = scrInterval.Value
End Sub

Private Sub scrSpeed_Change()
    txtSpeed.Value = scrSpeed.Value
End Sub

Private Sub UserForm_Initialize()

    scrSpeed.Value = GetSetting(C_TITLE, "Sushi", "Speed", 8)
    txtSpeed.Value = scrSpeed.Value
    
    scrInterval.Value = GetSetting(C_TITLE, "Sushi", "Interval", 10)
    txtInterval.Value = scrInterval.Value
    
'    Dim strBuf As String
'
'    strBuf = GetSetting(C_TITLE, "Sushi", "Show", "1")
'
'    Dim c As Object
'    Dim i As Long
'
'    For i = 1 To Len(strBuf)
'        For Each c In Controls
'            If TypeName(c) = "CheckBox" Then
'                If Mid(strBuf, i, 1) = c.tag Then
'                    c.Value = True
'                End If
'            End If
'        Next
'    Next
    
End Sub
