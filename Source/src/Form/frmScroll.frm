VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScroll 
   Caption         =   "マウスのホイールスクロール行数"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3915
   OleObjectBlob   =   "frmScroll.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmScroll"
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

Private Sub cmdOk_Click()

     Call SaveSetting(C_TITLE, "ScrollLine", "ScrollLine", txtScroll.Text)
     Call SaveSetting(C_TITLE, "ScrollLine", "DefaultLine", txtDefault.Text)
     Unload Me
    
End Sub

Private Sub spnDefault_SpinDown()
    txtDefault.Text = spinDown(txtDefault.Text)

End Sub

Private Sub spnDefault_SpinUp()
    txtDefault.Text = spinUp(txtDefault.Text)

End Sub

Private Sub spnScroll_SpinDown()
    txtScroll.Text = spinDown(txtScroll.Text)
End Sub

Private Sub spnScroll_SpinUp()
    txtScroll.Text = spinUp(txtScroll.Text)
End Sub


Private Sub txtDefault_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case Chr$(KeyAscii)
        Case "0" To "9"
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtScroll_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case Chr$(KeyAscii)
        Case "0" To "9"
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()
    
    txtScroll.Text = GetSetting(C_TITLE, "ScrollLine", "ScrollLine", 1)
    txtDefault = GetSetting(C_TITLE, "ScrollLine", "DefaultLine", 3)
    
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
    If lngValue < 1 Then
        lngValue = 1
    End If
    spinDown = lngValue

End Function
