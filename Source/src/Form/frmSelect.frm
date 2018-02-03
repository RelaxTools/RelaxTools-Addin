VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelect 
   Caption         =   "セルの再選択"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   OleObjectBlob   =   "frmSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelect"
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

Private ret As VBA.VbMsgBoxResult

Public Function Start(ByRef lngRowSel As Long, _
                      ByRef lngRowGap As Long, _
                      ByRef lngColSel As Long, _
                      ByRef lngColGap As Long) As VBA.VbMsgBoxResult
                      

    txtRowSel.Text = lngRowSel
    txtRowGap.Text = lngRowGap
    txtColSel.Text = lngColSel
    txtColGap.Text = lngColGap
    
    ret = vbCancel

    Me.show vbModal

    lngRowSel = txtRowSel.Text
    lngRowGap = txtRowGap.Text
    lngColSel = txtColSel.Text
    lngColGap = txtColGap.Text

    Start = ret

End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ret = vbOK
    Unload Me
End Sub
Private Sub spnColGap_SpinDown()
    txtColGap.Text = spinDown(txtColGap.Text)
End Sub

Private Sub spnColGap_SpinUp()
    txtColGap.Text = spinUp(txtColGap.Text)
End Sub

Private Sub spnColSel_SpinDown()
    txtColSel.Text = spinDown(txtColSel.Text)
End Sub

Private Sub spnColSel_SpinUp()
    txtColSel.Text = spinUp(txtColSel.Text)
End Sub

Private Sub spnRowGap_SpinDown()
    txtRowGap.Text = spinDown(txtRowGap.Text)
End Sub
Private Sub spnRowGap_SpinUp()
    txtRowGap.Text = spinUp(txtRowGap.Text)
End Sub
Private Sub spnRowSel_SpinDown()
    txtRowSel.Text = spinDown(txtRowSel.Text)
End Sub
Private Sub spnRowSel_SpinUp()
    txtRowSel.Text = spinUp(txtRowSel.Text)
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
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDown = lngValue

End Function


