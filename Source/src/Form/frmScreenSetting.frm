VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScreenSetting 
   Caption         =   "Excelスクショモード設定"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8400.001
   OleObjectBlob   =   "frmScreenSetting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmScreenSetting"
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

Private Sub chkPageBreakEnable_Change()

    Dim c As control

    For Each c In Controls
        If c.Tag = "P" Then
        
            c.enabled = chkPageBreakEnable.Value
        
        End If
    
    Next
End Sub

Private Sub chkZoomEnable_Change()

    Dim c As control

    For Each c In Controls
        If c.Tag = "Z" Then
        
            c.enabled = chkZoomEnable.Value
        
        End If
    
    Next
    
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim blnZoomEnable As Boolean
    Dim lngZoomNum As Long
    Dim blnSave As Boolean
    Dim lngBlankNum As Long
    Dim blnPageBreakEnable As Boolean
    Dim lngPageBreakNun As Long
    
    If chkZoomEnable.Value Then
        If IsNumeric(txtZoomNum.Text) Then
        Else
            MsgBox "画像の縮小率には数値を入力してください", vbOKOnly + vbExclamation, C_TITLE
            txtZoomNum.SetFocus
            Exit Sub
        End If
        Select Case Val(txtZoomNum.Text)
            Case 10 To 200
            Case Else
                MsgBox "画像の縮小率には10～200%を入力してください", vbOKOnly + vbExclamation, C_TITLE
                txtZoomNum.SetFocus
                Exit Sub
        End Select
    Else
        txtZoomNum.Text = "100"
    End If
    
    If IsNumeric(txtBlankNum.Text) Then
    Else
        MsgBox "画像の間隔には数値を入力してください", vbOKOnly + vbExclamation, C_TITLE
        txtBlankNum.SetFocus
        Exit Sub
    End If
    Select Case Val(txtBlankNum.Text)
        Case 0 To 99
        Case Else
            MsgBox "行数の間隔には0～99を入力してください", vbOKOnly + vbExclamation, C_TITLE
            txtBlankNum.SetFocus
            Exit Sub
    End Select
    
    If chkPageBreakEnable.Value Then
        If IsNumeric(txtPageBreakNum.Text) Then
        Else
            MsgBox "改ページの間隔には数値を入力してください", vbOKOnly + vbExclamation, C_TITLE
            txtPageBreakNum.SetFocus
            Exit Sub
        End If
        Select Case Val(txtPageBreakNum.Text)
            Case 0 To 99
            Case Else
                MsgBox "改ページの間隔には0～99を入力してください", vbOKOnly + vbExclamation, C_TITLE
                txtPageBreakNum.SetFocus
                Exit Sub
        End Select
    Else
        txtPageBreakNum.Text = "1"
    End If
    
    blnZoomEnable = chkZoomEnable.Value
    lngZoomNum = Val(txtZoomNum.Value)
    blnSave = chkSave.Value
    lngBlankNum = Val(txtBlankNum.Value)
    blnPageBreakEnable = chkPageBreakEnable.Value
    lngPageBreakNun = Val(txtPageBreakNum.Value)
    
    SetScreenSetting blnZoomEnable, lngZoomNum, blnSave, lngBlankNum, blnPageBreakEnable, lngPageBreakNun
    
    Unload Me
    
End Sub

Private Sub spnBlankNum_SpinDown()
    txtBlankNum.Text = LineSpinDown(txtBlankNum.Text)
End Sub

Private Sub spnBlankNum_SpinUp()
    txtBlankNum.Text = LineSpinUp(txtBlankNum.Text)
End Sub
Private Sub spnPageBreakNum_SpinDown()
    txtPageBreakNum.Text = LineSpinDown(txtPageBreakNum.Text)
End Sub

Private Sub spnPageBreakNum_SpinUp()
    txtPageBreakNum.Text = LineSpinUp(txtPageBreakNum.Text)
End Sub

Private Sub spnZoomNum_SpinDown()
    txtZoomNum.Text = ZoomSpinDown(txtZoomNum.Text)
End Sub

Private Sub spnZoomNum_SpinUp()
    txtZoomNum.Text = ZoomSpinUp(txtZoomNum.Text)
End Sub

Private Sub UserForm_Initialize()

    Dim blnZoomEnable As Boolean
    Dim lngZoomNum As Long
    Dim blnSave As Boolean
    Dim lngBlankNum As Long
    Dim blnPageBreakEnable As Boolean
    Dim lngPageBreakNun As Long

    GetScreenSetting blnZoomEnable, lngZoomNum, blnSave, lngBlankNum, blnPageBreakEnable, lngPageBreakNun
    
    txtZoomNum.Text = lngZoomNum
    txtBlankNum.Text = lngBlankNum
    txtPageBreakNum.Text = lngPageBreakNun
    
    chkZoomEnable.Value = blnZoomEnable
    chkSave.Value = blnSave
    chkPageBreakEnable.Value = blnPageBreakEnable
    
    Call chkZoomEnable_Change
    Call chkPageBreakEnable_Change

End Sub
Private Function LineSpinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 1
    
    If lngValue > 99 Then
        lngValue = 99
    End If
    
    LineSpinUp = lngValue

End Function

Private Function LineSpinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 1
    
    If lngValue < 0 Then
        lngValue = 0
    End If
    
    LineSpinDown = lngValue

End Function
Private Function ZoomSpinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 1
    
    If lngValue > 200 Then
        lngValue = 200
    End If
    
    ZoomSpinUp = lngValue

End Function

Private Function ZoomSpinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 1
    
    If lngValue < 10 Then
        lngValue = 10
    End If
    
    ZoomSpinDown = lngValue

End Function
