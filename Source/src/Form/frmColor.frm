VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmColor 
   Caption         =   "色の選択(2003互換)"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3030
   OleObjectBlob   =   "frmColor.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmColor"
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

Private mColor As Long
Private mResult As VbMsgBoxResult
Public Function Start(ByRef lngColor As Long) As VbMsgBoxResult

    Dim c As control
    
    mColor = lngColor
    
    For Each c In frmColor.Controls
    
        If TypeName(c) = "Image" Then
    
            '色の初期設定
            c.BackColor = ActiveWorkbook.Colors(Val(Right(c.Name, 2)))
        
            If c.BackColor = lngColor Then
                c.SpecialEffect = fmSpecialEffectSunken
            Else
                c.SpecialEffect = fmSpecialEffectRaised
            End If
            
        End If
    Next

    mResult = vbCancel

    Me.Show vbModal

    If mResult = vbOK Then
        lngColor = mColor
    End If

    Start = mResult
    
End Function

Private Sub cmdColorDlg_Click()

    Dim lngColor As Long
    Dim result As VbMsgBoxResult


    lngColor = mColor

    result = rlxGetColorDlg(lngColor)
'    result = frmColor.Start(lngColor)

    If result > 0 Then
    
        mResult = vbOK
        mColor = result
        Unload Me
    
    End If

End Sub

Private Sub Image01_Click()
    Call selectImage(Image01)
End Sub
Private Sub Image02_Click()
    Call selectImage(Image02)
End Sub
Private Sub Image03_Click()
    Call selectImage(Image03)
End Sub
Private Sub Image04_Click()
    Call selectImage(Image04)
End Sub
Private Sub Image05_Click()
    Call selectImage(Image05)
End Sub
Private Sub Image06_Click()
    Call selectImage(Image06)
End Sub
Private Sub Image07_Click()
    Call selectImage(Image07)
End Sub
Private Sub Image08_Click()
    Call selectImage(Image08)
End Sub
Private Sub Image09_Click()
    Call selectImage(Image09)
End Sub
Private Sub Image10_Click()
    Call selectImage(Image10)
End Sub
Private Sub Image11_Click()
    Call selectImage(Image11)
End Sub
Private Sub Image12_Click()
    Call selectImage(Image12)
End Sub
Private Sub Image13_Click()
    Call selectImage(Image13)
End Sub
Private Sub Image14_Click()
    Call selectImage(Image14)
End Sub
Private Sub Image15_Click()
    Call selectImage(Image15)
End Sub
Private Sub Image16_Click()
    Call selectImage(Image16)
End Sub
Private Sub Image33_Click()
    Call selectImage(Image33)
End Sub
Private Sub Image34_Click()
    Call selectImage(Image34)
End Sub
Private Sub Image35_Click()
    Call selectImage(Image35)
End Sub
Private Sub Image36_Click()
    Call selectImage(Image36)
End Sub
Private Sub Image37_Click()
    Call selectImage(Image37)
End Sub
Private Sub Image38_Click()
    Call selectImage(Image38)
End Sub
Private Sub Image39_Click()
    Call selectImage(Image39)
End Sub
Private Sub Image40_Click()
    Call selectImage(Image40)
End Sub
Private Sub Image41_Click()
    Call selectImage(Image41)
End Sub
Private Sub Image42_Click()
    Call selectImage(Image42)
End Sub
Private Sub Image43_Click()
    Call selectImage(Image43)
End Sub
Private Sub Image44_Click()
    Call selectImage(Image44)
End Sub
Private Sub Image45_Click()
    Call selectImage(Image45)
End Sub
Private Sub Image46_Click()
    Call selectImage(Image46)
End Sub
Private Sub Image47_Click()
    Call selectImage(Image47)
End Sub
Private Sub Image48_Click()
    Call selectImage(Image48)
End Sub
Private Sub Image49_Click()
    Call selectImage(Image49)
End Sub
Private Sub Image50_Click()
    Call selectImage(Image50)
End Sub
Private Sub Image51_Click()
    Call selectImage(Image51)
End Sub
Private Sub Image52_Click()
    Call selectImage(Image52)
End Sub
Private Sub Image53_Click()
    Call selectImage(Image53)
End Sub
Private Sub Image54_Click()
    Call selectImage(Image54)
End Sub
Private Sub Image55_Click()
    Call selectImage(Image55)
End Sub
Private Sub Image56_Click()
    Call selectImage(Image56)
End Sub


Private Sub selectImage(ByRef i As image)
    
    mResult = vbOK
    mColor = i.BackColor
    Unload Me

End Sub

