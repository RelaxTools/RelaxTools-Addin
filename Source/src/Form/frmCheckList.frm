VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCheckList 
   Caption         =   "逆Ｌ罫線設定"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5340
   OleObjectBlob   =   "frmCheckList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCheckList"
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

Private colBorder As New Collection


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim lngPos As Long

    lngPos = getBorder()

    Call SaveSetting(C_TITLE, "CheckList", "LineNo", lngPos)

    Call SaveSetting(C_TITLE, "CheckList", "Head", chkHead.Value)
    Call SaveSetting(C_TITLE, "CheckList", "Bottom", chkBottom.Value)
    Call SaveSetting(C_TITLE, "CheckList", "Left", chkLeft.Value)
    Call SaveSetting(C_TITLE, "CheckList", "Right", chkRight.Value)
    
    Unload Me

End Sub

Private Sub lbl01_Click()
    setBorder 1
End Sub
Private Sub lbl02_Click()
    setBorder 2
End Sub
Private Sub lbl03_Click()
    setBorder 3
End Sub
Private Sub lbl04_Click()
    setBorder 4
End Sub
Private Sub lbl05_Click()
    setBorder 5
End Sub
Private Sub lbl06_Click()
    setBorder 6
End Sub
Private Sub lbl07_Click()
    setBorder 7
End Sub
Private Sub lbl08_Click()
    setBorder 8
End Sub
Private Sub lbl09_Click()
    setBorder 9
End Sub

Private Sub lbl10_Click()
    setBorder 10
End Sub

Private Sub lbl11_Click()
    setBorder 11
End Sub
Private Sub lbl12_Click()
    setBorder 12
End Sub
Private Sub lbl13_Click()
    setBorder 13
End Sub
Private Sub lbl14_Click()
    setBorder 14
End Sub
Function getBorder() As Long
    
    Dim i As Long
    
    getBorder = 2
    
    For i = 1 To colBorder.Count
        If colBorder(i).BorderStyle = fmBorderStyleSingle Then
            getBorder = i
            Exit Function
        End If
    Next

End Function


Sub setBorder(ByVal lngNo As Long)

    Dim i As Long
    
    For i = 1 To colBorder.Count

        If i = lngNo Then
            colBorder(i).BorderStyle = fmBorderStyleSingle
        Else
            colBorder(i).BorderStyle = fmBorderStyleNone
        End If

    Next

End Sub

Private Sub UserForm_Initialize()

    colBorder.Add lbl01
    colBorder.Add lbl02
    colBorder.Add lbl03
    colBorder.Add lbl04
    colBorder.Add lbl05
    colBorder.Add lbl06
    colBorder.Add lbl07
    colBorder.Add lbl08
    colBorder.Add lbl09
    colBorder.Add lbl10
    colBorder.Add lbl11
    colBorder.Add lbl12
    colBorder.Add lbl13
    colBorder.Add lbl14
    
    Dim lngPos As Long
    
    lngPos = Val(GetSetting(C_TITLE, "CheckList", "LineNo", 2))

    chkHead.Value = CBool(GetSetting(C_TITLE, "CheckList", "Head", True))
    chkBottom.Value = CBool(GetSetting(C_TITLE, "CheckList", "Bottom", True))
    chkLeft.Value = CBool(GetSetting(C_TITLE, "CheckList", "Left", True))
    chkRight.Value = CBool(GetSetting(C_TITLE, "CheckList", "Right", True))
    
    setBorder lngPos

End Sub
