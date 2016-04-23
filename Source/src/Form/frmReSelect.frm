VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReSelect 
   Caption         =   "選択範囲の復元"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   OleObjectBlob   =   "frmReSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmReSelect"
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

Dim mstrSelRenge As String
Dim mlngLeft As Long
Dim mlngTop As Long

Private Sub cboRange_Change()
    Dim lngLeft As Long
    Dim lngTop As Long
    
    If chkOffset.value Then
        lngLeft = getLeft(Range(cboRange.Text))
        lngTop = getTop(Range(cboRange.Text))
        
        lngLeft = mlngLeft - lngLeft
        lngTop = mlngTop - lngTop
        
        Range(cboRange.Text).Offset(lngTop, lngLeft).Select
    Else
        '選択範囲を復元
        Range(cboRange.Text).Select
    End If
End Sub

Private Sub chkOffset_Click()
    Call cboRange_Change
End Sub

Private Sub cmdCancel_Click()
    '選択範囲を復元
    Range(mstrSelRenge).Select
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim strBuf As String
    Dim lngCount As Long
    Dim i As Long

    strBuf = cboRange.Text
    If Len(Trim(strBuf)) = 0 Then
        Exit Sub
    End If
    
    lngCount = 1
    For i = 0 To cboRange.ListCount - 1
        If cboRange.List(i) <> cboRange.Text Then
            strBuf = strBuf & vbTab & cboRange.List(i)
            lngCount = lngCount + 1
            'リストは最大１０
            If lngCount >= 20 Then
                Exit For
            End If
        End If
    Next
    SaveSetting C_TITLE, "ReSelect", "Range", strBuf
    SaveSetting C_TITLE, "ReSelect", "Offset", chkOffset.value
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    Dim strRange() As String
    Dim strBuf As String
    Dim i As Long
    
    '現在の選択範囲を保存
    mstrSelRenge = Selection.Address
    
    mlngLeft = getLeft(Selection)
    mlngTop = getTop(Selection)
    

    strBuf = GetSetting(C_TITLE, "ReSelect", "Range", "")
    strRange = Split(strBuf, vbTab)
    
    For i = LBound(strRange) To UBound(strRange)
        cboRange.AddItem strRange(i)
    Next
    If cboRange.ListCount > 0 Then
        cboRange.ListIndex = 0
    End If
    
    chkOffset.value = GetSetting(C_TITLE, "ReSelect", "Offset", False)
    
End Sub
Private Function getLeft(ByVal s As Range) As Long

    Dim lngMin As Long
    Dim r As Range
    lngMin = 2147483647

    For Each r In s.Areas
        If lngMin > r.Column Then
            lngMin = r.Column
        End If
    Next
    getLeft = lngMin
End Function
Private Function getTop(ByVal s As Range) As Long

    Dim lngMin As Long
    Dim r As Range
    
    lngMin = 2147483647

    For Each r In s.Areas
        If lngMin > r.row Then
            lngMin = r.row
        End If
    Next
    getTop = lngMin
End Function

