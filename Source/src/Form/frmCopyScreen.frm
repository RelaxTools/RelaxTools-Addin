VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCopyScreen 
   Caption         =   "選択範囲のピクチャ化設定"
   ClientHeight    =   2640
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3648
   OleObjectBlob   =   "frmCopyScreen.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCopyScreen"
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
Private mResult As VBA.VbMsgBoxResult

Private Sub chkBackColor_Change()
    lblHead.enabled = Not (chkBackColor.Value)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()

    Dim blnFillVisible As Boolean
    Dim strFillColor As String
    Dim blnLine As Boolean

    blnFillVisible = chkBackColor.Value
    
    strFillColor = "&H" & Right$("00000000" & Hex(lblHead.BackColor), 8)
    
    blnLine = chkLine.Value
    
    Call setCopyScreenSetting(blnFillVisible, strFillColor, blnLine)

    Unload Me
End Sub
Private Sub lblHead_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult

    lngColor = lblHead.BackColor
    Result = frmColor.Start(lngColor)

    If Result = vbOK Then
        lblHead.BackColor = lngColor
    End If

End Sub

Private Sub UserForm_Initialize()

    Dim blnFillVisible As Boolean
    Dim lngFillColor As Long
    Dim blnLine As Boolean

    Call getCopyScreenSetting(blnFillVisible, lngFillColor, blnLine)
    
    chkBackColor.Value = blnFillVisible
    
    lblHead.BackColor = lngFillColor
    
    chkLine.Value = blnLine
    
End Sub
