VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCopyScreen 
   Caption         =   "選択範囲のピクチャ化設定"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3645
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
    lblHead.enabled = Not (chkBackColor.value)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()

    Dim blnFillVisible As Boolean
    Dim strFillColor As String
    Dim blnLine As Boolean

    blnFillVisible = chkBackColor.value
    
    strFillColor = "&H" & Right$("00000000" & Hex(lblHead.BackColor), 8)
    
    blnLine = chkLine.value
    
    Call setCopyScreenSetting(blnFillVisible, strFillColor, blnLine)

    Unload Me
End Sub
Private Sub lblHead_Click()

    Dim lngColor As Long
    Dim result As VbMsgBoxResult

    lngColor = lblHead.BackColor
    result = frmColor.Start(lngColor)

    If result = vbOK Then
        lblHead.BackColor = lngColor
    End If

End Sub

Private Sub UserForm_Initialize()

    Dim blnFillVisible As Boolean
    Dim lngFillColor As Long
    Dim blnLine As Boolean

    Call getCopyScreenSetting(blnFillVisible, lngFillColor, blnLine)
    
    chkBackColor.value = blnFillVisible
    
    lblHead.BackColor = lngFillColor
    
    chkLine.value = blnLine
    
End Sub
