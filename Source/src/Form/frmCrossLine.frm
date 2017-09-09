VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCrossLine 
   Caption         =   "十字カーソル設定"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
   OleObjectBlob   =   "frmCrossLine.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCrossLine"
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInit_Click()

    lblEven.BackColor = &H8000&
    txtCol.Value = "2"
    
    lblFont.BackColor = &H8000&
    
End Sub

Private Sub cmdOk_Click()

    Dim strFillVisible As String
    Dim strFillColor As String
    Dim strFillTransparency As String
    Dim strLineVisible As String
    Dim strLineColor As String
    Dim strLineWeight As String
    Dim lngType As Long
    Dim blnGuid As Boolean
    Dim blnEdit As Boolean
    Dim strFontColor As String
    Dim blnLineWidth As Boolean
    
    
    Select Case Val(txtCol.Value)
        Case 0.25! To 100
        Case Else
            MsgBox "線の幅は0.25～10を入力してください。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
    End Select
    
    Select Case True
        Case optHolizon.Value
            lngType = C_HOLIZON
        Case optVertical.Value
            lngType = C_VERTICAL
        Case Else
            lngType = C_ALL
    End Select
    
    strLineColor = "&H" & Right$("00000000" & Hex(lblEven.BackColor), 8)
    strLineWeight = txtCol.Value
    
    blnGuid = chkGuid.Value
    
    strFontColor = "&H" & Right$("00000000" & Hex(lblFont.BackColor), 8)
    
    Call setCrossLineSetting(lngType, strLineColor, strLineWeight, blnGuid, strFontColor)

    Unload Me
End Sub


Private Sub CommandButton1_Click()

End Sub

Private Sub lblFont_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult


    lngColor = lblFont.BackColor

    Result = frmColor.Start(lngColor)

    If Result = vbOK Then
        lblFont.BackColor = lngColor
    End If
    
End Sub

Private Sub lblEven_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult

    lngColor = lblEven.BackColor

    Result = frmColor.Start(lngColor)

    If Result = vbOK Then
        lblEven.BackColor = lngColor
    End If

End Sub

Private Sub spnCol_SpinDown()
    txtCol.Text = spinDown2(txtCol.Text)
End Sub

Private Sub spnCol_SpinUp()
    txtCol.Text = spinUp2(txtCol.Text)
End Sub

Private Function spinUp2(ByVal vntValue As Variant) As Variant

    Dim lngValue As Single

    lngValue = Val(vntValue)
    lngValue = lngValue + 0.25!
    If lngValue > 10 Then
        lngValue = 10
    End If
    spinUp2 = lngValue

End Function

Private Function spinDown2(ByVal vntValue As Variant) As Variant

    Dim lngValue As Single

    lngValue = Val(vntValue)
    lngValue = lngValue - 0.25!
    If lngValue < 0.25! Then
        lngValue = 0.25!
    End If
    spinDown2 = lngValue

End Function

Private Sub UserForm_Initialize()

    Dim lngLineColor As Long
    Dim lngFontColor As Long
    Dim sngLineWeight As Single
    Dim lngType As Long
    Dim blnGuid As Boolean

    Call getCrossLineSetting(lngType, lngLineColor, sngLineWeight, blnGuid, lngFontColor)
    
    Select Case lngType
        Case C_HOLIZON
            optHolizon.Value = True
        Case C_VERTICAL
            optVertical.Value = True
        Case Else
            optAll.Value = True
    End Select

    
    chkGuid.Value = blnGuid
    
    lblEven.BackColor = lngLineColor
    
    txtCol.Value = sngLineWeight
    
    lblFont.BackColor = lngFontColor

    
End Sub
