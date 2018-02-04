VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCrossLine 
   Caption         =   "十字カーソル設定"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
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

Private Sub chkFillVisible_Change()
        
    fraFill.enabled = chkFillVisible.Value
    fraFillColor.enabled = chkFillVisible.Value
    fraFillTransparency.enabled = chkFillVisible.Value
    lblClick.enabled = chkFillVisible.Value
    lblPercent.enabled = chkFillVisible.Value
    lblFillColor.enabled = chkFillVisible.Value
    txtFillTransparency.enabled = chkFillVisible.Value
    spnFillTransparency.enabled = chkFillVisible.Value

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInit_Click()

    chkFillVisible.Value = False
    lblFillColor.BackColor = &H50B000
    txtFillTransparency.Value = "50"

    lblEven.BackColor = &H50B000
    txtCol.Value = "2"
    lblFont.BackColor = &H50B000
    lblBack.BackColor = &HFFFFFF
    txtGuidTransparency.Value = "50"
    
End Sub

Private Sub cmdOk_Click()

    Dim blnFillVisible As Boolean
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
    Dim strBackColor As Long
    Dim strGuidTransparency As String
    
    
    Select Case Val(txtFillTransparency.Value)
        Case 0 To 100
        Case Else
            MsgBox "透明度は0～100%を入力してください。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
    End Select
    
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
    
    
    If chkFillVisible.Value Then
        blnFillVisible = False
    Else
        blnFillVisible = True
    End If
    
    strFillColor = "&H" & Right$("00000000" & Hex(lblFillColor.BackColor), 8)
    strFillTransparency = txtFillTransparency.Value
    
    strLineColor = "&H" & Right$("00000000" & Hex(lblEven.BackColor), 8)
    strLineWeight = txtCol.Value
    
    blnGuid = chkGuid.Value
    
    strFontColor = "&H" & Right$("00000000" & Hex(lblFont.BackColor), 8)
    
    strBackColor = "&H" & Right$("00000000" & Hex(lblBack.BackColor), 8)
    
    strGuidTransparency = txtGuidTransparency.Value
    
    Call setCrossLineSetting(lngType, blnFillVisible, strFillColor, strFillTransparency, strLineVisible, strLineColor, strLineWeight, blnGuid, strFontColor, blnEdit, blnLineWidth, strBackColor, strGuidTransparency)


    Unload Me
End Sub


Private Sub CommandButton1_Click()

End Sub

Private Sub lblBack_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult

    lngColor = lblBack.BackColor

    Result = frmColor.Start(lngColor)

    If Result = vbOK Then
        lblBack.BackColor = lngColor
    End If
    
End Sub

Private Sub lblFillColor_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult

    lngColor = lblFillColor.BackColor

    Result = frmColor.Start(lngColor)

    If Result = vbOK Then
        lblFillColor.BackColor = lngColor
    End If
    
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
Private Function spinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Single

    lngValue = Val(vntValue)
    lngValue = lngValue + 5
    If lngValue > 100 Then
        lngValue = 100
    End If
    spinUp = lngValue

End Function

Private Function spinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Single

    lngValue = Val(vntValue)
    lngValue = lngValue - 5
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDown = lngValue

End Function



Private Sub spnFillTransparency_SpinDown()
    txtFillTransparency.Text = spinDown(txtFillTransparency.Text)
End Sub

Private Sub spnFillTransparency_SpinUp()
    txtFillTransparency.Text = spinUp(txtFillTransparency.Text)
End Sub

Private Sub spnGuidTransparency_SpinDown()
    txtGuidTransparency.Text = spinDown(txtGuidTransparency.Text)
End Sub

Private Sub spnGuidTransparency_SpinUp()
    txtGuidTransparency.Text = spinUp(txtGuidTransparency.Text)
End Sub

Private Sub UserForm_Initialize()

    Dim blnFillVisible As Boolean
    Dim lngFillColor As Long
    Dim dblFillTransparency As Double
    Dim lngLineVisible As Long
    Dim lngLineColor As Long
    Dim lngFontColor As Long
    Dim sngLineWeight As Single
    Dim strOnAction As String
    Dim lngType As Long
    Dim blnGuid As Boolean
    Dim blnEdit As Boolean
    Dim blnLineWidth As Boolean
    Dim lngBackColor As Long
    Dim dblGuidTransparency As Double

    Call getCrossLineSetting(lngType, blnFillVisible, lngFillColor, dblFillTransparency, lngLineVisible, lngLineColor, sngLineWeight, strOnAction, blnGuid, lngFontColor, blnEdit, blnLineWidth, lngBackColor, dblGuidTransparency)
    
    
    
    
    Select Case lngType
        Case C_HOLIZON
            optHolizon.Value = True
        Case C_VERTICAL
            optVertical.Value = True
        Case Else
            optAll.Value = True
    End Select

    If blnFillVisible Then
        chkFillVisible.Value = False
    Else
        chkFillVisible.Value = True
    End If
    
    lblFillColor.BackColor = lngFillColor
    txtFillTransparency.Value = dblFillTransparency
    
    chkGuid.Value = blnGuid
    
    lblEven.BackColor = lngLineColor
    
    txtCol.Value = sngLineWeight
    
    lblFont.BackColor = lngFontColor
    lblBack.BackColor = lngBackColor

    txtGuidTransparency.Value = dblGuidTransparency
    
End Sub
