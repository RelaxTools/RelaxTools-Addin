VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFusen 
   Caption         =   "付箋の設定"
   ClientHeight    =   8385.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12390
   OleObjectBlob   =   "frmFusen.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFusen"
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

Private Sub cmdHelp_Click()
    
    If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    Dim WSH As Object
    
    Set WSH = CreateObject("WScript.Shell")
    
    Call WSH.Run(C_STAMP_URL)
    
    Set WSH = Nothing

End Sub

Private Sub cmdOk_Click()
    
    Dim strText As String
    Dim strTag As String
    Dim blnPrint As Boolean
    
    Dim strWidth  As String
    Dim strHeight  As String
    
    Dim strFormat As String
    Dim strUserDate  As String
    Dim strFusenDate As String
    
    Dim strFont  As String
    Dim strSize  As String
    
    Dim strHorizontalAnchor  As String
    Dim strVerticalAnchor  As String
    
    Dim blnAutoSize  As Boolean
    Dim blnOverFlow As Boolean
    Dim blnWordWrap As Boolean
    
    
    strText = txtText.Text
    strTag = txtTag.Text
    blnPrint = chkPrint.Value
    
    strWidth = txtWidth.Text
    strHeight = txtHeight.Text
    
    strFormat = txtFormat.Text
    strUserDate = txtUserDate.Text
    
    Select Case True
        Case optSystemDate.Value
            strFusenDate = C_FUSEN_DATE_SYSTEM
        Case Else
            strFusenDate = C_FUSEN_DATE_USER
    End Select
    
    strFont = cboFont.Text
    strSize = txtSize.Text
    
    strHorizontalAnchor = cboHorizontalAnchor.ListIndex
    strVerticalAnchor = cboVerticalAnchor.ListIndex
    
    blnAutoSize = chkAutoSize.Value
    blnOverFlow = chkOverflow.Value
    blnWordWrap = chkWordWrap.Value
    
    If strFusenDate = C_FUSEN_DATE_USER Then
        If IsDate(strUserDate) Then
        Else
            MsgBox "指定日付には有効な日付をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
            txtUserDate.SetFocus
            Exit Sub
        End If
    End If
    
    If IsNumeric(strHeight) Then
    Else
        MsgBox "高さには数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
        txtHeight.SetFocus
        Exit Sub
    End If
    
    If CDbl(strHeight) < 0 Then
        MsgBox "高さは０以上を入力してください。", vbExclamation + vbOKOnly, C_TITLE
        txtHeight.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(strWidth) Then
    Else
        MsgBox "幅には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
        txtWidth.SetFocus
        Exit Sub
    End If
    
    If CDbl(strWidth) < 0 Then
        MsgBox "幅は０以上を入力してください。", vbExclamation + vbOKOnly, C_TITLE
        txtWidth.SetFocus
        Exit Sub
    End If
        
    SaveSetting C_TITLE, "Fusen", "Text", strText
    SaveSetting C_TITLE, "Fusen", "Tag", strTag
    SaveSetting C_TITLE, "Fusen", "PrintObject", blnPrint
    
    SaveSetting C_TITLE, "Fusen", "Width", strWidth
    SaveSetting C_TITLE, "Fusen", "Height", strHeight
    
    SaveSetting C_TITLE, "Fusen", "UserDate", strUserDate
    SaveSetting C_TITLE, "Fusen", "Format", strFormat
    SaveSetting C_TITLE, "Fusen", "FusenDate", strFusenDate
    
    SaveSetting C_TITLE, "Fusen", "Font", strFont
    SaveSetting C_TITLE, "Fusen", "Size", strSize
    
    SaveSetting C_TITLE, "Fusen", "HorizontalAnchor", strHorizontalAnchor
    SaveSetting C_TITLE, "Fusen", "VerticalAnchor", strVerticalAnchor
    
    SaveSetting C_TITLE, "Fusen", "AutoSize", blnAutoSize
    SaveSetting C_TITLE, "Fusen", "OverFlow", blnOverFlow
    SaveSetting C_TITLE, "Fusen", "WordWrap", blnWordWrap
    
    Unload Me
    
End Sub

Private Function spinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 0.5
    spinUp = lngValue

End Function

Private Function spinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 0.5
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDown = lngValue

End Function
Private Function spinUpSize(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 1
    spinUpSize = lngValue

End Function

Private Function spinDownSize(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 1
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDownSize = lngValue

End Function

Private Sub spnHeight_SpinDown()
    txtHeight.Text = spinDown(txtHeight.Text)
End Sub

Private Sub spnHeight_SpinUp()
    txtHeight.Text = spinUp(txtHeight.Text)
End Sub

Private Sub spnSize_SpinDown()
    txtSize.Text = spinDownSize(txtSize.Text)
End Sub

Private Sub spnSize_SpinUp()
    txtSize.Text = spinUpSize(txtSize.Text)
End Sub

Private Sub spnWidth_SpinDown()
    txtWidth.Text = spinDown(txtWidth.Text)
End Sub

Private Sub spnWidth_SpinUp()
    txtWidth.Text = spinUp(txtWidth.Text)
End Sub

Private Sub UserForm_Initialize()

    Dim strText As String
    Dim strTag As String
    Dim varPrint As Variant
    
    Dim strWidth  As String
    Dim strHeight  As String
    
    Dim strFormat As String
    Dim strUserDate  As String
    Dim strFusenDate As String
    
    Dim strFont  As String
    Dim strSize  As String
    
    Dim varHorizontalAnchor  As Variant
    Dim varVerticalAnchor  As Variant
    
    Dim varAutoSize  As Variant
    Dim varOverFlow As Variant
    Dim varWordWrap As Variant
    
    Dim i As Long
    
    Call getSettingFusen(strText, strTag, varPrint, strWidth, strHeight, strFormat, strUserDate, strFusenDate, strFont, strSize, varHorizontalAnchor, varVerticalAnchor, varAutoSize, varOverFlow, varWordWrap)
    
    txtText.Text = strText
    txtTag.Text = strTag
    chkPrint.Value = varPrint
    
    txtWidth.Text = strWidth
    txtHeight.Text = strHeight
    
    txtFormat.Text = strFormat
    
    Select Case strFusenDate
        Case C_FUSEN_DATE_SYSTEM
            optSystemDate.Value = True
        Case Else
            optUserDate.Value = True
    End Select
    
    txtUserDate.Text = strUserDate
    
    ActiveCell.Select
    With Application.CommandBars("Formatting").Controls(1)
        For i = 1 To .ListCount
            cboFont.AddItem .List(i)
        Next i
    End With
    cboFont.Text = strFont
    txtSize.Text = strSize
    
    cboHorizontalAnchor.AddItem "左"
    cboHorizontalAnchor.AddItem "中"
    cboHorizontalAnchor.AddItem "右"
    cboHorizontalAnchor.ListIndex = varHorizontalAnchor
    
    cboVerticalAnchor.AddItem "上"
    cboVerticalAnchor.AddItem "中"
    cboVerticalAnchor.AddItem "下"
    cboVerticalAnchor.ListIndex = varVerticalAnchor
    
    chkAutoSize.Value = varAutoSize
    
    lblUser.Caption = "ユーザ名:" & Application.UserName
    
    chkOverflow.Value = varOverFlow
    chkWordWrap.Value = varWordWrap
    
    
    txtFormat.AddItem "yyyy/mm/dd"
    txtFormat.AddItem "yyyy.mm.dd"
    txtFormat.AddItem "'yy.mm.dd"
    txtFormat.AddItem "ge.m.d"
    txtFormat.AddItem "gge.m.d"
    txtFormat.AddItem "ggge年m月d日"
    
#If VBA7 Then
#Else
    chkOverflow.enabled = False
#End If

End Sub
