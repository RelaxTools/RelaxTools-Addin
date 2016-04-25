VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDoc 
   Caption         =   "Excel方眼紙　ユーザ設定"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   OleObjectBlob   =   "frmDoc.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDoc"
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

Private Sub chkSize_Click()
    txtCol.enabled = chkSize.value
'    txtRow.enabled = chkSize.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    If Not IsNumeric(txtFont.value) Then
        MsgBox "フォントサイズに数値を入力してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    If chkSize.value Then
        If Not IsNumeric(txtCol.value) Then
            MsgBox "列の幅に数値を入力してください。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
'        If Not IsNumeric(txtRow.Value) Then
'            MsgBox "行の高さに数値を入力してください。", vbOKOnly + vbExclamation, C_TITLE
'            Exit Sub
'        End If
    End If

    SaveSetting C_TITLE, "FormatCell", "Size", chkSize.value
    SaveSetting C_TITLE, "FormatCell", "Bunrui", optBunrui1.value
    SaveSetting C_TITLE, "FormatCell", "Font", cmbFont.Text
    SaveSetting C_TITLE, "FormatCell", "Point", txtFont.value
    SaveSetting C_TITLE, "FormatCell", "Col", txtCol.value
'    SaveSetting C_TITLE, "FormatCell", "Row", txtRow.Value
    Unload Me

End Sub

Private Sub txtCol_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), Asc(".")
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtFont_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case Asc("0") To Asc("9"), Asc(".")
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtRow_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), Asc(".")
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()

    Dim strFont As String
    Dim pos As Long
    Dim i As Long
    Dim blnNormal As Boolean
    
    If GetSetting(C_TITLE, "FormatCell", "Bunrui", True) Then
        optBunrui1.value = True
        optBunrui2.value = False
    Else
        optBunrui1.value = False
        optBunrui2.value = True
    End If
    strFont = GetSetting(C_TITLE, "FormatCell", "Font", "ＭＳ ゴシック")
    txtFont.value = GetSetting(C_TITLE, "FormatCell", "Point", "9")
    txtCol.value = GetSetting(C_TITLE, "FormatCell", "Col", "8.5")
'    txtRow.Value = GetSetting(C_TITLE, "FormatCell", "Row", "11.25")
    chkSize.value = GetSetting(C_TITLE, "FormatCell", "Size", False)
    txtCol.enabled = chkSize.value
'    txtRow.enabled = chkSize.Value

    With Application.CommandBars("Formatting").Controls(1)
        For i = 1 To .ListCount
            cmbFont.AddItem .List(i)
            If strFont = .List(i) Then
                pos = i - 1
            End If
        Next i
    End With
    cmbFont.ListIndex = pos
    
End Sub
