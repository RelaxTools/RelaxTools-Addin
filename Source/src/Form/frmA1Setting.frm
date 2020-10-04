VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmA1Setting 
   Caption         =   "ホームポジション（A1)設定"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "frmA1Setting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmA1Setting"
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

Private Sub chkRatio_Click()
        
    cboPercent.enabled = chkRatio.Value

End Sub

Private Sub chkView_Click()
    cboView.enabled = chkView.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Select Case Val(cboPercent.Text)
        Case 10 To 400
        Case Else
            MsgBox "倍率は１０～４００％の間で入力してください。", vbOKOnly + vbExclamation, C_TITLE
            cboPercent.SetFocus
            Exit Sub
    End Select

    SaveSetting C_TITLE, "A1Setting", "ratio", chkRatio.Value
    SaveSetting C_TITLE, "A1Setting", "percent", cboPercent.Value
    SaveSetting C_TITLE, "A1Setting", "ViewEnable", chkView.Value
    SaveSetting C_TITLE, "A1Setting", "View", cboView.ListIndex
    Unload Me
    
End Sub


Private Sub UserForm_Initialize()

    chkRatio.Value = GetSetting(C_TITLE, "A1Setting", "ratio", False)
    chkView.Value = GetSetting(C_TITLE, "A1Setting", "ViewEnable", False)
    
    cboPercent.Clear
    cboPercent.AddItem "25"
    cboPercent.AddItem "50"
    cboPercent.AddItem "75"
    cboPercent.AddItem "100"
    cboPercent.AddItem "200"
    cboPercent.AddItem "400"
    
    cboView.AddItem "標準"
    cboView.AddItem "ページレイアウト"
    cboView.AddItem "改ページプレビュー"
    
    cboPercent.Value = GetSetting(C_TITLE, "A1Setting", "percent", "100")
    cboView.ListIndex = Val(GetSetting(C_TITLE, "A1Setting", "View", "0"))
    
    cboPercent.enabled = chkRatio.Value
    cboView.enabled = chkView.Value

End Sub
