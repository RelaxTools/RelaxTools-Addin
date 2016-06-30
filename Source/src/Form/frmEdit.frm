VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEdit 
   Caption         =   "セルの拡大表示＋編集"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13365
   OleObjectBlob   =   "frmEdit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmEdit"
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
Private mblnArrowKeyFlg As Boolean
'Private mblnValue  As Boolean
Private Sub cmbFont_Change()

    txtEdit.Font.Name = cmbFont.Text

End Sub

Private Sub cmbSize_Change()
    txtEdit.Font.Size = Val(cmbSize.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFormatSql_Click()

    txtEdit.Text = rlxFormatSql(txtEdit.Text)
    txtEdit.SelStart = Len(frmEdit.txtEdit.Text)
    txtEdit.SetFocus
    txtEdit.SelStart = 0
    
End Sub



Private Sub cmdFormatXML_Click()

    txtEdit.Text = FormatXML(txtEdit.Text)
    txtEdit.SelStart = Len(frmEdit.txtEdit.Text)
    txtEdit.SetFocus
    txtEdit.SelStart = 0
    
End Sub

Private Sub cmdOK_Click()
    
    On Error Resume Next
    Err.Clear
    ActiveCell.Formula = Replace(txtEdit.Text, vbCrLf, vbLf)
    
    If Err.Number = 0 Then
        Unload Me
    Else
        MsgBox "式の設定に失敗しました。式が正しくない可能性があります。", vbOKOnly + vbExclamation, C_TITLE
    End If

End Sub

Private Sub cmdReload_Click()

    On Error GoTo e
    Err.Clear
    ActiveCell.Formula = Replace(txtEdit.Text, vbCrLf, vbLf)
    
    txtEdit.Text = Replace(Replace(ActiveCell.Formula, vbCrLf, vbLf), vbLf, vbCrLf)
    txtValue.Text = Replace(Replace(ActiveCell.value, vbCrLf, vbLf), vbLf, vbCrLf)
    
'    optValue.Value = True
    
    Exit Sub
e:
    MsgBox "式の設定に失敗しました。式が正しくない可能性があります。", vbOKOnly + vbExclamation, C_TITLE
    txtValue.Text = C_ERROR
End Sub

Private Sub cmdUTF8_Click()
    
    Dim bytBuf() As Byte
    Dim utf8 As UTF8Encoding
    
    Set utf8 = New UTF8Encoding
    
    bytBuf = StrConv(txtEdit.Text, vbFromUnicode)
    txtEdit.Text = utf8.GetString(bytBuf)

    Set utf8 = Nothing

End Sub



Private Sub optFormura_Click()
    Call changeValue
End Sub

Private Sub optValue_Click()
    Call changeValue
End Sub

Private Sub txtEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = mblnArrowKeyFlg
End Sub

Private Sub txtEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
     Select Case KeyCode
        Case 37 To 40
            mblnArrowKeyFlg = True
        Case Else
            mblnArrowKeyFlg = False
    End Select

End Sub

Private Sub txtEdit_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mblnArrowKeyFlg = False
End Sub

Private Sub UserForm_Initialize()

    mblnArrowKeyFlg = False
    
    Dim strFont As String
    Dim strSize As String
    
    strFont = GetSetting(C_TITLE, "Edit", "Font", "ＭＳ ゴシック")
    strSize = GetSetting(C_TITLE, "Edit", "Size", "12")
    
    Dim i As Long
    Dim pos As Long
    
    With Application.CommandBars("Formatting").Controls(1)
        For i = 1 To .ListCount
            cmbFont.AddItem .List(i)
            If strFont = .List(i) Then
                pos = i - 1
            End If
        Next i
    End With

    cmbFont.ListIndex = pos
    txtEdit.Font.Name = strFont
    
    cmbSize.AddItem "6"
    cmbSize.AddItem "8"
    cmbSize.AddItem "9"
    cmbSize.AddItem "10"
    cmbSize.AddItem "11"
    cmbSize.AddItem "12"
    cmbSize.AddItem "14"
    cmbSize.AddItem "16"
    cmbSize.AddItem "18"
    cmbSize.AddItem "20"
    cmbSize.AddItem "22"
    cmbSize.AddItem "24"
    cmbSize.AddItem "26"
    cmbSize.AddItem "28"
    cmbSize.AddItem "36"
    cmbSize.AddItem "48"
    cmbSize.AddItem "72"
    
    cmbSize.Text = strSize
    
    txtEdit.Text = String$(100, vbCrLf)
    txtEdit.SelStart = Len(frmEdit.txtEdit.Text)
    
    txtFormura.Text = Replace(Replace(ActiveCell.Formula, vbCrLf, vbLf), vbLf, vbCrLf)
    
    Err.Clear
    On Error Resume Next
    txtValue.Text = Replace(Replace(ActiveCell.value, vbCrLf, vbLf), vbLf, vbCrLf)
    If Err.Number <> 0 Then
        txtValue.Text = C_ERROR
    End If
    
    txtEdit.Text = txtFormura.Text
    txtEdit.SelStart = 0
    
    optFormura.value = True
    
End Sub

Private Sub UserForm_Terminate()

    Dim strSize As String

    SaveSetting C_TITLE, "Edit", "Font", cmbFont.Text
    
    strSize = cmbSize.Text
    If Val(strSize) = 0 Then
        txtEdit.Font.Size = 12
    Else
        txtEdit.Font.Size = Val(strSize)
    End If
    SaveSetting C_TITLE, "Edit", "Size", strSize

End Sub
Private Sub changeValue()
    Dim r As Range
    
    If optValue.value Then
        txtEdit.BackColor = &H8000000F
        txtFormura.Text = txtEdit.Text
        txtEdit.Text = txtValue.Text
'        txtEdit.Locked = True
'        cmdFormatSql.enabled = False
        cmdReload.enabled = False
        cmdOk.enabled = False
    Else
        txtEdit.BackColor = vbWhite
        txtEdit.Text = txtFormura.Text
'        txtEdit.Locked = False
'        cmdFormatSql.enabled = True
        cmdReload.enabled = True
        cmdOk.enabled = True
    End If
    
'    txtEdit.SetFocus
'    SendKeys "^A"
    txtEdit.SelStart = Len(frmEdit.txtEdit.Text)
    txtEdit.SetFocus
    txtEdit.SelStart = 0
'    txtEdit.SelLength = Len(frmEdit.txtEdit.Text)

End Sub
