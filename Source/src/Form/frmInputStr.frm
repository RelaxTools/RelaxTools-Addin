VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputStr 
   Caption         =   "文字列の追加"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6150
   OleObjectBlob   =   "frmInputStr.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmInputStr"
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

Private Sub cmdOk_Click()

    Dim strBuf As String
    Dim i As Long
    Dim lngCount As Long
    Dim strSearch() As String
    
    strBuf = txtInput.Text
    lngCount = 1
    For i = 0 To txtInput.ListCount - 1
        If txtInput.List(i) <> txtInput.Text Then
            strBuf = strBuf & vbTab & txtInput.List(i)
            lngCount = lngCount + 1
            'リストは最大１０
            If lngCount >= 10 Then
                Exit For
            End If
        End If
    Next
    SaveSetting C_TITLE, "Insert", "InputStr", strBuf

    txtInput.Clear
    strSearch = Split(strBuf, vbTab)
    
    For i = LBound(strSearch) To UBound(strSearch)
        txtInput.AddItem strSearch(i)
    Next
    If txtInput.ListCount > 0 Then
        txtInput.ListIndex = 0
    End If
    
    mResult = vbOK
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()

    Dim strBuf As String
    Dim strSearch() As String
    Dim strReplace() As String
    Dim i As Long

    strBuf = GetSetting(C_TITLE, "Insert", "InputStr", "")
    strSearch = Split(strBuf, vbTab)
    
    For i = LBound(strSearch) To UBound(strSearch)
        txtInput.AddItem strSearch(i)
    Next
    If txtInput.ListCount > 0 Then
        txtInput.ListIndex = 0
    End If

    With txtInput
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Public Function Start() As String


    Me.show vbModal
    If mResult = vbOK Then
    
        Dim strBuf As String
        
        strBuf = txtInput.Text
        
        Select Case True
            Case InStr(strBuf, "\\n") > 0
                strBuf = Replace(strBuf, "\\n", "\n")
                
            Case InStr(strBuf, "\n") > 0
                strBuf = Replace(strBuf, "\n", vbLf)
        
        End Select
        
        Start = strBuf
    Else
        Start = ""
    End If

End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

