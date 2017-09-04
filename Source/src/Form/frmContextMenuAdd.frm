VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContextMenuAdd 
   Caption         =   "カテゴリー編集"
   ClientHeight    =   1605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "frmContextMenuAdd.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmContextMenuAdd"
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
Private mblnMode As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim i As Long
    
    With frmContextMenu.lstMenu1
        If mblnMode Then
            For i = 0 To .ListCount - 1
                If txtCat.Text = .List(i) Then
                    MsgBox "すでに登録されています。", vbExclamation + vbOKOnly, C_TITLE
                    Exit Sub
                End If
            Next
             .AddItem ""
             .List(.ListCount - 1, 0) = Me.txtCat.Text
             .List(.ListCount - 1, 1) = ""
        Else
            For i = 0 To .ListCount - 1
                If txtCat.Text = .List(i) And i <> .ListIndex Then
                    MsgBox "すでに登録されています。", vbExclamation + vbOKOnly, C_TITLE
                    Exit Sub
                End If
            Next
            .List(.ListIndex, 0) = Me.txtCat.Text
        End If
    End With
    Unload Me
End Sub

Sub Start(ByVal Mode As Boolean)

    mblnMode = Mode
    
    If mblnMode Then
        lblCat.Caption = "カテゴリの追加"
    Else
        lblCat.Caption = "カテゴリの編集"
        Me.txtCat.Text = frmContextMenu.lstMenu1.Text
    End If
    
    Me.Show

End Sub


