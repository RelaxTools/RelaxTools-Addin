VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFavCategory 
   Caption         =   "カテゴリー編集"
   ClientHeight    =   1605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "frmFavCategory.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFavCategory"
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
Private mlngMode As Long


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim i As Long
    Dim fav As favoriteDTO
    Dim c As Variant
    
    With frmFavorite.lstCategory
        Select Case mlngMode
            Case C_FAVORITE_ADD
                For i = 0 To .ListCount - 1
                    If txtCat.Text = .List(i) Then
                        MsgBox "すでに登録されています。", vbExclamation + vbOKOnly, C_TITLE
                        Exit Sub
                    End If
                Next
                 .AddItem Me.txtCat.Text
            Case C_FAVORITE_MOD
                For i = 0 To .ListCount - 1
                    If txtCat.Text = .List(i) And i <> .ListIndex Then
                        MsgBox "すでに登録されています。", vbExclamation + vbOKOnly, C_TITLE
                        Exit Sub
                    End If
                Next
                If frmFavorite.mobjCategory.Exists(.List(.ListIndex)) Then
                    frmFavorite.mobjCategory.key(.List(.ListIndex)) = Me.txtCat.Text
                    Dim cat As Variant
                    Set cat = frmFavorite.mobjCategory.Item(Me.txtCat.Text)
                    
                    i = 0
                    For Each c In cat
                    
                        Set fav = cat.Item(c)
                        fav.Category = Me.txtCat.Text
                    
                    Next
                End If
                .List(.ListIndex) = Me.txtCat.Text
                
        End Select
    End With
    Unload Me
End Sub

Sub Start(ByVal lngMode As Long)

    mlngMode = lngMode
    
    Select Case lngMode
        Case C_FAVORITE_ADD
            lblCat.Caption = "カテゴリの追加"
        Case C_FAVORITE_MOD
            lblCat.Caption = "カテゴリの編集"
            Me.txtCat.Text = frmFavorite.lstCategory.Text
    End Select
    
    Me.Show

End Sub

Private Sub UserForm_Activate()
'    Call AllwaysOnTop
End Sub

