VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHoldBook 
   Caption         =   "ピン留め中のブック"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11640
   OleObjectBlob   =   "frmHoldBook.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmHoldBook"
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
'Private Const C_NO As Long = 0
'Private Const C_BOOK As Long = 1
'
'
'Private mDic As Object
'
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdDel_Click()
'
'    Dim KEY As String
'    Dim i As Long
'
'    If lstBook.ListCount = 0 Or lstBook.ListIndex < 0 Then
'        Exit Sub
'    End If
'
'    i = lstBook.ListIndex
'
'    KEY = lstBook.List(i, C_BOOK)
'
'    If mDic.Exists(KEY) Then
'        mDic.Remove KEY
'        SaveHoldList mDic
'        lstBook.RemoveItem i
'        If i < lstBook.ListCount - 1 Then
'            lstBook.ListIndex = i
'        Else
'            lstBook.ListIndex = lstBook.ListCount - 1
'        End If
'    End If
'
'End Sub
'
'Private Sub UserForm_Initialize()
'
'    disp
'
'End Sub
'Sub disp()
'
'    Set mDic = GetHoldList()
'    Dim h As HoldDto
'
'    Dim c As Variant
'    Dim i As Long
'
'    i = 0
'    For Each c In mDic.keys
'
'        Set h = mDic.Item(c)
'        lstBook.AddItem ""
'        lstBook.List(i, C_NO) = i + 1
'        lstBook.List(i, C_BOOK) = h.FullName
'        i = i + 1
'    Next
'
'    If i > 0 Then
'        lstBook.ListIndex = 0
'    End If
'
'End Sub
'Private Sub UserForm_Terminate()
'
'
'    Set mDic = Nothing
'
'
'End Sub
Private Sub UserForm_Click()

End Sub
