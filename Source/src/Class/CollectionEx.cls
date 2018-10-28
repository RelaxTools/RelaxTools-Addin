VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'  Collection ラッパークラス
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private col As Collection

Private Sub Class_Initialize()
    Set col = New Collection
End Sub
Private Sub Class_Terminate()
    Set col = Nothing
End Sub
Public Function Add(obj As Variant, Optional Key As Variant)
    If IsMissing(Key) Then
        col.Add obj
    Else
        col.Add obj, Key
    End If
End Function
Public Function Exists(ByVal Key As String) As Boolean

    On Error Resume Next

    Err.Clear
    col.Item Key
    
    Select Case Err.Number
        Case 0
            Exists = True
        Case 5
            Exists = False
        Case Else
            Exists = False
    End Select

    Err.Clear

End Function

Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = col.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = col.Count
End Property
Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    
    If Me.Exists(Key) Then
        If VBA.IsObject(col.Item(Key)) Then
            Set Item = col.Item(Key)
        Else
            Item = col.Item(Key)
        End If
    Else
        Err.Raise vbObjectError + 512 + 1, , "キーが存在しません。"
    End If
    
End Property
Public Sub Remove(Key As Variant)
    col.Remove Key
End Sub
'--------------------------------------------------------------
'  コレクションのソート
'--------------------------------------------------------------
Public Sub Sort()

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim col2 As Collection
    Dim var() As Variant
    Dim wk As String

    'Collectionが空ならなにもしない
    If col Is Nothing Then
        Exit Sub
    End If

    'Collectionの要素数が０または１の場合ソート不要
    If col.Count <= 1 Then
        Exit Sub
    End If

    n = col.Count
    ReDim var(1 To n)

    For i = 1 To n
        var(i) = col.Item(i)
    Next

    '挿入ソート
    For i = 2 To n

        wk = var(i)

        If var(i - 1) > wk Then

            j = i

            Do

                var(j) = var(j - 1)

                j = j - 1

                If j = 1 Then
                    Exit Do
                End If

            Loop While var(j - 1) > wk
            var(j) = wk

        End If
    Next

    Set col2 = New Collection

    For i = 1 To n
        col2.Add var(i)
    Next

    Set col = col2
    Set col2 = Nothing

End Sub
