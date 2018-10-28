VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryEx"
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
'  Dictionary ラッパークラス (参照設定はしたくないけどインテリセンスしたいわがままクラス）
'-----------------------------------------------------------------------------------------------------
Option Explicit
Public Enum CompareMethod
    BinaryCompare = 0
    TextCompare = 1
    DatabaseCompare = 2
End Enum

Private dic As Object

Private Sub Class_Initialize()
    Set dic = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set dic = Nothing
End Sub
Public Function Add(ByVal KeyValue As String, obj As Variant)
    dic.Add KeyValue, obj
End Function
Public Function Exists(ByVal KeyValue As String) As Boolean
    Exists = dic.Exists(KeyValue)
End Function
Public Function keys() As Variant
    keys = dic.keys
End Function
Public Function SortedKeys() As Variant
    SortedKeys = Sort(dic.keys)
End Function
Public Function Items() As Variant
    Items = dic.Items
End Function
Public Property Get CompareMode() As CompareMethod
    CompareMode = dic.CompareMode
End Property
Public Property Let CompareMode(Mode As CompareMethod)
    dic.CompareMode = Mode
End Property
Public Property Get Count() As Long
    Count = dic.Count
End Property
Public Property Get Item(ByVal KeyValue As String) As Variant
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(dic.Item(KeyValue)) Then
        Set Item = dic.Item(KeyValue)
    Else
        Item = dic.Item(KeyValue)
    End If
End Property
Public Property Let Item(ByVal KeyValue As String, NewItem As Variant)
    dic.Item(KeyValue) = NewItem
End Property
Public Property Let Key(ByVal OldKey As String, NewKey As Variant)
    dic.Key(OldKey) = NewKey
End Property
Public Sub Remove(ByVal KeyValue As String)
    dic.Remove KeyValue
End Sub
Public Sub RemoveAll()
    dic.RemoveAll
End Sub
Private Function Sort(strKey As Variant) As Variant

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim wk As String

    '要素数
    n = UBound(strKey) - LBound(strKey) + 1
    
    'ソート不要
    If n <= 1 Then
        GoTo e
    End If

    '挿入ソート
    For i = 1 To n - 1

        wk = strKey(i)

        If strKey(i - 1) > wk Then

            j = i

            Do

                strKey(j) = strKey(j - 1)

                j = j - 1

                If j = 0 Then
                    Exit Do
                End If

            Loop While strKey(j - 1) > wk
            strKey(j) = wk

        End If
    Next

e:
    Sort = strKey

End Function
