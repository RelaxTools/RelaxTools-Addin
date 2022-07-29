Attribute VB_Name = "basFunction"
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
Private Const C_SHA256 As Long = 1
Private Const C_SHA384 As Long = 2
Private Const C_SHA512 As Long = 3
'--------------------------------------------------------------
'　SHA256算出関数
'--------------------------------------------------------------
Public Function GetSHA256(ハッシュ値算出範囲 As Range) As Variant
Attribute GetSHA256.VB_Description = "ハッシュ値(SHA256)を取得する関数。.Net3.5のインストールが別途必要"
Attribute GetSHA256.VB_ProcData.VB_Invoke_Func = " \n20"
    Application.Volatile
    GetSHA256 = ComputeSHA(C_SHA256, ハッシュ値算出範囲)

End Function
'--------------------------------------------------------------
'　SHA384算出関数
'--------------------------------------------------------------
Public Function GetSHA384(ハッシュ値算出範囲 As Range) As Variant
Attribute GetSHA384.VB_Description = "ハッシュ値(SHA384)を取得する関数。.Net3.5のインストールが別途必要"
Attribute GetSHA384.VB_ProcData.VB_Invoke_Func = " \n20"
    Application.Volatile
    GetSHA384 = ComputeSHA(C_SHA384, ハッシュ値算出範囲)

End Function
'--------------------------------------------------------------
'　SHA2512算出関数
'--------------------------------------------------------------
Public Function GetSHA512(ハッシュ値算出範囲 As Range) As Variant
Attribute GetSHA512.VB_Description = "ハッシュ値(SHA512)を取得する関数。.Net3.5のインストールが別途必要"
Attribute GetSHA512.VB_ProcData.VB_Invoke_Func = " \n20"
    Application.Volatile
    GetSHA512 = ComputeSHA(C_SHA512, ハッシュ値算出範囲)

End Function
Private Function ComputeSHA(lngType As Long, r As Range) As Variant

    Dim objUTF8 As Object
    Dim objSHA As Object
    
    On Error GoTo e

10: Set objUTF8 = CreateObject("System.Text.UTF8Encoding")

    Select Case lngType
        Case C_SHA256
            Set objSHA = CreateObject("System.Security.Cryptography.SHA256Managed")
        Case C_SHA384
            Set objSHA = CreateObject("System.Security.Cryptography.SHA384Managed")
        Case C_SHA512
            Set objSHA = CreateObject("System.Security.Cryptography.SHA512Managed")
    End Select
    
    Dim str As String

    'Rangeの文字列結合（ワークシート関数のConcatを流用）
    str = Application.WorksheetFunction.Concat(r)

    'バイト読み込み
    Dim code() As Byte
    code = objUTF8.GetBytes_4(str)

    'ハッシュ値計算
    Dim hashValue() As Byte
    hashValue = objSHA.ComputeHash_2(code)

    '16進数へ変換
    Dim description As String
    
    description = ""
    
    Dim i As Long
    For i = LBound(hashValue()) To UBound(hashValue())
        description = description & Right("0" & Hex(hashValue(i)), 2)
    Next i

    'return
    ComputeSHA = LCase(description)

    Exit Function

e:
    Select Case Erl
        Case 10
            ComputeSHA = ".Net3.5がインストールされてない可能性があります。Windows機能の有効化で.Net3.5をインストールしてください。"
        Case Else
            ComputeSHA = "原因不明のエラー"
    End Select

End Function
