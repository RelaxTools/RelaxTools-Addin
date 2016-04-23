VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CryptoServiceProvider"
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
Option Explicit
Public Enum CryptoServiceProviderConstants
    HashTypeMD5
    HashTypeSHA1
    HashTypeSHA256
    HashTypeSHA384
    HashTypeSHA512
End Enum
Private mHashType As CryptoServiceProviderConstants
Public Property Let HashType(h As CryptoServiceProviderConstants)
    mHashType = h
End Property
Public Function ComputeHash(ByRef bytBuf() As Byte) As String

    Dim md As Object
    Dim bytRet() As Byte
    Dim strBuf As String
    Dim i As Long
    Dim lngFrom As Long
    Dim lngSize As Long
    
    On Error GoTo e
    Select Case mHashType
        Case HashTypeMD5
            Set md = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
        Case HashTypeSHA1
            Set md = CreateObject("System.Security.Cryptography.SHA1Managed")
        Case HashTypeSHA256
            Set md = CreateObject("System.Security.Cryptography.SHA256Managed")
        Case HashTypeSHA384
            Set md = CreateObject("System.Security.Cryptography.SHA384Managed")
        Case HashTypeSHA512
            Set md = CreateObject("System.Security.Cryptography.SHA512Managed")
    End Select

    lngFrom = LBound(bytBuf())
    lngSize = UBound(bytBuf()) - lngFrom + 1

    bytRet = md.ComputeHash_3(bytBuf, lngFrom, lngSize)
    
    md.Clear
    
    Set md = Nothing
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    
    For i = LBound(bytRet) To UBound(bytRet)
        sb.Append Right$("00" & Hex(bytRet(i)), 2)
    Next
    
    ComputeHash = LCase(sb.ToString)
    
    Set sb = Nothing
    
    Exit Function
e:
    ComputeHash = "*ê∂ê¨ïsâ¬*"
    
End Function

