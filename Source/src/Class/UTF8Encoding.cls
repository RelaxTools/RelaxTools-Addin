VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTF8Encoding"
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

Private utf8 As Object

Public Function GetString(ByRef bytBuf() As Byte) As String

    GetString = utf8.GetString(bytBuf)

End Function
Public Function getBytes(ByRef strBuf As String) As Byte()

    Dim bytBuf() As Byte
    Dim lngBuf As Long
    Dim bytRet() As Byte
    
    Dim i As Long
    Dim lngPos As Long

    bytBuf = strBuf
    bytRet = String(Len(strBuf) * 2, vbNullChar)
    lngPos = 0

    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = bytBuf(i + 1) * &H100& + bytBuf(i)
    
        Select Case lngBuf

            Case Is < &H80&
            
                bytRet(lngPos) = lngBuf
                lngPos = lngPos + 1
            
            Case Is < &H800&
            
                bytRet(lngPos) = &HC0& Or Fix(lngBuf / (2 ^ 6))
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (lngBuf And &H3F&)
                lngPos = lngPos + 1
      
            Case &HD800& To &HDBFF&

                Dim lngHigh As Long
                Dim lngLow As Long

                lngHigh = lngBuf
                lngLow = bytBuf(i + 3) * &H100& + bytBuf(i + 2)

                bytRet(lngPos) = &HF0& Or (Fix((lngHigh + &H40&) / (2 ^ 8)) And &HF0&)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (Fix((lngHigh + &H40&) / (2 ^ 2)) And &H3F&)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or ((Fix(lngLow / (2 ^ 6)) And &HF&) + ((lngHigh * (2 ^ 4)) And &H30&))
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (lngLow And &H3F&)
                lngPos = lngPos + 1
          
            Case Else
                
                bytRet(lngPos) = &HE0& Or Fix(lngBuf / (2 ^ 12))
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (Fix(lngBuf / (2 ^ 6)) And &H3F&)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (lngBuf And &H3F&)
                lngPos = lngPos + 1
      
        End Select
    
    Next
    
    getBytes = LeftB(bytRet, lngPos)

End Function

Private Sub Class_Initialize()

    Set utf8 = CreateObject("System.Text.UTF8Encoding")

End Sub

Private Sub Class_Terminate()

    Set utf8 = Nothing

End Sub
