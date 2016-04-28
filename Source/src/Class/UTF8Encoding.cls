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

'Private utf8 As Object

Private Sub Class_Initialize()

'    Set utf8 = CreateObject("System.Text.UTF8Encoding")

End Sub

Private Sub Class_Terminate()

'    Set utf8 = Nothing

End Sub
Public Function GetString(ByRef bytBuf() As Byte) As String

    Dim bytRet() As Byte
    Dim i As Long
    Dim lngPos As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    
    i = LBound(bytBuf)
    ReDim bytRet(0 To UBound(bytBuf) * 2 - 1)
    
    
    Do Until i >= UBound(bytBuf)
    
        b0 = bytBuf(i): i = i + 1
    
        Select Case True

'            // UTF-8:   [&hxx xxxx]
'            // Unicode: [0000 0000] [&hxx xxxx]
            Case (b0 < &H80&)
                bytRet(lngPos) = b0: lngPos = lngPos + 1
                bytRet(lngPos) = 0:  lngPos = lngPos + 1

'            // UTF-8:   [110y yyyy] [1&hx xxxx]
'            // Unicode: [0000 0yyy] [yyxx xxxx]
'            Case ((b0 And &HE0&) = &HC0 And (b0 And &H1E&) <> 0)
            Case ((b0 And &HF0&) = &HC0)
            
                b1 = bytBuf(i): i = i + 1
                Dim c As Long
                c = ((LShift(b0, 6)) And &H7C0&) Or (b1 And &H3F&)
                
                bytRet(lngPos) = LByte(c): lngPos = lngPos + 1
                bytRet(lngPos) = UByte(c): lngPos = lngPos + 1
                
'                count -= 1;
'                continue;

'            // UTF-8:   [1110 zzzz] [10yy yyyy] [1&hx xxxx]
'            // Unicode: [zzzz yyyy] [yyxx xxxx]
            Case ((b0 And &HF0&) = &HE0&)
            
                b1 = bytBuf(i): i = i + 1
                b2 = bytBuf(i): i = i + 1
            
                c = ((LShift(b0, 12)) And &HF000&) Or ((LShift(b1, 6)) And &HFC0&) Or (b2 And &H3F&)
                
                bytRet(lngPos) = LByte(c)
                lngPos = lngPos + 1
                bytRet(lngPos) = UByte(c)
                lngPos = lngPos + 1

'            // UTF-8:   [1111 0uuu] [10uu zzzz] [10yy yyyy] [1&hx xxxx]*
'            // Unicode: [1101 10ww] [wwzz zzyy] (high surrogate)
'            //          [1101 11yy] [yyxx xxxx] (low surrogate)
'            //          * uuuuu = wwww + 1
            Case ((b0 And &HF0) = &HF0&)
                
                b1 = bytBuf(i): i = i + 1
                b2 = bytBuf(i): i = i + 1
                b3 = bytBuf(i): i = i + 1
                
                Dim uuuuu As Long
                Dim wwww As Long
                Dim zzzz As Long
                Dim yyyyyy As Long
                Dim xxxxxx As Long
                Dim hs As Long
                Dim ls As Long

'                // decode bytes into surrogate characters
                uuuuu = ((LShift(b0, 2)) And &H1C&) Or ((RShift(b1, 4)) And &H3&)
'                If (uuuuu > &H10) Then
'                    invalidSurrogate(uuuuu);
'                End If
                wwww = uuuuu - 1
                zzzz = b1 And &HF&
                yyyyyy = b2 And &H3F&
                xxxxxx = b3 And &H3F&
                
                hs = &HD800& Or ((LShift(wwww, 6)) And &H3C0&) Or (LShift(zzzz, 2)) Or (RShift(yyyyyy, 4))
                ls = &HDC00& Or ((LShift(yyyyyy, 6)) And &H3C0&) Or xxxxxx

                bytRet(lngPos) = LByte(hs)
                lngPos = lngPos + 1
                bytRet(lngPos) = UByte(hs)
                lngPos = lngPos + 1

                bytRet(lngPos) = LByte(ls)
                lngPos = lngPos + 1
                bytRet(lngPos) = UByte(ls)
                lngPos = lngPos + 1

        
        End Select
            
    Loop
    
    GetString = LeftB(bytRet, lngPos)
        

'    GetString = utf8.GetString(bytBuf)

End Function



Public Function getBytes(ByVal strBuf As String) As Byte()

    Dim bytBuf() As Byte
    Dim lngBuf As Long
    Dim bytRet() As Byte
    
    Dim i As Long
    Dim lngPos As Long
    
    If strBuf = "" Then
        Exit Function
    End If

    bytBuf = strBuf
    
    'バッファを最大 1文字×4バイト分確保
    ReDim bytRet(0 To (Len(strBuf) * 4) - 1)
    
    lngPos = 0

    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf

            Case Is < &H80&
            
                'UTF-8(ASCII)
                bytRet(lngPos) = lngBuf
                lngPos = lngPos + 1
            
            Case Is < &H800&
            
                'UTF-8(2バイト)
                bytRet(lngPos) = &HC0& Or RShift(lngBuf, 6)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (lngBuf And &H3F&)
                lngPos = lngPos + 1
      
            Case &HD800& To &HDBFF&

                Dim lngHigh As Long
                Dim lngLow As Long

                lngHigh = lngBuf
                
                i = i + 2
                lngLow = LShift(bytBuf(i + 1), 8) + bytBuf(i)
                
                'サロゲート(UTF-16→Unicode)
                lngBuf = &H10000 + (lngHigh - &HD800&) * &H400& + (lngLow - &HDC00&)
                
                'UTF-8(4バイト)
                bytRet(lngPos) = &HF0& Or RShift(lngBuf, 18)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (RShift(lngBuf, 12) And &H3F&)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (RShift(lngBuf, 6) And &H3F&)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (lngBuf And &H3F&)
                lngPos = lngPos + 1
          
            Case Else
                
                'UTF-8(3バイト)
                bytRet(lngPos) = &HE0& Or RShift(lngBuf, 12)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (RShift(lngBuf, 6) And &H3F&)
                lngPos = lngPos + 1
                
                bytRet(lngPos) = &H80& Or (lngBuf And &H3F&)
                lngPos = lngPos + 1
      
        End Select
    
    Next
    
    getBytes = LeftB(bytRet, lngPos)

End Function
Private Function LByte(ByVal lngValue As Long) As Long
    LByte = lngValue And &HFF&
End Function
Private Function UByte(ByVal lngValue As Long) As Long
    UByte = RShift((lngValue And &HFF00&), 8)
End Function
Private Function RShift(ByVal lngValue As Long, ByVal lngKeta As Long) As Long
    RShift = lngValue \ (2 ^ lngKeta)
End Function
Private Function LShift(ByVal lngValue As Long, ByVal lngKeta As Long) As Long
    LShift = lngValue * (2 ^ lngKeta)
End Function
