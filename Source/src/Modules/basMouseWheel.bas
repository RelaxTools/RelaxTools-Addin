Attribute VB_Name = "basMouseWheel"
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
' 以下URLを参考にしています
'--------------------------------------------------------------
' エクセルの学校
' http://www.excel.studio-kazu.jp/kw/20141022104627.html
'--------------------------------------------------------------
Option Explicit
Option Private Module
Private MH As New MouseWheel
'--------------------------------------------------------------
' 必ずこのメソッドにてインスタンスを取得すること。
'--------------------------------------------------------------
Public Property Get GetInstance() As MouseWheel

    If MH Is Nothing Then
        Set MH = New MouseWheel
    End If

    Set GetInstance = MH

End Property
    
'--------------------------------------------------------------
' SetWindowsHookExのコールバックメソッド(直接呼ばないこと)
'--------------------------------------------------------------
#If VBA7 And Win64 Then
    Public Function MouseEventHookHandler(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
    
        On Error GoTo ErrorHandler
        MouseEventHookHandler = GetInstance.MouseLLHookProc(uMsg, wParam, lParam)
        
ErrorHandler:
    End Function
#Else
    Public Function MouseEventHookHandler(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
        On Error GoTo ErrorHandler
        MouseEventHookHandler = GetInstance.MouseLLHookProc(uMsg, wParam, lParam)
        
ErrorHandler:
    End Function
#End If
