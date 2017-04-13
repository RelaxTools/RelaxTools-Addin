Attribute VB_Name = "Logger"
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
Option Private Module

#If Win64 And VBA7 Then
    Declare PtrSafe Function GetTickCount Lib "kernel32" Alias "GetTickCount64" () As LongLong
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Private mLog As New LogManager
Private colStack As New Collection

Public Const C_LOG_DEBUGWINDOW As Long = 1
Public Const C_LOG_LOGFILE As Long = 2
Public Const C_LOG_ALL As Long = 3

'インスタンスの取得
Private Function GetInstance() As LogManager

    If mLog Is Nothing Then
        Set mLog = New LogManager
    End If

    Set GetInstance = mLog

End Function

Public Sub LogTrace(Message As String)
    GetInstance.Log LogLevel.Trace, Message
End Sub

Public Sub LogInfo(Message As String)
    GetInstance.Log LogLevel.Info, Message
End Sub

Public Sub LogWarn(Message As String)
    GetInstance.Log LogLevel.Warn, Message
End Sub
Public Sub LogFlash()
    GetInstance.Flash
End Sub
Public Sub LogBegin(Message As String)
    
    If colStack Is Nothing Then
        Set colStack = New Collection
    End If
    colStack.Add GetTickCount
    
    GetInstance.Log LogLevel.Info, Message & ",BEGIN"

End Sub
Public Sub LogFinish(Message As String)
    
#If Win64 And VBA7 Then
    Dim t As LongLong
#Else
    Dim t As Long
#End If

    If colStack Is Nothing Then
        t = 0
    Else
        If colStack.count = 0 Then
            t = 0
        Else
            t = colStack.Item(colStack.count)
            colStack.remove colStack.count
        End If
    End If
    
    If t = 0 Then
        GetInstance.Log LogLevel.Info, Message & ",FINISH,[?]ms"
    Else
        t = GetTickCount - t
        GetInstance.Log LogLevel.Info, Message & ",FINISH,[" & t & "]ms"
    End If
    GetInstance.Flash
End Sub
Public Sub LogFatal(Message As String)
    GetInstance.Log LogLevel.Fatal, "エラーが発生しました。(" & Message & ")"
    GetInstance.Flash
End Sub
Public Property Let Level(ByVal l As LogLevel)
    GetInstance.Level = l
End Property
