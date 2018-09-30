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

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hhk As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hhk As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
#End If

Private Type POINT
    X As Long
    Y As Long
End Type
    
Private Type MouseLLHookStruct
    POINT As POINT
    mouseData As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Const GWL_HINSTANCE As Long = -6
Private Const WH_MOUSE_LL As Long = 14

Private mlngRefCount As Long

#If VBA7 And Win64 Then
    Private mHandle As LongPtr
#Else
    Private mHandle As Long
#End If

Private mColhWnd As Object

Public Function Install(ByVal strhWnd As String) As MouseWheel

    Dim MH As MouseWheel

    '参照カウンタ
    mlngRefCount = mlngRefCount + 1
    
    If mlngRefCount = 1 Then
        mHandle = SetWindowsHookEx(WH_MOUSE_LL, AddressOf basMouseWheel.MouseEventHookHandler, GetWindowLong(Application.hWnd, GWL_HINSTANCE), 0)
    End If
    
    If mHandle <> 0 Then
        Set MH = New MouseWheel
        
        If mColhWnd Is Nothing Then
            Set mColhWnd = CreateObject("Scripting.Dictionary")
        End If
        
        mColhWnd.Add strhWnd, MH
        Set Install = MH
    End If
    
End Function
Public Function UnInstall(ByVal strhWnd As String) As MouseWheel

    If mHandle <> 0 Then
        If mColhWnd.Exists(strhWnd) Then
            mColhWnd.Remove strhWnd
        End If
    End If
    
    If mlngRefCount = 1 Then
        If mHandle <> 0 Then
            Call UnhookWindowsHookEx(mHandle)
            mHandle = 0
        End If
    End If
    
    mlngRefCount = mlngRefCount - 1
    
    Set UnInstall = Nothing

End Function
'--------------------------------------------------------------
' SetWindowsHookExのコールバックメソッド(直接呼ばないこと)
'--------------------------------------------------------------
#If VBA7 And Win64 Then
    Public Function MouseEventHookHandler(ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    
        Dim m_MouseLLHookStruct As MouseLLHookStruct
        Dim MH As MouseWheel
        Dim strhWnd As String
        
        On Error GoTo ErrorHandler
        
        If uMsg < 0 Or wParam <> 522 Then
            GoTo ErrorHandler
        End If
    
        strhWnd = CStr(GetActiveWindow)
        
        If mColhWnd.Exists(strhWnd) Then
            
            CopyMemory m_MouseLLHookStruct.POINT.X, ByVal lParam, LenB(m_MouseLLHookStruct)
            
            Set MH = mColhWnd(strhWnd)
            Call MH.MouseLLHookProc(m_MouseLLHookStruct.mouseData)
            
            MouseEventHookHandler = 1
            Exit Function
        End If

ErrorHandler:
        MouseEventHookHandler = CallNextHookEx(mHandle, uMsg, wParam, lParam)
    End Function
#Else
    Public Function MouseEventHookHandler(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
        Dim m_MouseLLHookStruct As MouseLLHookStruct
        Dim MH As MouseWheel
        Dim strhWnd As String
        
        On Error GoTo ErrorHandler
        
        If uMsg < 0 Or wParam <> 522 Then
            GoTo ErrorHandler
        End If
    
        strhWnd = CStr(GetActiveWindow)
        
        If mColhWnd.Exists(strhWnd) Then
            
            CopyMemory m_MouseLLHookStruct.POINT.X, ByVal lParam, LenB(m_MouseLLHookStruct)
            
            Set MH = mColhWnd(strhWnd)
            Call MH.MouseLLHookProc(m_MouseLLHookStruct.mouseData)
            
            MouseEventHookHandler = 1
            Exit Function
        End If

ErrorHandler:
        MouseEventHookHandler = CallNextHookEx(mHandle, uMsg, wParam, lParam)
    End Function
#End If
