VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseWheel"
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
'--------------------------------------------------------------
' 以下URLを参考にしています
'--------------------------------------------------------------
' エクセルの学校
' http://www.excel.studio-kazu.jp/kw/20141022104627.html
'--------------------------------------------------------------
'--------------------------------------------------------------
' コンボやリストボックスでマウスホイールによるスクロール
' basMouseWheel.GetInstance でのみインスタンスを生成する
'--------------------------------------------------------------
Option Explicit

'メインイベント
'Public Event WheelUp(ByVal Name As String)
'Public Event WheelDown(ByVal Name As String)
Public Event WheelUp(ByRef obj As Object)
Public Event WheelDown(ByRef obj As Object)

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
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
Private Const GWL_HINSTANCE As Long = -6
Private Const WH_MOUSE_LL As Long = 14

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

Private mIsRunning As Boolean

#If VBA7 And Win64 Then
    Private mHandle As LongPtr
    Private mHWnd As LongPtr
    Private mActWnd As LongPtr
#Else
    Private mHandle As Long
    Private mHWnd As Long
    Private mActWnd As Long
#End If


'Private mstrName As String
Private mobj As Object
'Public Sub Activate()
'    mActWnd = GetActiveWindow
'End Sub
Public Property Set obj(s As Object)
    Set mobj = s
End Property
'Public Property Let Name(s As String)
'    mstrName = s
'End Property
Public Sub Install(f As UserForm)

    If mIsRunning = False Then
    
        mActWnd = FindWindowA("ThunderDFrame", f.Caption)
        mHWnd = Application.hwnd

        mHandle = SetWindowsHookEx(WH_MOUSE_LL, AddressOf basMouseWheel.MouseEventHookHandler, GetWindowLong(mHWnd, GWL_HINSTANCE), 0)
        
        If mHandle = 0 Then
        Else
            mIsRunning = True
        End If
    End If
End Sub
#If VBA7 And Win64 Then
    Public Function MouseLLHookProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
    
        Dim m_Return As LongPtr
        Dim m_MouseLLHookStruct As MouseLLHookStruct
    
        On Error GoTo ErrorHandler
        If ncode < 0 Or mobj Is Nothing Or GetActiveWindow <> mActWnd Then
            GoTo ErrorHandler
        End If
        
        Select Case wParam
            Case 522
                CopyMemory m_MouseLLHookStruct.POINT.X, ByVal lParam, LenB(m_MouseLLHookStruct)
                With m_MouseLLHookStruct
                    Select Case .mouseData
                        Case Is > 0
'                            RaiseEvent WheelUp(mstrName)
                            If mobj Is Nothing Then
                            Else
                                RaiseEvent WheelUp(mobj)
                            End If
                        Case Is < 0
'                            RaiseEvent WheelDown(mstrName)
                            If mobj Is Nothing Then
                            Else
                                RaiseEvent WheelDown(mobj)
                            End If
                    End Select
                End With
                MouseLLHookProc = 1
                Exit Function
        End Select
    
ErrorHandler:
    
        m_Return = CallNextHookEx(mHandle, ncode, wParam, lParam)
        MouseLLHookProc = m_Return
        
    End Function
#Else
    Public Function MouseLLHookProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
        Dim m_Return As Long
        Dim m_MouseLLHookStruct As MouseLLHookStruct
    
        On Error GoTo ErrorHandler
'            Debug.Print "MouseLLHookProc" & GetActiveWindow

        If ncode < 0 Or mobj Is Nothing Or GetActiveWindow <> mActWnd Then
            GoTo ErrorHandler
        End If
        
        Select Case wParam
            Case 522
                CopyMemory m_MouseLLHookStruct.POINT.X, ByVal lParam, LenB(m_MouseLLHookStruct)
                With m_MouseLLHookStruct
                    Select Case .mouseData
                        Case Is > 0
'                            RaiseEvent WheelUp(mstrName)
                            If mobj Is Nothing Then
                            Else
                                RaiseEvent WheelUp(mobj)
                            End If
                        Case Is < 0
'                            RaiseEvent WheelDown(mstrName)
                            If mobj Is Nothing Then
                            Else
                                RaiseEvent WheelDown(mobj)
                            End If
                    End Select
                End With
                MouseLLHookProc = 1
                Exit Function
        End Select
    
ErrorHandler:
    
        m_Return = CallNextHookEx(mHandle, ncode, wParam, lParam)
        MouseLLHookProc = m_Return
        
    End Function
#End If

Public Sub UnInstall()
    If mIsRunning Then
        Call UnhookWindowsHookEx(mHandle)
        mIsRunning = False
    End If
End Sub



