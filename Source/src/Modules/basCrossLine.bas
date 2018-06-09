Attribute VB_Name = "basCrossLine"
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
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbsize As Long) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As rect) As Long
    Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

    Private Type KEYBDINPUT
        wVk As Integer
        wScan As Integer
        dwFlags As LongPtr
        time As LongPtr
        dwExtraInfo As LongPtr
        dummy1 As Long
        dummy2 As Long
    End Type

#Else
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbsize As Long) As Long
    Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As rect) As Long
    Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

    Private Type KEYBDINPUT
        wVk As Integer
        wScan As Integer
        dwFlags As Long
        time As Long
        dwExtraInfo As Long
        dummy1 As Long
        dummy2 As Long
    End Type

#End If

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type INPUT_TYPE
    dwType As Long
    ki As KEYBDINPUT
End Type

Public Const INPUT_KEYBOARD As Integer = 1
Public Const VK_TAB As Integer = 9
Public Const VK_ENTER As Integer = 13
Public Const KEYEVENTF_KEYDOWN As Integer = 0
Public Const KEYEVENTF_KEYUP As Integer = 2

Public Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)

Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_LAYERED = &H80000
Public Const WS_CAPTION = &HC00000
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_SYSMENU = &H80000      '最大化／最小化／消去ボタンなど全て
Public Const LWA_ALPHA = 2

'------------------------------------------------------------------------------------------------------------------------
' 十字カーソルクリック時の処理
'------------------------------------------------------------------------------------------------------------------------
Sub rlxCrossClick()

    Dim a As POINTAPI
    Dim r As Range
        
    Dim blnSw As Boolean
    
    On Error Resume Next
    
    ThisWorkbook.disableCrossLine
    
    GetCursorPos a

    Set r = ActiveWindow.RangeFromPoint(a.X, a.Y)
    If r Is Nothing Then
    Else
    
        If ActiveCell.Address = r.Address Then
            If GetSetting(C_TITLE, "CrossLine", "Edit", False) Then

                Dim inputevents(0 To 1) As INPUT_TYPE
                
                With inputevents(0)
                    .dwType = INPUT_KEYBOARD
                    .ki.wVk = vbKeyF2
                    .ki.wScan = 0
                    .ki.dwFlags = KEYEVENTF_KEYDOWN
                    .ki.time = 0
                    .ki.dwExtraInfo = 0
                End With
                
                With inputevents(1)
                    .dwType = INPUT_KEYBOARD
                    .ki.wVk = vbKeyF2
                    .ki.wScan = 0
                    .ki.dwFlags = KEYEVENTF_KEYUP
                    .ki.time = 0
                    .ki.dwExtraInfo = 0
                End With

'                With inputevents(2)
'                    .dwType = INPUT_KEYBOARD
'                    .ki.wVk = vbKeyHome
'                    .ki.wScan = 0
'                    .ki.dwFlags = KEYEVENTF_KEYDOWN
'                    .ki.time = 0
'                    .ki.dwExtraInfo = 0
'                End With
'
'                With inputevents(3)
'                    .dwType = INPUT_KEYBOARD
'                    .ki.wVk = vbKeyHome
'                    .ki.wScan = 0
'                    .ki.dwFlags = KEYEVENTF_KEYUP
'                    .ki.time = 0
'                    .ki.dwExtraInfo = 0
'                End With
                
                SendInput 2, inputevents(0), Len(inputevents(0))

            End If
        Else
            r.Select
        End If
    
    End If
    ThisWorkbook.enableCrossLine

End Sub
