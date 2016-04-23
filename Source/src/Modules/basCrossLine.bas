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
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function SendInput Lib "user32.dll" (ByVal nInputs As LongPtr, pInputs As INPUT_TYPE, ByVal cbsize As LongPtr) As LongPtr
#Else
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbsize As Long) As Long
#End If

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As LongPtr
    time As LongPtr
    dwExtraInfo As LongPtr
    dummy1 As Long
    dummy2 As Long
End Type

Private Type INPUT_TYPE
    dwType As Long
    ki As KEYBDINPUT
End Type

Private Const INPUT_KEYBOARD As Integer = 1
Private Const VK_TAB As Integer = 9
Private Const VK_ENTER As Integer = 13
Private Const KEYEVENTF_KEYDOWN As Integer = 0
Private Const KEYEVENTF_KEYUP As Integer = 2

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
                    .ki.wVk = vbKeyHome
                    .ki.wScan = 0
                    .ki.dwFlags = KEYEVENTF_KEYDOWN
                    .ki.time = 0
                    .ki.dwExtraInfo = 0
                End With

                With inputevents(2)
                    .dwType = INPUT_KEYBOARD
                    .ki.wVk = vbKeyF2
                    .ki.wScan = 0
                    .ki.dwFlags = KEYEVENTF_KEYUP
                    .ki.time = 0
                    .ki.dwExtraInfo = 0
                End With
                
                With inputevents(3)
                    .dwType = INPUT_KEYBOARD
                    .ki.wVk = vbKeyHome
                    .ki.wScan = 0
                    .ki.dwFlags = KEYEVENTF_KEYUP
                    .ki.time = 0
                    .ki.dwExtraInfo = 0
                End With
                
                SendInput 4, inputevents(0), Len(inputevents(0))

            End If
        Else
            r.Select
        End If
    
    End If
    ThisWorkbook.enableCrossLine

End Sub
