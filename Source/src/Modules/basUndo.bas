Attribute VB_Name = "basUndo"
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

'--------------------------------------------------------------
'　Undoバッファエリア
'--------------------------------------------------------------
Public mUndo As New UndoRangeDTO
Public mUndoWS As Worksheet
Public mA1Save As Boolean

'--------------------------------------------------------------
'　選択範囲Undo
'--------------------------------------------------------------
Sub execUndo()

    Dim rr As Range
    
    On Error Resume Next

    If ActiveCell Is Nothing Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For Each rr In mUndo.destRange.Areas
        rr.Copy mUndo.sourceRange.Worksheet.Range(rr.Address)
    Next
    
    Application.ScreenUpdating = True

End Sub

'--------------------------------------------------------------
'　行追加Undo
'--------------------------------------------------------------
Sub execInsUndo()

    Dim rr As Range
    
    On Error Resume Next

    If ActiveCell Is Nothing Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    mUndo.sourceRange.Delete xlUp
    
    Application.ScreenUpdating = True

End Sub
'--------------------------------------------------------------
'　行削除Undo
'--------------------------------------------------------------
Sub execDelUndo()

    Dim rr As Range
    
    On Error Resume Next

    If ActiveCell Is Nothing Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    mUndo.destRange.Copy
    mUndo.sourceRange.Insert Shift:=xlDown
    
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True

End Sub
