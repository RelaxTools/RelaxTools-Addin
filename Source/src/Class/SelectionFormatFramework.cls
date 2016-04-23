VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionFormatFramework"
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
' SelectionFrameWork 1.00 2011/07/22 Y.Watanabe
'--------------------------------------------------------------
' Selectionに含まれるCellの数だけ処理を行うフレームワーク
'--------------------------------------------------------------
Option Explicit

'初期化イベント
Public Event SelectionInit(ByRef Cancel As Boolean, ByRef Undo As Boolean)
'メインイベント
Public Event SelectionMain(ByRef r As Range)
'メインイベント(Shape)
Public Event SelectionMainShape(ByRef r As Shape, ByRef Cancel As Boolean)
'終了イベント
Public Event SelectionTerm()

Public Sub Run()

    '変数宣言
    Dim r As Range
    Dim Cancel As Boolean
    Dim Undo As Boolean

    'キャンセルの初期化
    Cancel = False
    Undo = False
    
'    'Selection進捗バー
'    Dim objStatus As SelectionStatusBar
    
    On Error GoTo ErrHandle
    
    
    
    Dim blnRange As Boolean
    Dim blnObject As Boolean
    blnRange = False
    blnObject = False
    Select Case True
        Case ActiveWorkbook Is Nothing
        Case ActiveCell Is Nothing
        Case Selection Is Nothing
        Case TypeOf Selection Is Shape
        Case TypeOf Selection Is Range
            blnRange = True
        Case TypeOf Selection Is Object
            blnObject = True
        Case Else
    End Select
    If blnRange Or blnObject Then
    Else
        MsgBox "選択範囲が見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    'Shepeの場合の実行処理
    If blnObject Then
        If Selection.ShapeRange.count > 0 Then
        
            Application.ScreenUpdating = True
        
            Dim c As Shape
            
            For Each c In Selection.ShapeRange
                
                Select Case c.Type
                    Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                        
                        RaiseEvent SelectionMainShape(c, Cancel)
                        If Cancel Then
                            Exit For
                        End If
                    
                    Case msoGroup
                        grouprc c, Cancel
        
                    Case Else
        
                End Select
            Next
            
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If
    
'
'
'
'    If Selection Is Nothing Then
'        MsgBox "選択範囲が見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    If TypeOf Selection Is Range Then
'    Else
'        MsgBox "選択範囲が見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
    
    '---------------------------
    '初期化イベント
    '---------------------------
    RaiseEvent SelectionInit(Cancel, Undo)
    
    'キャンセルの場合
    If Cancel Then
        Exit Sub
    End If
    
    'Undoの場合
    If Undo Then
        
        ThisWorkbook.Worksheets("Undo").Cells.Clear
        
        Set mUndo.sourceRange = Selection
        Set mUndo.destRange = ThisWorkbook.Worksheets("Undo").Range(Selection.Address)
        
        Dim rr As Range
        For Each rr In mUndo.sourceRange.Areas
            rr.Copy mUndo.destRange.Worksheet.Range(rr.Address)
        Next
        
    End If
                
    '---------------------------
    'メインイベント
    '---------------------------
    Application.ScreenUpdating = False
    
    RaiseEvent SelectionMain(Selection)

    '---------------------------
    '終了イベント
    '---------------------------
    
    RaiseEvent SelectionTerm
    
    Application.ScreenUpdating = True
    
    If Undo Then
        'Undo
        Application.OnUndo "Undo", "execUndo"
    End If
    
    Exit Sub
ErrHandle:
    MsgBox "エラーが発生しました。", vbOKOnly, C_TITLE

End Sub


'再帰にてグループ以下のシェイプを検索
Private Sub grouprc(ByRef objShape As Shape, ByRef Cancel As Boolean)
    
    Dim c As Shape
    
    For Each c In objShape.GroupItems
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                
                RaiseEvent SelectionMainShape(c, Cancel)
                If Cancel Then
                    Exit For
                End If

            Case msoGroup
                grouprc c, Cancel
            
            Case Else
        End Select
    Next

End Sub

