VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionFrameWork"
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
Public Event SelectionInit(ByRef Cancel As Boolean, ByRef Undo As Boolean, ByRef func As Boolean)
'メインイベント
Public Event SelectionMain(ByRef r As Range, ByVal NotHoldFormat As Boolean, ByRef Cancel As Boolean)
'メインイベント(Shape)
Public Event SelectionMainShape(ByRef r As Shape, ByVal NotHoldFormat As Boolean, ByRef Cancel As Boolean)
'終了イベント
Public Event SelectionTerm()

Private mblnNotHoldFormat As Boolean

Public Sub Run()
    Logger.LogTrace TypeName(Me) & ".Run"

    '変数宣言
    Dim r As Range
    Dim Cancel As Boolean
    Dim Undo As Boolean
    Dim func As Boolean
    
    Dim a As Range

    Dim colMerge As Collection
    
    On Error GoTo ErrHandle
    
    'キャンセルの初期化
    Cancel = False
    
    'Undoの初期化
    Undo = False
    
    '式に対しても処理を行うか？
    func = False
    
    '
    mblnNotHoldFormat = CBool(GetSetting(C_TITLE, "Option", "NotHoldFormat", False))
    
    'Selection進捗バー
    Dim objStatus As SelectionStatusBar
    
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
        If Selection.ShapeRange.Count > 0 Then
        
            Application.ScreenUpdating = False
        
            Dim c As Shape
            
            For Each c In Selection.ShapeRange
                
                Select Case c.Type
                    Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                        
                        RaiseEvent SelectionMainShape(c, mblnNotHoldFormat, Cancel)
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
    
    Dim strRange As String
    
    strRange = Selection.Address
    
    '使われているセルと現在選択セルのAND部分のみ処理する。
    Set a = Intersect(ActiveSheet.UsedRange, Selection)
    
    '重複する選択が無い場合終了。
    If a Is Nothing Then
        Exit Sub
    End If
    
    Dim ret As VbMsgBoxResult
    If a.CountLarge > C_MAX_CELLS Then
        ret = MsgBox("大量のセルが選択されています。処理に時間がかかる＆元に戻せませんが続行しますか？", vbInformation + vbOKCancel + vbDefaultButton2, C_TITLE)
        If ret = vbCancel Then
            Exit Sub
        End If
    End If

    '---------------------------
    '初期化イベント
    '---------------------------
    RaiseEvent SelectionInit(Cancel, Undo, func)
    
    'キャンセルの場合
    If Cancel Then
        Exit Sub
    End If
    
    '大量セルの場合UNDO不可
    If ret = vbOK Then
        Undo = False
    End If

    'Undoの場合
    If Undo Then
    
        ThisWorkbook.Worksheets("Undo").Cells.Clear
        
        Set mUndo.sourceRange = a
        Set mUndo.destRange = ThisWorkbook.Worksheets("Undo").Range(a.Address)
        
        Dim rr As Range
        For Each rr In mUndo.sourceRange.Areas
            rr.Copy mUndo.destRange.Worksheet.Range(rr.Address)
        Next
    
    End If

    'Selection進捗バーInstance化
    Set objStatus = New SelectionStatusBar
    
    objStatus.MaxItems = a.Count
    
    Set colMerge = New Collection
    Dim strAddress As String
    
    Application.ScreenUpdating = False
    For Each r In a
    
        ''フィルタおよび非表示対策。
        If r.Rows.Hidden Or r.Columns.Hidden Or r.MergeArea(1).Address <> r.Address Then
            'フィルタまたは非表示の行・列の処理は行わない。
        Else
            
            'マージセル対策
            strAddress = r.MergeArea(1, 1).Address
            Dim e As Range
            
            On Error Resume Next
            Set e = Nothing
            Set e = colMerge(strAddress)
            
            On Error GoTo 0
            
            'アドレスが登録されていない場合イベントを起こす
            If Err.Number = 0 And Not (e Is Nothing) Then
            Else
                '---------------------------
                'メインイベント
                '---------------------------
                If IsEmpty(r.Value) Or IsError(r.Value) Or (Not func And r.HasFormula) Then
                Else
                    RaiseEvent SelectionMain(r, mblnNotHoldFormat, Cancel)
                    If Cancel Then
                        Exit For
                    End If
                End If
                
                colMerge.Add r, strAddress
            End If
            
        End If
        
        'Selection進捗バーの更新
        objStatus.Refresh
    Next
    Application.ScreenUpdating = True
    
    'Selection進捗バーのDispose
    objStatus.Dispose
    
    Set objStatus = Nothing

    Range(strRange).Select
    
    '---------------------------
    '終了イベント
    '---------------------------
    RaiseEvent SelectionTerm
    
    If Undo Then
        'Undo
        Application.OnUndo "Undo", MacroHelper.BuildPath("execUndo")
    End If
    
    
    Exit Sub
ErrHandle:
    Logger.LogFatal TypeName(Me) & ".Run" & "(" & Err.Description & ")"
    
End Sub

'再帰にてグループ以下のシェイプを検索
Private Sub grouprc(ByRef objShape As Shape, ByRef Cancel As Boolean)
    
    Dim c As Shape
    
    For Each c In objShape.GroupItems
        
        Select Case c.Type
            Case msoAutoShape, msoTextBox, msoCallout, msoFreeform
                
                RaiseEvent SelectionMainShape(c, mblnNotHoldFormat, Cancel)
                If Cancel Then
                    Exit For
                End If

            Case msoGroup
                grouprc c, Cancel
            
            Case Else
        End Select
    Next

End Sub
