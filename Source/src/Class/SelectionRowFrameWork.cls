VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionRowFrameWork"
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
' SelectionRowFrameWork 1.00 2011/07/22 Y.Watanabe
'--------------------------------------------------------------
' Selection.Areasを行単位に処理するフレームワーク
'--------------------------------------------------------------
Option Explicit

'初期化イベント
Public Event SelectionInit(ByRef rArea As Excel.Areas, ByRef Cancel As Boolean, ByRef Undo As Boolean)
'前処理イベント
Public Event SelectionBegin(ByRef rArea As Range, ByRef Cancel As Boolean)
'メイン処理イベント
Public Event SelectionMain(ByRef r As Range, ByVal row As Long, ByRef Cancel As Boolean)
'後処理イベント
Public Event SelectionFinal(ByRef rAarea As Range)
'終了イベント
Public Event SelectionTerm()

Public Sub Run()

    '変数宣言
    Dim r As Range
    Dim Cancel As Boolean
    Dim Undo As Boolean
    Dim lngRows As Long
    Dim lngCols As Long
    Dim lngCur As Long
    Dim i As Long
    
    'Selection進捗バー
    Dim objStatus As SelectionStatusBar
    
    On Error GoTo ErrHandle
    
    If Selection Is Nothing Then
        MsgBox "選択範囲が見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If

    If TypeOf Selection Is Range Then
    Else
        MsgBox "選択範囲が見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
'    If selection.count > C_MAX_CELLS Then
    If Selection.CountLarge > C_MAX_CELLS Then
        Dim ret As VbMsgBoxResult
        ret = MsgBox("大量のセルが選択されています。処理に時間がかかる＆元に戻せませんが続行しますか？", vbInformation + vbOKCancel + vbDefaultButton2, C_TITLE)
        If ret = vbCancel Then
            Exit Sub
        End If
    End If
        
    'キャンセルの初期化
    Cancel = False
    Undo = False
    
    '---------------------------
    '初期化イベント
    '---------------------------
    RaiseEvent SelectionInit(Selection.Areas, Cancel, Undo)

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
        
        Set mUndo.sourceRange = Selection
        Set mUndo.destRange = ThisWorkbook.Worksheets("Undo").Range(Selection.Address)
        
        Dim rr As Range
        For Each rr In mUndo.sourceRange.Areas
            rr.Copy mUndo.destRange.Worksheet.Range(rr.Address)
        Next
    
    End If

    'Selection進捗バーInstance化
    Set objStatus = New SelectionStatusBar
    
    '処理最大数をセット
    Dim lngMax As Long
    Dim nArea As Range
    lngMax = 0
    For Each nArea In Selection.Areas
        lngMax = lngMax + nArea.Rows.count
    Next
    objStatus.MaxItems = lngMax

    Dim rArea As Range
    
    Application.ScreenUpdating = False
    For Each rArea In Selection.Areas

        '---------------------------
        '前処理イベント
        '---------------------------
        RaiseEvent SelectionBegin(rArea, Cancel)
        If Cancel Then
            Exit For
        End If
        
        lngRows = rArea.Rows.count
        lngCols = rArea.Columns.count
    
        For lngCur = 1 To lngRows
        
            '１行分のRangeを作成
            
            Set r = Nothing
            
            For i = 1 To lngCols
                ''フィルタおよび非表示対策。
                If rArea(lngCur, i).Rows.Hidden Or rArea(lngCur, i).Columns.Hidden Then
                Else
                    If r Is Nothing Then
                        Set r = rArea(lngCur, i)
                    Else
                        Set r = Union(r, rArea(lngCur, i))
                    End If
                End If
            Next
        
            '---------------------------
            'メイン処理イベント
            '---------------------------
            If r Is Nothing Then
            Else
                RaiseEvent SelectionMain(r, lngCur, Cancel)
                If Cancel Then
                    Exit For
                End If
            End If
            
            'Selection進捗バーの更新
            objStatus.Refresh
        Next
        '---------------------------
        '後処理イベント
        '---------------------------
        RaiseEvent SelectionFinal(rArea)
    
    Next
    Application.ScreenUpdating = True
    
    'Selection進捗バーのDispose
    objStatus.Dispose
    
    Set objStatus = Nothing

    '---------------------------
    '終了イベント
    '---------------------------
    Application.ScreenUpdating = False
    
    RaiseEvent SelectionTerm
    
    Application.ScreenUpdating = True
    
    If Undo Then
        'Undo
        Application.OnUndo "Undo", "execUndo"
    End If
    
    Exit Sub
ErrHandle:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。", vbOKOnly, C_TITLE
End Sub


