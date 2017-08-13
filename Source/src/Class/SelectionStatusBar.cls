VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionStatusBar"
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
' Selection の処理を行う際に表示するステータスバーの表示クラス
'--------------------------------------------------------------
Option Explicit
Private mblnStatusBar As Boolean
'Private mlngMaxItem As Long
'Private mlngIndexStatusBar As Long
Private mlngMaxItem As Variant
Private mlngIndexStatusBar As Variant
Private mblnNotCallDispose As Boolean
Private mobjPointer As Excel.XlMousePointer
Private mobjCalculation As Excel.XlCalculation
Private mblnEnableEventsas As Boolean

'-------------------------------------
' Initialize
'-------------------------------------
Private Sub Class_Initialize()

    On Error Resume Next
    
    'ステータスバーの状態を記憶
    mblnStatusBar = Application.DisplayStatusBar
    
    'ステータスバーを表示
    Application.DisplayStatusBar = True

    'メンバ変数の初期化
    mlngIndexStatusBar = 0
    mblnNotCallDispose = True

    'Selectionより選択セル数を取得する。
    If Selection Is Nothing Then
        mlngMaxItem = 0
    Else
        mlngMaxItem = Selection.CountLarge
    End If

    'マウスポインタを砂時計にする。
    mobjPointer = Application.Cursor
    Application.Cursor = xlWait
    mobjCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    mblnEnableEventsas = Application.EnableEvents
    Application.EnableEvents = False
    
End Sub
'-------------------------------------
' Refresh
'    SelectionのForEachループ中に１度呼び出すこと。
'-------------------------------------
Public Sub Refresh()
    
    On Error Resume Next
    
    'ステータスバーにメッセージを表示
    mlngIndexStatusBar = mlngIndexStatusBar + 1   ''カウントアップ
    
    '５％ずつ表示。
    If Int(mlngMaxItem * 0.05) <> 0 Then
        If mlngIndexStatusBar Mod Int(mlngMaxItem * 0.05) = 0 Then
            Application.StatusBar = mlngIndexStatusBar & " / " & mlngMaxItem & " 件 処理中です..."
        End If
    Else
        Application.StatusBar = mlngIndexStatusBar & " / " & mlngMaxItem & " 件 処理中です..."
    End If
    DoEvents
    
End Sub
'-------------------------------------
' Dispose 必ず最後に呼び出すこと。
'-------------------------------------
Public Sub Dispose()

    On Error Resume Next
    
    mblnNotCallDispose = False
    
    Application.EnableEvents = mblnEnableEventsas
    Application.Calculation = mobjCalculation
    Application.Cursor = mobjPointer
    
    'ステータスバーの表示内容をExcelの既定値に戻す
    Application.StatusBar = False
    
    'ステータスバーをマクロの実行前の状態に戻す
    Application.DisplayStatusBar = mblnStatusBar

End Sub
'-------------------------------------
' Terminate
'-------------------------------------
Private Sub Class_Terminate()

    On Error Resume Next
    
    ''Dispose が呼ばれなかった場合の対策
    If mblnNotCallDispose Then
        Call Dispose
    End If
    
End Sub
'-------------------------------------
' カスタムプロパティ
'-------------------------------------
Public Property Let MaxItems(ByVal m As Long)
    mlngMaxItem = m
End Property
