VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionSetUniq"
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
Option Explicit
'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 宣言
'------------------------------------------------------------------------------------------------------------------------
Private WithEvents SFWork As SelectionFrameWork
Attribute SFWork.VB_VarHelpID = -1

'------------------------------------------------------------------------------------------------------------------------
' メンバ変数宣言部(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private mlngMax As Long
Private mlngMatch As Long

Private mobjCol As Collection
Private mobjDel As Range

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New SelectionFrameWork
End Sub
'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 開放
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set SFWork = Nothing
End Sub
'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork Run
'------------------------------------------------------------------------------------------------------------------------
Public Sub Run()
    SFWork.Run
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 前処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionInit(Cancel As Boolean, Undo As Boolean, Func As Boolean)

    On Error GoTo e
    
    'メンバ変数の初期化
    Set mobjCol = New Collection
    mlngMax = 0
    mlngMatch = 0
    
    Undo = True
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal NotHoldFormat As Boolean, Cancel As Boolean)
        
    Dim c As Range
    Dim blnMatch As Boolean
    
    Dim i As Long
    Dim count As Long
    
    On Error GoTo e
    
    If r.value = "" Then
        Exit Sub
    End If

    'コレクション内に一致するものがあるかどうかチェック
    blnMatch = False
    
    Err.Clear
    On Error Resume Next
    Set c = Nothing
    Set c = mobjCol(CStr(r.value))
    
    On Error GoTo 0
            
    If Err.Number = 0 And Not (c Is Nothing) Then
        blnMatch = True
    End If
    
    '一致するものが１つでも存在した場合
    If blnMatch Then
        mlngMatch = mlngMatch + 1
        If mobjDel Is Nothing Then
            Set mobjDel = r
        Else
            Set mobjDel = Union(mobjDel, r)
        End If
    Else
        '存在しない場合コレクションに追加
        mlngMax = mlngMax + 1
        Call mobjCol.Add(r, CStr(r.value))
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
    Cancel = True
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 終了処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionTerm()
    
    On Error GoTo e
    
    '重複したセルを削除
    If mobjDel Is Nothing Then
    Else
        mobjDel.Delete xlUp
    End If

    If mlngMatch > 0 Then
        MsgBox mlngMatch & " 個 削除しました。確認してください。", vbExclamation, C_TITLE
    Else
        MsgBox "一致するものはありませんでした。", vbInformation, C_TITLE
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

