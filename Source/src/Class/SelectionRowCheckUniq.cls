VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionRowCheckUniq"
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
Private WithEvents SFWork As SelectionRowFrameWork
Attribute SFWork.VB_VarHelpID = -1

'------------------------------------------------------------------------------------------------------------------------
' メンバ変数宣言部(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private mlngMatch As Long
'Private mobjCol As Collection
Private mobjCol As Object

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New SelectionRowFrameWork
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
Private Sub SFWork_SelectionInit(ByRef rAreas As Excel.Areas, Cancel As Boolean, Undo As Boolean)
    
    On Error GoTo e
    
    If rAreas.count > 1 Then
        MsgBox "複数のエリアには対応していません。", vbExclamation, C_TITLE
        Cancel = True
        Exit Sub
    End If

    Set mobjCol = CreateObject("Scripting.Dictionary")
    mlngMatch = 0

    Undo = True

    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal lngRow As Long, Cancel As Boolean)

    Dim strkey As String
    Dim v As Range
    Dim blnMatch As Boolean
    Dim c As Object

    On Error GoTo e
    
    'キーの作成
    strkey = ""
    For Each v In r
        If IsError(v.value) Then
        Else
            strkey = strkey & vbTab & v.value
        End If
    Next

    If Replace(strkey, vbTab, "") = "" Then
        Exit Sub
    End If

    'コレクション内に一致するものがあるかどうかチェック
    blnMatch = False
    
    If mobjCol.exists(strkey) Then
        Set c = mobjCol.Item(strkey)
        blnMatch = True
        With c.Interior
            .Color = vbCyan
            .Pattern = xlSolid
        End With
    End If
    
    '一致するものが１つでも存在した場合
    If blnMatch Then
        mlngMatch = mlngMatch + 1
        With r.Interior
            .Color = vbYellow
            .Pattern = xlSolid
        End With
    Else
        '存在しない場合コレクションに追加
        Call mobjCol.Add(strkey, r)
    End If

    Exit Sub
e:
    Call rlxErrMsg(err)
    Cancel = True
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 終了処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionTerm()
    
    On Error GoTo e
    
    If mlngMatch > 0 Then
        MsgBox mlngMatch & " 個 一致しました。確認してください。", vbExclamation, C_TITLE
    Else
        MsgBox "一致するものはありませんでした。", vbInformation, C_TITLE
    End If
    
    'Collectionの開放
    Set mobjCol = Nothing

    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub

