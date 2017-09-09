VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionInsertMiddle"
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
' SelectionAllFrameWork 宣言
'------------------------------------------------------------------------------------------------------------------------
Private WithEvents SFWork As SelectionFrameWork
Attribute SFWork.VB_VarHelpID = -1

'------------------------------------------------------------------------------------------------------------------------
' メンバ変数宣言部(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private mstrIns As String
'Private mstrPos As String
Private mlngIns As Long


'------------------------------------------------------------------------------------------------------------------------
' SelectionAllFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New SelectionFrameWork
End Sub

'------------------------------------------------------------------------------------------------------------------------
' SelectionAllFrameWork 開放
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set SFWork = Nothing
End Sub

'------------------------------------------------------------------------------------------------------------------------
' SelectionAllFrameWork Run
'------------------------------------------------------------------------------------------------------------------------
Public Sub Run()
    SFWork.Run
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 前処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionInit(Cancel As Boolean, Undo As Boolean, Func As Boolean)

    mlngIns = frmInputLength.Start("セルの何文字目の前に文字を挿入するかを入力してください。")
    If mlngIns = 0 Then
        Cancel = True
        Exit Sub
    End If

    mstrIns = frmInputStr.Start()
    If mstrIns = "" Then
        Cancel = True
        Exit Sub
    End If

    Undo = True

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal NotHoldFormat As Boolean, Cancel As Boolean)

    Dim f As String
    Dim b As String

    Dim lngPos As Long

    On Error GoTo e

    If IsError(r.Value) Then
        Exit Sub
    End If

    lngPos = rlxSetLimit(1, Len(r.Value) + 1, mlngIns)

    If Not (NotHoldFormat) And VarType(r.Value) = vbString Then
        r.Characters(lngPos, 0).Insert mstrIns
    Else

        f = Mid$(r.Value, 1, lngPos - 1)
        b = Mid$(r.Value, lngPos)

        If Len(r.PrefixCharacter) > 0 Then
            r.Value = r.PrefixCharacter & f & mstrIns & b
        Else
            r.Value = f & mstrIns & b
        End If
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
    Cancel = True

End Sub
