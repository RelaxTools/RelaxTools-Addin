VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionRowShiftSelect"
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
Private mSel As Range
Private mlngPos As Long

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

Private Sub SFWork_SelectionInit(rArea As Areas, Cancel As Boolean, Undo As Boolean)

    Dim strIns As String

    On Error GoTo e
    
    strIns = InputBox("選択範囲を左右シフトする行数を入力してください。" & vbCrLf & "（左シフト：マイナス、右シフト：整数）", C_TITLE)
    If strIns = "" Then
        Cancel = True
        Exit Sub
    End If
    
    mlngPos = Val(strIns)
    If mlngPos = 0 Then
        Cancel = True
        Exit Sub
    End If

    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub

Private Sub SFWork_SelectionBegin(rArea As Range, Cancel As Boolean)

    On Error GoTo e
    
    If mSel Is Nothing Then
        Set mSel = rArea.Offset(0, mlngPos)
    Else
        Set mSel = Union(mSel, rArea.Offset(0, mlngPos))
    End If

    Exit Sub
e:
    Call rlxErrMsg(err)
    Cancel = True
End Sub
Private Sub SFWork_SelectionTerm()

    On Error GoTo e
    
    If mSel Is Nothing Then
    Else
        mSel.Select
    End If

    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub
