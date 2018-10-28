VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionLineFeedInsert"
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
' 初期処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionInit(Cancel As Boolean, Undo As Boolean, func As Boolean)
    Undo = True
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal NotHoldFormat As Boolean, Cancel As Boolean)

    Dim s As String

    On Error GoTo e

    If IsError(r.Value) Then
        Exit Sub
    End If

    s = r.Value

    If Len(s) = 0 Then
        Exit Sub
    End If

    If Not (NotHoldFormat) And VarType(r.Value) = vbString Then
        Select Case r.Characters(r.Characters.Count, 1).Text
            Case vbCr, vbLf, vbCrLf
            Case Else
                r.Characters(r.Characters.Count + 1, 0).Insert vbLf
        End Select
    Else
        '最後の文字が改行か、それ以外か？
        Select Case Right$(s, 1)
            Case vbCr, vbLf, vbCrLf
            Case Else
                If Len(r.PrefixCharacter) > 0 Then
                    r.Value = r.PrefixCharacter & s & vbLf
                Else
                    r.Value = s & vbLf
                End If
        End Select
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
    Cancel = True
End Sub
