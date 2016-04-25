VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionSetClipboardCopy"
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
Private mstrBuf As String
Private mlngErr As Long

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
Private Sub SFWork_SelectionInit(Cancel As Boolean, Undo As Boolean, Func As Boolean)

    Undo = True
    mstrBuf = ""
    mlngErr = 0
    
End Sub

'------------------------------------------------------------------------------------------------------------------------
' メイン処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal NotHoldFormat As Boolean, Cancel As Boolean)

    On Error GoTo e
    
    If rlxIsFileExists(r.value) Or rlxIsFolderExists(r.value) Then
        If mstrBuf = "" Then
            mstrBuf = r.value
        Else
            mstrBuf = mstrBuf & vbTab & r.value
        End If
    Else
        r.Interior.Color = vbYellow
        mlngErr = mlngErr + 1
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(err)
    Cancel = True
End Sub
Private Sub SFWork_SelectionTerm()
    
    Dim strFiles() As String
        
    On Error GoTo e
    
    If mlngErr > 0 Then
        MsgBox "存在しないファイル/フォルダが" & mlngErr & "件の存在します。" & vbCrLf & "クリップボードへのコピーを中断しました。", vbOKOnly + vbExclamation, C_TITLE
    Else
        If mstrBuf <> "" Then
            strFiles = Split(mstrBuf, vbTab)
            SetCopyClipText strFiles
            MsgBox UBound(strFiles) + 1 & "件のファイル/フォルダをクリップボードにコピーしました。", vbOKOnly + vbInformation, C_TITLE
        End If
     End If

    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub
