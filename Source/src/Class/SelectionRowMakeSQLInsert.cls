VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionRowMakeSQLInsert"
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
Private mInsert As String
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
Private Sub SFWork_SelectionInit(rAreas As Excel.Areas, Cancel As Boolean, Undo As Boolean)

    On Error GoTo e
    
    If rAreas.count > 1 Then
        MsgBox "複数のエリアには対応していません。", vbExclamation, C_TITLE
        Cancel = True
        Exit Sub
    End If
    
    Dim strTable As String
    strTable = InputBox("INSERT文に追加するテーブル名を入力してください。")
    If strTable = "" Then
        Cancel = True
    End If

    mInsert = "=""INSERT INTO " & strTable & " VALUES("

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal lngRow As Long, Cancel As Boolean)
        
    Dim v As Range
    Dim last As Range
    Dim out As Range
    Dim strSQL As String
    
    Dim strComma As String
    Dim strQuat As String
    
    On Error GoTo e
    
    'キーの作成
    strSQL = ""
    For Each v In r
    
        If strSQL = "" Then
            strComma = ""
        Else
            strComma = ","
        End If
        
        If v.NumberFormatLocal = "@" Then
            strQuat = "'"
        Else
            strQuat = ""
        End If
        
        strSQL = strSQL & strComma & "" & strQuat & """ & " & v.Address(False, False, xlA1, False) & " & """ & strQuat & ""
        
    Next
    
    Set last = r.Columns(r.Columns.count)
    Set out = r.Worksheet.Cells(last.row, last.Column + 1)

    out.value = mInsert & strSQL & ");"""

    Exit Sub
e:
    Call rlxErrMsg(Err)
    Cancel = True
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 終了処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionTerm()
    

End Sub




