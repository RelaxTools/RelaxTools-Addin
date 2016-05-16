VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionRowSelectCell"
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
Private mstrSearch As String
Private mlngType As Long

Private mSel As Range

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
Private Sub SFWork_SelectionInit(rArea As Areas, Cancel As Boolean, Undo As Boolean)

    Dim ret As VBA.VbMsgBoxResult

    On Error GoTo e
    
    ret = frmSelectCell.Start(mstrSearch, mlngType)
    If ret = vbOK Then
    Else
        Cancel = True
        Exit Sub
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal row As Long, Cancel As Boolean)
        
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo e
    
    lngRow = row
    lngCol = 1
    
    Dim c As Range
    
    For Each c In r
    
        Select Case mlngType
            Case 0
                If mstrSearch = c.value Then
                    If mSel Is Nothing Then
                        Set mSel = c
                    Else
                        Set mSel = Union(mSel, c)
                    End If
                End If
            Case 1
                If mstrSearch <> c.value Then
                    If mSel Is Nothing Then
                        Set mSel = c
                    Else
                        Set mSel = Union(mSel, c)
                    End If
                End If
            Case Else
                If InStr(c.value, mstrSearch) Then
                    If mSel Is Nothing Then
                        Set mSel = c
                    Else
                        Set mSel = Union(mSel, c)
                    End If
                End If
        End Select
        
        lngCol = lngCol + 1
    Next
    
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
    
    If mSel Is Nothing Then
        MsgBox "選択対象のセルがありませんでした。", vbInformation, C_TITLE
    Else
        mSel.Select
        Set mSel = Nothing
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

