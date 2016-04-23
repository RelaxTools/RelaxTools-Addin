VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionRowSelect"
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
Private mlngRowSel As Long
Private mlngRowGap As Long
Private mlngColSel As Long
Private mlngColGap As Long

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
    
    'メンバ変数の初期化
    mlngRowSel = 1
    mlngRowGap = 1
    mlngColSel = 1
    mlngColGap = 1

    ret = frmSelect.Start(mlngRowSel, mlngRowGap, mlngColSel, mlngColGap)
    If ret = vbOK Then
    Else
        Cancel = True
        Exit Sub
    End If

    Exit Sub
e:
    Call rlxErrMsg(err)
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
    
        '選択行判定
        If isSelect(lngRow, lngCol) Then
            If mSel Is Nothing Then
                Set mSel = c
            Else
                Set mSel = Union(mSel, c)
            End If
        End If
        lngCol = lngCol + 1
    Next
    
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
    
    If mSel Is Nothing Then
    Else
        mSel.Select
        Set mSel = Nothing
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub


Private Function isSelect(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean

    Dim lngRowLeap As Long
    Dim lngColLeap As Long
    
    Dim blnRowSel As Boolean
    Dim blnColSel As Boolean

    isSelect = False

    lngRowLeap = mlngRowSel + mlngRowGap
    lngColLeap = mlngColSel + mlngColGap

    If lngRowLeap = 0 Then
        blnRowSel = True
    Else
        Select Case (lngRow Mod lngRowLeap)
            Case 1 To mlngRowSel
                blnRowSel = True
            Case Else
                blnRowSel = False
        End Select
    End If
    
    If lngColLeap = 0 Then
        blnColSel = True
    Else
        Select Case (lngCol Mod lngColLeap)
            Case 1 To mlngColSel
                blnColSel = True
            Case Else
                blnColSel = False
        End Select
    End If
    '行、列共に選択の場合にのみ選択となる。
    If blnRowSel And blnColSel Then
        isSelect = True
    End If

End Function
