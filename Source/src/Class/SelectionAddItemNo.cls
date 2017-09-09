VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionAddItemNo"
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
Private mstrItemName As String
Private mlngPos As Long

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
Private Sub SFWork_SelectionInit(Cancel As Boolean, Undo As Boolean, Func As Boolean)

    Undo = True
    mlngPos = Selection(1).Row
    
End Sub

Public Property Let ItemName(ByVal strBuf As String)
    mstrItemName = strBuf
End Property

Private Sub SFWork_SelectionMain(r As Range, ByVal NotHoldFormat As Boolean, Cancel As Boolean)

    Dim lngCnt As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnFind As Boolean

    Dim strBuf As String
    Dim lngLevel As Long

    On Error GoTo e
    
    lngCol = r.Column
    lngRow = r.Row - 1
    
    strBuf = ""
    blnFind = False

    For lngCnt = lngRow To mlngPos Step -1
    
        strBuf = Cells(lngCnt, lngCol).Value
        lngLevel = Cells(lngCnt, lngCol).IndentLevel

'        If lngLevel < r.IndentLevel Then
'            Exit For
'        End If

'        If lngLevel >= r.IndentLevel Then
            '箇条書きが存在する行の場合
            If rlxHasItemNo(strBuf, mstrItemName) Then
                blnFind = True
                Exit For
            End If
'        End If
    Next

   
    Dim strNewNum As String
    If blnFind Then
        '次の番号を振る
        strNewNum = rlxGetItemNext(strBuf, mstrItemName)
    Else
        '見つからない場合自分のレベルで最初の番号にする
        strNewNum = rlxGetItemNext("", mstrItemName)
    End If

    '文字列に変換
    r.NumberFormatLocal = "@"

    '段落番号の削除
    delItemNo r

    '段落番号を設定する
    setItemNo r, strNewNum
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
    Cancel = True
    
End Sub



