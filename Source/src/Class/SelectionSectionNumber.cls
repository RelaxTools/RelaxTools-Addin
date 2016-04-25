VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionSectionNumber"
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
Private mlngIndent As Long


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
    
End Sub

Public Property Let Indent(ByVal lngBuf As Long)
    mlngIndent = lngBuf
End Property

Private Sub SFWork_SelectionMain(r As Range, ByVal NotHoldFormat As Boolean, Cancel As Boolean)

    Dim lngCnt As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnFind As Boolean
    Dim blnSearch As Boolean

    Dim strBuf As String
    Dim lngLevel As Long

    Dim WB As Workbook
    Dim WS As Worksheet
    Dim i As Long
    Dim j As Long
    Const C_NOT_FOUND As Long = -1

    On Error GoTo e
    
    lngCol = r.Column
    lngRow = C_NOT_FOUND 'r.row - 1
    
    strBuf = ""
    lngLevel = 0
    blnFind = False
    blnSearch = False

    setIndent r, mlngIndent
    
    '現在のセルの段落番号である場合
    If Not rlxIsSectionNo(r.value) Then
        Exit Sub
    End If
'    Application.ScreenUpdating = False
    
    Set WB = r.Parent.Parent
    
    j = 0
    For i = r.Parent.Index To 1 Step -1
    
        '表示されているシートのみで判定
        Set WS = WB.Worksheets(i)
        
        If WS.visible = xlSheetVisible Then
        
            If lngRow = C_NOT_FOUND Then
                lngRow = r.row - 1
            Else
                lngRow = WS.UsedRange.Item(WS.UsedRange.count).row
            End If
            
            For lngCnt = lngRow To 1 Step -1
            
                strBuf = WS.Cells(lngCnt, lngCol).value
                lngLevel = WS.Cells(lngCnt, lngCol).IndentLevel
        
                '段落番号が存在しない行の場合
                If (rlxHasSectionNo(strBuf, lngLevel) And lngLevel <= r.IndentLevel) Then
                    blnFind = True
                    GoTo pass
                End If
            
            Next
            
            blnSearch = True
            j = j + 1
            If j >= 2 Then
                Exit For
            End If
        End If
        
        Set WS = Nothing
        
    Next
    Set WB = Nothing
pass:
    
    If blnSearch And blnFind Then
    
        Dim strMsg As String
    
        strMsg = ""
        strMsg = strMsg & "１つ前のシートから段落番号を引き継ぎますか？" & vbCrLf
        strMsg = strMsg & "引き継がない場合は最初から採番します。" & vbCrLf & vbCrLf
        strMsg = strMsg & "☆制限☆" & vbCrLf
        strMsg = strMsg & "　引き継ぐためには同じ列に段落番号が記述されている必要があります。"
        
        'メッセージ
        If MsgBox(strMsg, vbQuestion + vbYesNo, C_TITLE) <> vbYes Then
            strBuf = ""
        End If
            
    End If

    Dim strNewNum As String
        
    '次の番号を振る
    strNewNum = rlxGetSectionNext(strBuf, lngLevel, r.IndentLevel)
    
    '段落番号の削除
    delSectionNo r
    
    '段落番号を設定する
    setSectionNo r, strNewNum
    
'    Application.ScreenUpdating = True
    
    Exit Sub
e:
'    Application.ScreenUpdating = True
    Call rlxErrMsg(err)
    Cancel = True
    
End Sub


