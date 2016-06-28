VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionFormatDec"
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
Private WithEvents SFWork As SelectionFormatFramework
Attribute SFWork.VB_VarHelpID = -1
Private mRange As Range

Public BoadersIndex As XlBordersIndex
Public LineStyle       As Variant
Public Weight As Variant

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New SelectionFormatFramework
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

Private Sub SFWork_SelectionInit(Cancel As Boolean, Undo As Boolean)

    Undo = True
   
    Dim lngCol As Long
    Dim lngMax As Long
    
    lngMax = (2 ^ Selection.Rows.count)
    
    If Selection(1).Column + lngMax - 1 > Application.Columns.count Then
    Else
        If Selection.Columns.count < lngMax Then
            If MsgBox("デシジョンを表現する列が足りません。拡張しますか？", vbOKCancel + vbQuestion, C_TITLE) = vbOK Then
                lngCol = (2 ^ Selection.Rows.count) - Selection.Columns.count
                Range(Selection, Selection.Offset(0, lngCol)).Select
            End If
        End If
    End If
    
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range)
        
   Dim lngRow As Long
    Dim lngCol As Long
    Dim sw As Long
    
    Dim lngSeed As Long
    
    For lngRow = 1 To r.Rows.count
    
        sw = 1
        lngSeed = 2 ^ (r.Rows.count - lngRow)
        
        For lngCol = 1 To r.Columns.count
    
            If sw Mod 2 = 1 Then
                r(lngRow, lngCol).value = "Y"
            Else
                r(lngRow, lngCol).value = "N"
            End If
            
            If (lngCol Mod lngSeed) = 0 Then
                sw = sw + 1
            End If
            
            '行の２乗以上はパターンが重複するのでパス
            If lngCol >= (2 ^ r.Rows.count) Then
                Exit For
            End If
                
        Next
    Next
    
End Sub






