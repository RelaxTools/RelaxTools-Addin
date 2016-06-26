VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionFormatCheckList"
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
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range)
        
    Dim bLine As Variant
    Dim bStyle As Variant

    '線の太さ
    bStyle = xlContinuous
    bLine = xlHairline
    
    Select Case Val(GetSetting(C_TITLE, "CheckList", "LineNo", 2))
        Case 1
            bLine = xlHairline
            bStyle = xlLineStyleNone
        Case 2
            bLine = xlHairline
            bStyle = xlContinuous
        Case 3
            bLine = xlThin
            bStyle = xlDot
        Case 4
            bLine = xlThin
            bStyle = xlDashDotDot
        Case 5
            bLine = xlThin
            bStyle = xlDashDot
        Case 6
            bLine = xlThin
            bStyle = xlDash
        Case 7
            bLine = xlThin
            bStyle = xlContinuous
        Case 8
            bLine = xlMedium
            bStyle = xlDashDotDot
        Case 9
            bLine = xlMedium
            bStyle = xlSlantDashDot
        Case 10
            bLine = xlMedium
            bStyle = xlDashDot
        Case 11
            bLine = xlMedium
            bStyle = xlDash
        Case 12
            bLine = xlMedium
            bStyle = xlContinuous
        Case 13
            bLine = xlThick
            bStyle = xlContinuous
        Case 14
            bLine = xlThick
            bStyle = xlDouble
    End Select
    
    
    
    

'    外周に線を引く (おこのみでコメント)
    If CBool(GetSetting(C_TITLE, "CheckList", "Left", True)) Then
        With r.Borders(xlEdgeLeft)
            .Weight = bLine
            .LineStyle = bStyle
        End With
    End If
    
    If CBool(GetSetting(C_TITLE, "CheckList", "Head", True)) Then
        With r.Borders(xlEdgeTop)
            .Weight = bLine
            .LineStyle = bStyle
        End With
    End If
    
    If CBool(GetSetting(C_TITLE, "CheckList", "Bottom", True)) Then
        With r.Borders(xlEdgeBottom)
            .Weight = bLine
            .LineStyle = bStyle
        End With
    End If
    
    If CBool(GetSetting(C_TITLE, "CheckList", "Right", True)) Then
        With r.Borders(xlEdgeRight)
            .Weight = bLine
            .LineStyle = bStyle
        End With
    End If

    '選択エリア全体に横中線を引く
    With r.Borders(xlInsideHorizontal)
        .Weight = bLine
        .LineStyle = bStyle
    End With
    
    '選択エリア全体に縦中線を消す
    r.Borders(xlInsideVertical).LineStyle = xlNone
    
    '１列目を選択
    Dim s As Range
    Set s = Application.Range(Selection(1), r(1).Offset(r.Rows.count - 1, 0))

    '１列目の横中線を削除、右側に線を引く
    If s Is Nothing Then
    Else
        s.Borders(xlInsideHorizontal).LineStyle = xlNone
        
        With s.Borders(xlEdgeRight)
            .Weight = bLine
            .LineStyle = bStyle
        End With
        
    End If

    '左上のセルの右側だけ線を削除
    r(1).Borders(xlEdgeRight).LineStyle = xlNone
    
End Sub




