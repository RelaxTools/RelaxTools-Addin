VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeDrawEvidenceArrow"
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
Private WithEvents SFWork As ShapeFramework
Attribute SFWork.VB_VarHelpID = -1

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New ShapeFramework
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


Private Sub SFWork_SelectionMain(obj As Object, pos As ClickPosition, margin As Long, Cancel As Boolean)
    
    Dim c As Long
    Dim r As Long
    
    c = ActiveWindow.ActivePane.ScrollColumn
    r = ActiveWindow.ActivePane.ScrollRow
    
    Set obj = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, Selection.Left + (Selection.width / 2), Selection.Top + Selection.Height - (C_NORMAL_HEIGHT * 25), Selection.Left + (Selection.width / 2), Selection.Top + Selection.Height)
    
    With obj
        
        With obj.Line
            .Weight = 2.25
            .DashStyle = msoLineSingle
            .Style = msoLineSingle
            .Transparency = 0#
            .visible = msoTrue
            .ForeColor.SchemeColor = 10
            .BackColor.RGB = RGB(255, 255, 255)
            .EndArrowheadStyle = msoArrowheadOpen
            .EndArrowheadLength = msoArrowheadLong
            .EndArrowheadWidth = msoArrowheadWide
        End With
            
        .width = Selection.Left + (Selection.width / 2)
        .Height = Selection.Top + Selection.Height
            
        obj.Select
        obj.Placement = xlMove

    End With
    
    If CBool(GetSetting(C_TITLE, "Shape", "PickShape", True)) = False Then
        Cancel = True
    End If
    
End Sub




