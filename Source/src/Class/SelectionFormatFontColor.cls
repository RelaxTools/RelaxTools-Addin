VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionFormatFontColor"
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
' SelectionFrameWork êÈåæ
'------------------------------------------------------------------------------------------------------------------------
Private WithEvents SFWork As SelectionFormatFramework
Attribute SFWork.VB_VarHelpID = -1
Private Const C_COLOR_OTHER As String = "99"

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork çÏê¨
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New SelectionFormatFramework
End Sub

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork äJï˙
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
' éÂèàóù(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range)
        
    Dim s As String
  
    On Error GoTo e
    
    If r Is Nothing Then
    Else
        Dim strColorNo As String
        strColorNo = GetSetting(C_TITLE, "Color2003", "font", C_COLOR_OTHER)
        
        If strColorNo = C_COLOR_OTHER Then
            With r.Font
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
            End With
        Else
            r.Font.Color = ThisWorkbook.Colors(Val(strColorNo))
        End If
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub

'------------------------------------------------------------------------------------------------------------------------
' ShapeópéÂèàóù(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMainShape(r As Shape, Cancel As Boolean)

    On Error GoTo e
    
    If r.TextFrame2.HasText Then
    
        Dim strColorNo As String
        strColorNo = GetSetting(C_TITLE, "Color2003", "font", C_COLOR_OTHER)
        
        If strColorNo = C_COLOR_OTHER Then
            r.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        Else
            r.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ThisWorkbook.Colors(Val(strColorNo))
        End If
    
    End If
    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub

