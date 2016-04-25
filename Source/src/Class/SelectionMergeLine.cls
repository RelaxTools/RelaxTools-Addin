VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionMergeLine"
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
Private mr As Range
'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 宣言
'------------------------------------------------------------------------------------------------------------------------
Private WithEvents SFWork As SelectionColFramework
Attribute SFWork.VB_VarHelpID = -1

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New SelectionColFramework
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



Private Sub SFWork_SelectionInit(rArea As Areas, Cancel As Boolean, Undo As Boolean)
    Undo = True
    Application.DisplayAlerts = False
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------

Private Sub SFWork_SelectionMain(r As Range, ByVal col As Long, ByRef Cancel As Boolean)

    Dim s As Range
    Dim p As Range
    Dim blnFirst As Boolean
    
    On Error GoTo e
    
    For Each s In r
        
        If (Not IsEmpty(s.value)) And (Not IsError(s.value)) Then
            If mr Is Nothing Then
                Set mr = s
            Else
                If s.value = mr(mr.count).value Then
                    Set mr = Union(mr, s)
                Else
                    mr.Borders(xlInsideHorizontal).LineStyle = xlNone
                    blnFirst = True
                    For Each p In mr
                        If blnFirst Then
                            blnFirst = False
                        Else
                            p.Font.Color = vbWhite
                        End If
                    Next
                    Set mr = s
                End If
            End If
        End If
    Next
    
    Exit Sub
e:
    Call rlxErrMsg(err)
    Cancel = True
End Sub

Private Sub SFWork_SelectionTerm()
    Dim p As Range
    Dim blnFirst As Boolean
    
    On Error GoTo e
    
    mr.Borders(xlInsideHorizontal).LineStyle = xlNone
    blnFirst = True
    For Each p In mr
        If blnFirst Then
            blnFirst = False
        Else
            p.Font.Color = vbWhite
        End If
    Next
    Application.DisplayAlerts = True
    
    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub

