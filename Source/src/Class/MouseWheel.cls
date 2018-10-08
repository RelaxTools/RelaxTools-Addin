VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseWheel"
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
'--------------------------------------------------------------
' 以下URLを参考にしています
'--------------------------------------------------------------
' エクセルの学校
' http://www.excel.studio-kazu.jp/kw/20141022104627.html
'--------------------------------------------------------------
'--------------------------------------------------------------
' コンボやリストボックスでマウスホイールによるスクロール
' basMouseWheel.GetInstance でのみインスタンスを生成する
'--------------------------------------------------------------
Option Explicit

'メインイベント
Public Event WheelUp(ByRef obj As Object)
Public Event WheelDown(ByRef obj As Object)

#If VBA7 And Win64 Then
    Private hWnd As LongPtr
#Else
    Private hWnd As Long
#End If

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal obj As Object, hWnd As LongPtr) As LongPtr
#Else
    Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal obj As Object, hWnd As Long) As Long
#End If


Private mobj As Object
Public Property Set obj(s As Object)
    Set mobj = s
End Property

Public Sub Install(frm As UserForm)
    WindowFromAccessibleObject frm, hWnd
    basMouseWheel.Install Me, CStr(hWnd)
End Sub
Public Sub Uninstall()
    basMouseWheel.Uninstall CStr(hWnd)
End Sub

Public Sub MouseLLHookProc(ByVal mouseData As Long)

    If mobj Is Nothing Then
        Exit Sub
    End If
    
    Select Case mouseData
        Case Is > 0
            If mobj Is Nothing Then
            Else
                RaiseEvent WheelUp(mobj)
            End If
        Case Is < 0
            If mobj Is Nothing Then
            Else
                RaiseEvent WheelDown(mobj)
            End If
    End Select
    
End Sub




