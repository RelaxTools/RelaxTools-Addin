VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStyle 
   Caption         =   "スタイルの削除"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5265
   OleObjectBlob   =   "frmStyle.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1

Private Sub cmdAll_Click()

    Dim i As Long
    
    If MsgBox("すべてのスタイルを削除します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    For i = lstStyle.ListCount - 1 To 0 Step -1
        ActiveWorkbook.Styles(lstStyle.List(i)).Delete
        lstStyle.RemoveItem i
    Next
    
    If lstStyle.ListCount > 0 Then
        lstStyle.Selected(0) = True
    Else
        cmdAll.enabled = False
        cmdDel.enabled = False
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()

    Dim i As Long
    
    On Error Resume Next
    
    For i = lstStyle.ListCount - 1 To 0 Step -1
        If lstStyle.Selected(i) Then
            ActiveWorkbook.Styles(lstStyle.List(i)).Delete
            lstStyle.RemoveItem i
        End If
    Next
    
    If lstStyle.ListCount > 0 Then
        lstStyle.Selected(0) = True
    Else
        cmdAll.enabled = False
        cmdDel.enabled = False
    End If
    
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdSearch_Click()

    Dim s As Style
    
    On Error Resume Next
    
    lstStyle.Clear
    
    lblGauge.visible = False
    
    Dim mMm As MacroManager
    Dim lngCnt As Long
    
    Set mMm = New MacroManager
    Set mMm.Form = Me
    mMm.Disable
    
    mMm.StartGauge ActiveWorkbook.Styles.count
    lngCnt = 1
    
    For Each s In ActiveWorkbook.Styles
    
        If s.BuiltIn Then
        Else
            If chkStyle.Value Then
                If Not SearchStyle(s.NameLocal) Then
                    lstStyle.AddItem s.NameLocal
                End If
            Else
                lstStyle.AddItem s.NameLocal
            End If
        End If
        
        lngCnt = lngCnt + 1
        mMm.DisplayGauge lngCnt
    Next
    mMm.Enable
    Set mMm = Nothing
    
    lblGauge.visible = False
    
    If lstStyle.ListCount > 0 Then
        lstStyle.Selected(0) = True
        cmdAll.enabled = True
        cmdDel.enabled = True
    Else
        cmdAll.enabled = False
        cmdDel.enabled = False
    End If
    
End Sub

Private Sub lstStyle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Set MW.obj = lstStyle
End Sub

Private Sub UserForm_Initialize()

    lblGauge.visible = False
    
    Call cmdSearch_Click
    
    Set MW = basMouseWheel.GetInstance
    MW.Install
End Sub
Function SearchStyle(ByVal strBuf As String) As Boolean

    Dim WS As Worksheet
    Dim r As Range
    
    SearchStyle = False
    
    For Each WS In ActiveWorkbook.Worksheets
        For Each r In WS.UsedRange
            If r.Style = strBuf Then
                SearchStyle = True
                Exit Function
            End If
        Next
    Next

End Function

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Set MW.obj = Nothing
End Sub
Private Sub MW_WheelDown(obj As Object)

    If obj.ListCount = 0 Then Exit Sub
    obj.TopIndex = obj.TopIndex + 3
    
End Sub

Private Sub MW_WheelUp(obj As Object)

    Dim lngPos As Long

    If obj.ListCount = 0 Then Exit Sub
    lngPos = obj.TopIndex - 3

    If lngPos < 0 Then
        lngPos = 0
    End If

    obj.TopIndex = lngPos

End Sub
Private Sub UserForm_Terminate()
    MW.UnInstall
    Set MW = Nothing
End Sub
