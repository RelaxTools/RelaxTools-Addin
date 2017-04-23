VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacroManager"
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
Private mFrm As UserForm
Private mlngMax As Long
Private mstrBackup As String
Private mlngColor As Long
Private Sub Class_Initialize()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

End Sub

Private Sub Class_Terminate()

    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    mFrm.lblBack.Caption = mstrBackup
    mFrm.lblBack.BackColor = mlngColor


End Sub
Sub DispGuidance(ByVal strValue As String)

    If mFrm Is Nothing Then
        Exit Sub
    End If
    mFrm.lblBack.Caption = strValue
    DoEvents

End Sub
Sub StartGauge(ByVal lngValue As Long)
    mlngMax = lngValue
    
    mFrm.lblGauge.width = 0
    mFrm.lblGauge.visible = True
    mFrm.lblGauge.ForeColor = vbWhite
    mFrm.lblBack.BackColor = vbWhite

End Sub
Sub DisplayGauge(ByVal lngValue As Long)

    Dim dblValue As Double

    If mFrm Is Nothing Then
        Exit Sub
    End If
        Dim strMessage As String
    If lngValue > mlngMax Then
        lngValue = mlngMax
    End If
    dblValue = (CDbl(lngValue) / mlngMax)
    mFrm.lblGauge.width = mFrm.lblBack.width * dblValue
    
    strMessage = Space$(Fix(mFrm.lblBack.width * 0.16)) & Format$(rlxRoundDown(dblValue * 100, 0), "0") & "%"
    mFrm.lblBack.Caption = strMessage
    mFrm.lblGauge.Caption = strMessage
    DoEvents
    DoEvents
    DoEvents
End Sub
Property Set Form(frm As UserForm)
    Set mFrm = frm
    
    On Error Resume Next
    mstrBackup = mFrm.lblBack.Caption
    mlngColor = mFrm.lblBack.BackColor
    
End Property
Sub Disable()
    If mFrm Is Nothing Then
        Exit Sub
    End If

    Dim c As control
    
    For Each c In mFrm.Controls
    
        If c.Tag <> "e" Then
            c.enabled = False
        End If
    
    Next
    DoEvents

End Sub
Sub Enable()
    If mFrm Is Nothing Then
        Exit Sub
    End If

    Dim c As control
    
    For Each c In mFrm.Controls
    
        If c.Tag <> "e" Then
            c.enabled = True
        End If
    
    Next

End Sub

