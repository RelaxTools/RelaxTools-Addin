Attribute VB_Name = "basSushi"
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
Option Private Module
' 32-bit Function version.
' ドライブ名からネットワークドライブを取得
#If VBA7 And Win64 Then
    'VBA7 = Excel2010以降。赤くコンパイルエラーになって見えますが問題ありません。
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If
Private strBuf As String
Private strBuf2 As String
Private blnFlg As Boolean
Public mlngP As Long
Public mSushiEnable As Boolean

Dim mlngInterval As Long
Dim mlngSpeed As Long
Dim mstrValue As String


''--------------------------------------------------------------------
''  スシを流すの押下状態の取得
''--------------------------------------------------------------------
'Sub sushiPressed(control As IRibbonControl, ByRef returnValue)
'
'    returnValue = mSushiEnable
'
'End Sub
'--------------------------------------------------------------------
'  スシを流すの押下時イベント
'--------------------------------------------------------------------
'Sub sushiOnAction(control As IRibbonControl, pressed As Boolean)
Sub sushiOnAction(control As IRibbonControl)

    On Error GoTo e

    Select Case control.id
        Case "sushiShow"
            mSushiEnable = True
        Case "sushiStop"
            mSushiEnable = False
        Case "sushiSetting"
            frmSushiSetting.Show
            Exit Sub
    End Select

    Call RefreshRibbon

    If mSushiEnable Then

        mlngSpeed = Val(GetSetting(C_TITLE, "Sushi", "Speed", 8))
        mlngInterval = Val(GetSetting(C_TITLE, "Sushi", "Interval", 10))
        mstrValue = GetSetting(C_TITLE, "Sushi", "Show", "1")

        Application.OnTime Now, "SushiGoRound"
        
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
Public Sub StopSushi()
    mSushiEnable = False
End Sub
'--------------------------------------------------------------------
'  スシ設定のEnabled/Disabled
'--------------------------------------------------------------------
Sub getSushiEnabled(control As IRibbonControl, ByRef enabled)

    On Error GoTo e
    
    enabled = Not (mSushiEnable)

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'
'Private Sub showSushi()
'
'    strBuf = ""
'
'    Dim i As Long
'    Dim j As Long
'    Dim k As Long
'    Dim v As Long
'
'    k = 1
'    Do While mSushiEnable
'
'        v = Val(Mid(mstrValue, k, 1))
'        strBuf = ThisWorkbook.Worksheets("sushi").Cells(1, v).Value & strBuf
'
'        k = k + 1
'        If k > Len(mstrValue) Then
'            k = 1
'        End If
'
'        j = 0
'        Do While mSushiEnable And j < mlngInterval
'            For i = 1 To Fix(mlngSpeed / 10)
'                DoEvents
'                Sleep 10
'            Next
'            strBuf = Left("　" & strBuf, 180)
'            Application.StatusBar = strBuf
'            j = j + 1
'        Loop
'    Loop
'
'    Application.StatusBar = False
'
'End Sub

'Sub sushisetting()
'    frmSushiSetting.Show
'End Sub

Sub SushiGoRound()

    Dim c As New Collection
    Dim objMae As Object
    Dim f As Object
    Dim i As Long
    Dim k As Long
    Dim j As Long
    Dim lngTopfrgin As Long
    Dim lngLeftfrgin As Long
    Dim lngMove As Long
    Dim lngWait As Long

    Dim lngMaisu As Long

    lngMaisu = mlngInterval
'    Dim a As New Transparent
'    a.Init
    j = 1
    For i = 1 To lngMaisu
        Set f = New frmSushi
        If i = 1 Then
            f.Show
        End If
        f.Left = Application.Left - i * 40
        f.Top = Application.Top + Application.Height - 36
        f.Tag = "→"
        f.Neta = Mid(mstrValue, j, 1)
        j = j + 1
        If j > Len(mstrValue) Then
            j = 1
        End If
        c.Add f
    Next
    


    Do While mSushiEnable
        For i = 1 To c.count

            Set f = c(i)

            Sleep 10

            '移動量
            lngMove = mlngSpeed


            lngTopfrgin = 10
            lngLeftfrgin = 40

            '→
            If f.Tag = "→" Then
                f.Left = f.Left + lngMove
                f.Top = Application.Top + Application.Height - lngLeftfrgin

                If i <> c.count Then
                    If f.Left > Application.Left + 40 And c(i + 1).visible = False Then
                        c(i + 1).Show
                    End If
                End If

                DoEvents
                If f.Left > (Application.Left + Application.width - 36) Then
                    f.Tag = "↑"
                End If
            End If

            '↑
            If f.Tag = "↑" Then
                f.Top = f.Top - lngMove
                f.Left = Application.Left + Application.width - lngLeftfrgin
                DoEvents
                If f.Top < (Application.Top + lngTopfrgin) Then
                    f.Tag = "←"
                End If
            End If

            '←
            If f.Tag = "←" Then
                f.Left = f.Left - lngMove
                f.Top = Application.Top + lngTopfrgin
                DoEvents
                If f.Left < (Application.Left + lngTopfrgin) Then
                    f.Tag = "↓"
                End If
            End If

            '↓
            If f.Tag = "↓" Then
                f.Top = f.Top + lngMove
                f.Left = Application.Left + lngTopfrgin
                DoEvents
                If f.Top > (Application.Top + Application.Height - 50) Then
                    f.Tag = "→"
                End If
            End If

            If mSushiEnable = False Then
                Exit Do
            End If

        Next
    Loop
    For Each f In c
        Unload f
    Next
'    a.Term
End Sub
