VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCommonOption 
   Caption         =   "RelaxTools共通設定"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12105
   OleObjectBlob   =   "frmCommonOption.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCommonOption"
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
Private mblnSpin As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim lngType As Long
    
    
    If rlxIsNumber(txtSleep.Text) Then
        Select Case Val(txtSleep.Text)
            Case 0 To 500
            Case Else
                MsgBox "数値を入力してください。0～500ms", vbOKOnly + vbExclamation, C_TITLE
                txtSleep.SetFocus
                Exit Sub
        End Select
    Else
        MsgBox "数値を入力してください。0～500ms", vbOKOnly + vbExclamation, C_TITLE
        txtSleep.SetFocus
        Exit Sub
    End If
    
    Select Case True
        Case optDebugWindow.Value
            lngType = C_LOG_DEBUGWINDOW
        Case optLogfile.Value
            lngType = C_LOG_LOGFILE
        Case optAll.Value
            lngType = C_LOG_ALL
    End Select
    Call SaveSetting(C_TITLE, "Log", "LogType", lngType)
    Call SaveSetting(C_TITLE, "Log", "Level", cboLogLevel.ListIndex)

    Call SaveSetting(C_TITLE, "Option", "OnRepeat", chkOnRepeat.Value)
    Call SaveSetting(C_TITLE, "Option", "NotHoldFormat", chkNotHoldFormat.Value)
    Call SaveSetting(C_TITLE, "Option", "ClipboardSleep", txtSleep.Text)
    
    Call SaveSetting(C_TITLE, "Option", "ExitMode", chkExitMode.Value)
    
    Logger.Level = cboLogLevel.ListIndex
    
    Unload Me

End Sub

Private Sub cmdOpen_Click()

    On Error GoTo e

    With CreateObject("WScript.Shell")
        .Run (rlxGetAppDataFolder & "Log")
    End With
    
    Exit Sub
e:
    MsgBox "ログフォルダを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE

End Sub

Private Sub txtSleep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case &H30 To &H39
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()
    
    Dim strBuf As String
    
    chkOnRepeat.Value = CBool(GetSetting(C_TITLE, "Option", "OnRepeat", True))
    chkNotHoldFormat.Value = CBool(GetSetting(C_TITLE, "Option", "NotHoldFormat", False))
    chkExitMode.Value = CBool(GetSetting(C_TITLE, "Option", "ExitMode", False))
    
    strBuf = ""
    strBuf = strBuf & "・セルの最後に文字列挿入" & vbCrLf
    strBuf = strBuf & "・セルの先頭に文字列挿入" & vbCrLf
    strBuf = strBuf & "・セルのn文字目に文字列挿入" & vbCrLf
    strBuf = strBuf & "・セルの先頭からn文字削除" & vbCrLf
    strBuf = strBuf & "・セルの最後からn文字削除" & vbCrLf
    strBuf = strBuf & "・セルの指定文字から左側を削除" & vbCrLf
    strBuf = strBuf & "・セルの指定文字から右側を削除" & vbCrLf
    strBuf = strBuf & "・セルのn文字目以前削除" & vbCrLf
    strBuf = strBuf & "・セルのn文字目以降削除" & vbCrLf
    strBuf = strBuf & "・セルのすべての改行を削除" & vbCrLf
    strBuf = strBuf & "・右1桁削除" & vbCrLf
    strBuf = strBuf & "・左1桁削除" & vbCrLf
    strBuf = strBuf & "・セルの最後に改行を追加" & vbCrLf
    strBuf = strBuf & "・セルの最後の改行を削除" & vbCrLf
    strBuf = strBuf & "・セルの前後のスペース削除" & vbCrLf
    
    lblTaisho.Caption = strBuf
    
    cboLogLevel.AddItem "Trace"
    cboLogLevel.AddItem "Info"
    cboLogLevel.AddItem "Warn"
    cboLogLevel.AddItem "Fatal"
    cboLogLevel.AddItem "None"
    
    cboLogLevel.ListIndex = CLng(GetSetting(C_TITLE, "Log", "Level", LogLevel.Info))
    
    Dim lngType As Long
    lngType = CLng(GetSetting(C_TITLE, "Log", "LogType", C_LOG_LOGFILE))
    
    Select Case lngType
        Case C_LOG_DEBUGWINDOW
            optDebugWindow.Value = True
        Case C_LOG_LOGFILE
            optLogfile.Value = True
        Case C_LOG_ALL
            optAll.Value = True
    End Select
    
    strBuf = ""
    strBuf = strBuf & "NetBeansやSqlDeveloperなどのJavaアプリや" & vbCrLf
    strBuf = strBuf & "クリップボードを扱うツールを同時使用した際に" & vbCrLf
    strBuf = strBuf & "不安定になる場合があります。" & vbCrLf
    strBuf = strBuf & "数値を大きくすると誤動作が軽減します。0～500ms" & vbCrLf
    strBuf = strBuf & "大きくしすぎると処理スピードが遅くなるので注意。"
    lblSleep.Caption = strBuf
    
    txtSleep.Text = GetSetting(C_TITLE, "Option", "ClipboardSleep", 0)
    
End Sub
Private Sub spnSleep_SpinDown()
    If mblnSpin Then
        Exit Sub
    End If
    mblnSpin = True
    txtSleep.Text = spinDown(txtSleep.Text)
    mblnSpin = False
End Sub

Private Sub spnSleep_SpinUp()
    If mblnSpin Then
        Exit Sub
    End If
    mblnSpin = True
    txtSleep.Text = spinUp(txtSleep.Text)
    mblnSpin = False
End Sub
Private Function spinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long

    lngValue = Val(vntValue)
    lngValue = lngValue + 5
    If lngValue > 500 Then
        lngValue = 500
    End If
    spinUp = lngValue

End Function

Private Function spinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long

    lngValue = Val(vntValue)
    lngValue = lngValue - 5
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDown = lngValue

End Function
