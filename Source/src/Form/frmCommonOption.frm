VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCommonOption 
   Caption         =   "RelaxTools共通設定"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim lngType As Long
    
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

Private Sub UserForm_Initialize()
    
    Dim strBuf As String
    
    chkOnRepeat.Value = CBool(GetSetting(C_TITLE, "Option", "OnRepeat", True))
    chkNotHoldFormat.Value = CBool(GetSetting(C_TITLE, "Option", "NotHoldFormat", False))
    
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
    
End Sub
