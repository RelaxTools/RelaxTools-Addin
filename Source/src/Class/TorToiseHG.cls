VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TorToiseHG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [RelaxTools-Addin] v4
'
' Copyright (c) 2009 Yasuhiro Watanabe
' https://HGhub.com/RelaxTools/RelaxTools-Addin
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

Implements IVersion

Private Const EXE_NAME As String = "thgw"

Private Const CMD_ADD As String = "add "
Private Const CMD_COMMIT As String = "commit "
Private Const CMD_DIFF As String = "vdiff "
Private Const CMD_LOG As String = "log "
Private Const CMD_REVERT As String = "revert "
Private Const CMD_BROUSER As String = " "
Private Const CMD_ABOUT As String = "about "
Private Const CMD_HELP As String = "help "

Private Sub IVersion_Add()

    Dim strCommand As String
    Dim strPath As String
    Dim strFile As String
    Dim WB As Workbook
    Dim strBook As String
    
'    On Error GoTo e
    
    Set WB = ActiveWorkbook
    
    If WB Is Nothing Then
        Exit Sub
    End If
    
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "ブックが保存されていません。" & vbCrLf & "ブックを保存してから再実行してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
'    If MsgBox("ブックをバージョン管理システムに追加します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
'        Exit Sub
'    End If

    strPath = rlxGetFullpathFromPathName(WB.FullName)
    strCommand = CMD_ADD & GetFile(WB.FullName)
    
    Run strPath, strCommand
    
'    MsgBox "追加しました。", vbOKOnly + vbExclamation, C_TITLE
    
    Exit Sub
e:
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Brouser()

    Dim strCommand As String
    Dim strPath As String
    Dim strFile As String
    Dim WB As Workbook
    Dim strBook As String
    
'    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    If IsError(WB) Then
        Exit Sub
    End If
    
    strPath = rlxGetFullpathFromPathName(WB.FullName)
    strCommand = ""
    
    Run strPath, strCommand
    
    Exit Sub
e:
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Cleanup()

End Sub

Private Sub IVersion_Commit()

    Dim strCommand As String
    Dim strPath As String
    Dim strFile As String
    Dim WB As Workbook
    
'    On Error GoTo e
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    strPath = rlxGetFullpathFromPathName(WB.FullName)
    
    strCommand = CMD_COMMIT & GetFile(WB.FullName)
    Run strPath, strCommand


    Exit Sub
e:
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Diff()

    Dim strCommand As String
    Dim strPath As String
    Dim strFile As String
    Dim WB As Workbook
    
'    On Error GoTo e
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    strPath = rlxGetFullpathFromPathName(WB.FullName)
    strCommand = CMD_DIFF & GetFile(WB.FullName)
    Run strPath, strCommand

    Exit Sub
e:
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Help()

    Dim strCommand As String
    Dim strPath As String
    Dim strFile As String
    Dim WB As Workbook
    
'    On Error GoTo e
    
    Set WB = ActiveWorkbook
    
    strPath = rlxGetFullpathFromPathName(WB.FullName)
    strCommand = CMD_HELP
    Run strPath, strCommand

    
    Exit Sub
e:
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Locked()

End Sub

Private Sub IVersion_Pull()

End Sub

Private Sub IVersion_Push()

End Sub

Private Sub IVersion_RevisionGraph()
    
End Sub

Private Sub IVersion_Tag()

End Sub

Private Sub IVersion_Unlocked()

End Sub
Private Sub IVersion_Log()

    Dim strCommand As String
    Dim strPath As String
    Dim strFile As String
    Dim WB As Workbook
    Dim strBook As String
    
'    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    If IsError(WB) Then
        Exit Sub
    End If
    
'    On Error Resume Next
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    WB.ChangeFileAccess Mode:=xlReadOnly
    
    strPath = rlxGetFullpathFromPathName(WB.FullName)
    strCommand = CMD_LOG & GetFile(WB.FullName)
    Run strPath, strCommand


    WB.ChangeFileAccess Mode:=xlReadWrite
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Revert()
    
    Dim strCommand As String
    Dim strPath As String
    Dim strFile As String
    Dim WB As Workbook
    
'    On Error GoTo e
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
'    On Error Resume Next
        
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    WB.ChangeFileAccess Mode:=xlReadOnly
    
    strPath = rlxGetFullpathFromPathName(WB.FullName)
    strCommand = CMD_REVERT & GetFile(WB.FullName)
    Run strPath, strCommand

    WB.ChangeFileAccess Mode:=xlReadWrite
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub


Private Sub IVersion_Update()
    
End Sub
Private Sub IVersion_Ver()

    Dim strCommand As String
    Dim strPath As String
    Dim strFile As String
    Dim WB As Workbook
    
'    On Error GoTo e
    
    Set WB = ActiveWorkbook
    
    strPath = rlxGetFullpathFromPathName(WB.FullName)
    strCommand = CMD_ABOUT
    Run strPath, strCommand

    
    Exit Sub
e:
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub Run(ByVal strPath As String, ByVal strExe As String)
    On Error Resume Next

    err.Clear
    ChDir strPath
    With CreateObject("WScript.Shell")
        .Run EXE_NAME & " " & strExe, 1, True
    End With
    If err.Number <> 0 Then
        MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
    End If
    
    Exit Sub
e:
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Function IsError(WB As Workbook, Optional ByVal Upd As Boolean = False) As Boolean

    Dim strBook As String
    Dim ret As Boolean
    
    IsError = True

    If WB Is Nothing Then
        Exit Function
    End If
    
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "ブックが保存されていません。" & vbCrLf & "ブックを保存してから再実行してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Function
    End If
    
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。保存しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            MsgBox "処理を中断しました。", vbOKOnly + vbExclamation, C_TITLE
            Exit Function
        Else
            Application.DisplayAlerts = False
            WB.Save
            Application.DisplayAlerts = True
            WB.Saved = True
        End If
    End If
    
    IsError = False
    
End Function


Private Function GetFile(ByVal strFullName As String) As String

    GetFile = """" & rlxGetFullpathFromFileName(strFullName) & """"

End Function


