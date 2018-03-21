VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TortoiseSVN"
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

Implements IVersion

Private Const EXE_NAME As String = "TortoiseProc.exe"
Private Const OPT_COMMAND As String = "/command:"
Private Const OPT_PATH As String = "/path:"
Private Const OPT_END0 As String = "/closeonend:0"
Private Const OPT_END1 As String = "/closeonend:1"

Private Const CMD_ADD As String = "add "
Private Const CMD_COMMIT As String = "commit "
Private Const CMD_DIFF As String = "diff "
Private Const CMD_LOCK As String = "lock "
Private Const CMD_UNLOCK As String = "unlock "
Private Const CMD_LOG As String = "log "
Private Const CMD_UPDATE As String = "update "
Private Const CMD_REVERT As String = "revert "
Private Const CMD_BROUSER As String = "repobrowser "
Private Const CMD_CLEANUP As String = "cleanup "
Private Const CMD_ABOUT As String = "about "
Private Const CMD_HELP As String = "help "

Private Sub IVersion_Add()

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    
    If WB Is Nothing Then
        Exit Sub
    End If
    
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 0, 0
        If .IsSvnItem Then
            MsgBox "すでにバージョン管理に登録されているファイルです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
    
    If MsgBox("ブックをバージョン管理システムに追加します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    strCommand = CMD_ADD & GetPath(WB.FullName) & OPT_END0
    Run strCommand
    
    MsgBox "追加しました。", vbOKOnly + vbExclamation, C_TITLE
    
    Exit Sub
e:
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Brouser()

    Dim strCommand As String
    Dim WB As Workbook
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    WB.ChangeFileAccess Mode:=xlReadOnly
    
    strCommand = CMD_BROUSER & GetPath(WB.FullName) & OPT_END0
    Run strCommand
    
    

    WB.ChangeFileAccess Mode:=xlReadWrite
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Cleanup()

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_CLEANUP & "/PATH:" & rlxGetFullpathFromPathName(strBook) & " " & OPT_END0
    If Run(strCommand) < 0 Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    WB.Close False

    Workbooks.Open strBook
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Commit()

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_COMMIT & GetPath(WB.FullName) & OPT_END0
    If Run(strCommand) < 0 Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    WB.Close False
    Set WB = Nothing
    
    Set WB = Workbooks.Open(strBook)
    WB.Windows(1).visible = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    WB.Windows(1).visible = True
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Diff()

    Dim strCommand As String
    Dim WB As Workbook
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_DIFF & GetPath(WB.FullName) & OPT_END0
    Run strCommand

    Exit Sub
e:
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Help()

    Dim strCommand As String
    Dim WB As Workbook
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strCommand = CMD_HELP
    Run strCommand
    
    Exit Sub
e:
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Locked()

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 0, 0
        If .IsSvnItem Then
        Else
            MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo WB.FullName, 0, 0
        If .IsLocked Then
            MsgBox "ブックはロック済みです。", vbOKCancel + vbQuestion, C_TITLE
            Exit Sub
        End If
    End With
    
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。破棄しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    Else
        If MsgBox("ロックします。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    CreateObject("Scripting.FileSystemObject").GetFile(WB.FullName).Attributes = 0
    strCommand = CMD_LOCK & GetPath(WB.FullName) & OPT_END0
    WB.Close False
    DoEvents
    Set WB = Nothing
    
    Run strCommand
    
    Set WB = Workbooks.Open(strBook)
    WB.Windows(1).visible = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    WB.Windows(1).visible = True
    Exit Sub
e:
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
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

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 0, 0
        If .IsSvnItem Then
        Else
            MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo WB.FullName, 0, 0
        If Not .IsLocked Then
            MsgBox "ブックはロックされていません。", vbOKCancel + vbQuestion, C_TITLE
            Exit Sub
        End If
    End With
    
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。破棄しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    End If
    
    On Error Resume Next
        
    strCommand = CMD_UNLOCK & GetPath(WB.FullName) & OPT_END0
    If Run(strCommand) < 0 Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    WB.Close False
    Set WB = Nothing
    
    
    Set WB = Workbooks.Open(strBook)
    WB.Windows(1).visible = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    WB.Windows(1).visible = True
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub
Private Sub IVersion_Log()

    Dim strCommand As String
    Dim WB As Workbook
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    WB.ChangeFileAccess Mode:=xlReadOnly
    
    strCommand = CMD_LOG & GetPath(WB.FullName) & OPT_END0
    Run strCommand

    WB.ChangeFileAccess Mode:=xlReadWrite
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Revert()
    
    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 0, 0
        If .IsSvnItem Then
        Else
            MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
        
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。破棄しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    Else
        If MsgBox("変更を元に戻します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strCommand = CMD_REVERT & GetPath(WB.FullName) & OPT_END0
    WB.Close False
    DoEvents
    Set WB = Nothing
    
    Run strCommand
    Set WB = Workbooks.Open(strBook)
    WB.Windows(1).visible = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    WB.Windows(1).visible = True
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub


Private Sub IVersion_Update()
    
    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 0, 0
        If .IsSvnItem Then
        Else
            MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
        
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。破棄しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    Else
        If MsgBox("HEADに更新します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    End If
    
    On Error Resume Next

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strCommand = CMD_UPDATE & GetPath(WB.FullName) & OPT_END0
    WB.Close False
    DoEvents
    Set WB = Nothing
    
    Run strCommand
    Set WB = Workbooks.Open(strBook)
    WB.Windows(1).visible = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    WB.Windows(1).visible = True
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub
Private Sub IVersion_Ver()

    Dim strCommand As String
    Dim WB As Workbook
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strCommand = CMD_ABOUT
    Run strCommand
    
    Exit Sub
e:
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Function Run(ByVal strExe As String)
    On Error Resume Next

    Err.Clear
    With CreateObject("WScript.Shell")
        Run = .Run(EXE_NAME & " " & OPT_COMMAND & strExe, 1, True)
    End With
    If Err.Number <> 0 Then
        MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
    End If
    
    Exit Function
e:
    MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Function
Private Function GetPath(ByVal strBook As String) As String
    
    GetPath = OPT_PATH & """" & strBook & """ "

End Function
Private Function IsError(WB As Workbook, Optional ByVal Upd As Boolean = False) As Boolean

    Dim strBook As String
    Dim ret As Boolean
    
    IsError = True

    If WB Is Nothing Then
        Exit Function
    End If
    
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
        Exit Function
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 0, 0
        If .IsSvnItem Then
        Else
            MsgBox "バージョン管理外のブックです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Function
        End If
        
        If Upd And (.HasModifications Or Not WB.Saved) Then
            MsgBox "ブックは変更されているため、更新することができません。", vbOKOnly + vbExclamation, C_TITLE
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
    
    End With
    
    IsError = False
    
End Function

