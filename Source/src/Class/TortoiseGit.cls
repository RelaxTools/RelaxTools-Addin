VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TortoiseGit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVersion

Private Const EXE_NAME As String = "TortoiseGitProc.exe"
Private Const OPT_COMMAND As String = "/command:"
Private Const OPT_PATH As String = "/path:"
Private Const OPT_END0 As String = "/closeonend:0"
Private Const OPT_END1 As String = "/closeonend:1"

Private Const CMD_ADD As String = "add "
Private Const CMD_COMMIT As String = "commit "
Private Const CMD_DIFF As String = "diff "
Private Const CMD_LOG As String = "log "
Private Const CMD_TAG As String = "tag "
Private Const CMD_REVISIONGRAPH As String = "revisiongraph "
Private Const CMD_PUSH As String = "push "
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
        MsgBox "ブックが保存されていません。" & vbCrLf & "ブックを保存してから再実行してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    If MsgBox("ブックをバージョン管理システムに追加します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    strCommand = CMD_ADD & GetPath(WB.FullName) & OPT_END0
    Run strCommand
    
    MsgBox "追加しました。", vbOKOnly + vbExclamation, C_TITLE
    
    Exit Sub
e:
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Brouser()

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
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strCommand = CMD_CLEANUP & "/PATH:" & rlxGetFullpathFromPathName(strBook) & " " & OPT_END0
    WB.Close False
    
    Run strCommand

    Workbooks.Open strBook
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Commit()

    Dim strCommand As String
    Dim WB As Workbook
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_COMMIT & GetPath(WB.FullName) & OPT_END0
    Run strCommand

    Exit Sub
e:
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
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
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
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
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Locked()

End Sub

Private Sub IVersion_Pull()

End Sub

Private Sub IVersion_Push()

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_PUSH & "/PATH:" & rlxGetFullpathFromPathName(strBook) & " " & OPT_END0
    Run strCommand
    
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
    
End Sub

Private Sub IVersion_RevisionGraph()

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_REVISIONGRAPH & "/PATH:" & rlxGetFullpathFromPathName(strBook) & " " & OPT_END0
    Run strCommand
    
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
    
End Sub

Private Sub IVersion_Tag()

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_TAG & "/PATH:" & rlxGetFullpathFromPathName(strBook) & " " & OPT_END0
    Run strCommand
    
    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE

End Sub

Private Sub IVersion_Unlocked()

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
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub IVersion_Revert()
    
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
    
    strCommand = CMD_REVERT & GetPath(WB.FullName) & OPT_END0
    Run strCommand

    WB.ChangeFileAccess Mode:=xlReadWrite
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Exit Sub
e:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub


Private Sub IVersion_Update()
    
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
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub Run(ByVal strExe As String)
    On Error Resume Next

    err.Clear
    With CreateObject("WScript.Shell")
        .Run EXE_NAME & " " & OPT_COMMAND & strExe, 1, True
    End With
    If err.Number <> 0 Then
        MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
    End If
    
    Exit Sub
e:
    MsgBox "TortoiseGitの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub
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


