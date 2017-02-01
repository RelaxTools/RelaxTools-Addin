VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TortoiseSVN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVersion

Private Const EXE_NAME As String = "TortoiseProc.exe"
Private Const OPT_COMMAND As String = "/command:"
Private Const OPT_PATH As String = "/path:"
Private Const OPT_END0 As String = "/closeonend:0"
Private Const OPT_END1 As String = "/closeonend:1"

Private Const CMD_ADD As String = "add"
Private Const CMD_COMMIT As String = "commit"
Private Const CMD_DIFF As String = "diff"
Private Const CMD_LOCK As String = "lock"
Private Const CMD_UNLOCK As String = "unlock"
Private Const CMD_LOG As String = "log"
Private Const CMD_UPDATE As String = "update"
Private Const CMD_REVERT As String = "revert"
Private Const CMD_BROUSER As String = "repobrowser"
Private Const CMD_CLEANUP As String = "cleanup"
Private Const CMD_ABOUT As String = "about"
Private Const CMD_HELP As String = "help"

Private Sub IVersion_Add()

    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    Set WB = ActiveWorkbook
    
    If WB Is Nothing Then
        Exit Sub
    End If
    
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "ブックが保存されていません。" & vbCrLf & "ブックを保存してから再実行してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 1, 1
        If .IsSvnItem Then
            MsgBox "すでにバージョン管理に登録されているファイルです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
        
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。保存しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            MsgBox "処理を中断しました。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        Else
            WB.Save
            WB.Saved = True
        End If
    End If
    
    strCommand = CMD_ADD & " " & GetPath(WB) & OPT_END0
    Run strCommand, WB
    
    MsgBox "追加しました。", vbOKOnly + vbExclamation, C_TITLE

End Sub

Private Sub IVersion_Brouser()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    WB.ChangeFileAccess Mode:=xlReadOnly
    
    strCommand = CMD_BROUSER & " " & GetPath(WB) & OPT_END0
    Run strCommand, WB

    WB.ChangeFileAccess Mode:=xlReadWrite
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub IVersion_Cleanup()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    Set WB = ActiveWorkbook
    strCommand = CMD_CLEANUP & " /PATH:" & rlxGetFullpathFromPathName(WB.FullName) & OPT_END0
    
    Run strCommand, WB
    
End Sub

Private Sub IVersion_Commit()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_COMMIT & " " & GetPath(WB) & OPT_END0
    Run strCommand, WB

End Sub

Private Sub IVersion_Diff()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    strCommand = CMD_DIFF & " " & GetPath(WB) & OPT_END0
    Run strCommand, WB
End Sub

Private Sub IVersion_Help()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    strCommand = CMD_HELP
    Run strCommand, WB
    
End Sub

Private Sub IVersion_Locked()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo WB.FullName, 1, 1
        If .IsLocked Then
            MsgBox "ブックは既にロックされています。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
    
    strCommand = CMD_LOCK & " " & GetPath(WB) & OPT_END0
    Run strCommand, WB
End Sub
Private Sub IVersion_Unlocked()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo WB.FullName, 1, 1
        If .IsLocked Then
        Else
            MsgBox "ブックはロックされていません。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
    
    strCommand = CMD_UNLOCK & " " & GetPath(WB) & OPT_END0
    Run strCommand, WB
End Sub
Private Sub IVersion_Log()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    If IsError(WB) Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    WB.ChangeFileAccess Mode:=xlReadOnly
    
    strCommand = CMD_LOG & " " & GetPath(WB) & OPT_END0
    Run strCommand, WB

    WB.ChangeFileAccess Mode:=xlReadWrite
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Private Sub IVersion_Revert()
    
    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    Set WB = ActiveWorkbook
    
    If WB Is Nothing Then
        Exit Sub
    End If
    
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "ブックが保存されていません。" & vbCrLf & "ブックを保存してから再実行してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo WB.FullName, 1, 1
        If .IsSvnItem Then
        Else
            MsgBox "バージョン管理外のファイルです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
    
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。破棄しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            MsgBox "処理を中断しました。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End If
    
    On Error Resume Next
        
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strCommand = CMD_REVERT & " " & GetPath(WB) & OPT_END0
    WB.Close False
    
    Run strCommand, WB

    Workbooks.Open strBook
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


Private Sub IVersion_Update()
    
    Dim strCommand As String
    Dim WB As Workbook
    Dim strBook As String
    
    Set WB = ActiveWorkbook
    
    If WB Is Nothing Then
        Exit Sub
    End If
    
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "ブックが保存されていません。" & vbCrLf & "ブックを保存してから再実行してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 1, 1
        If .HasModifications Then
            MsgBox "作業フォルダのブックは変更されているため、更新することができません。" & vbCrLf & "変更を元に戻す場合は「変更の取り消し」を行ってください。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
        If .IsSvnItem Then
        Else
            MsgBox "バージョン管理外のファイルです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End With
    
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。破棄しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            MsgBox "処理を中断しました。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End If
    
    If MsgBox("HEADに更新します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    On Error Resume Next
        
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strCommand = CMD_UPDATE & " " & GetPath(WB) & OPT_END0
    WB.Close False
    
    Run strCommand, WB

    Workbooks.Open strBook
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Private Sub Run(ByVal strExe As String, WB As Workbook)
    On Error Resume Next
    

    err.Clear
    With CreateObject("WScript.Shell")
        .Run EXE_NAME & " " & OPT_COMMAND & strExe, 1, True
    End With
    If err.Number <> 0 Then
        MsgBox "TortoiseSVNの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
    End If
    
    
End Sub
Private Function GetPath(WB As Workbook) As String
    GetPath = OPT_PATH & """" & WB.FullName & """ "
End Function
Private Function IsError(WB As Workbook) As Boolean

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
    
    With CreateObject("SubWCRev.object")
        .GetWCInfo strBook, 1, 1
        If .IsSvnItem Then
        Else
            MsgBox "バージョン管理外のファイルです。", vbOKOnly + vbExclamation, C_TITLE
            Exit Function
        End If
    End With
        
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。保存しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            MsgBox "処理を中断しました。", vbOKOnly + vbExclamation, C_TITLE
            Exit Function
        Else
            WB.Save
            WB.Saved = True
        End If
    End If
    
    IsError = False
    
End Function

Private Sub IVersion_Ver()

    Dim strCommand As String
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    strCommand = CMD_ABOUT
    Run strCommand, WB
    
End Sub
