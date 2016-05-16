Attribute VB_Name = "basGit"
Option Explicit

Sub GitLog()

    Dim strCommand As String
    Dim cmd As CommandLine
    Dim strSysout As String
    Dim strFile As String
    Dim strPath As String
    Dim exitcode As Long
    
    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    strPath = rlxGetFullpathFromPathName(ActiveWorkbook.FullName)
    strFile = rlxGetFullpathFromFileName(ActiveWorkbook.FullName)
    
    strSysout = ""
    
    strCommand = "git log --date=iso --pretty=format:""[%ad] %an : %s"" " & strFile
    
    Set cmd = New CommandLine
    
    exitcode = cmd.Run(strPath, strCommand, strSysout)
    
    If exitcode <> 0 Then
        strSysout = "処理中にエラーが発生しました。ExitCode : " & exitcode & vbLf & vbLf & strSysout
    End If

    frmGitResult.Start strSysout

    Set cmd = Nothing

End Sub
Sub GitReset()

    Dim strCommand As String
    Dim cmd As CommandLine
    Dim strSysout As String
    
    Dim strPath As String
    Dim strFile As String
    Dim exitcode As Long
    Dim strBook As String
    Dim blnReadOnly As Boolean
    
    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    
    strBook = ActiveWorkbook.FullName
    
    If rlxIsFileExists(strBook) Then
    
        If MsgBox("変更を取り消して前回コミットの状態に戻します。" & vbCrLf & "よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    
        Application.ScreenUpdating = False
        
        blnReadOnly = ActiveWorkbook.ReadOnly
        
        ActiveWorkbook.Close
        
    
        strPath = rlxGetFullpathFromPathName(strBook)
        strFile = rlxGetFullpathFromFileName(strBook)
        
        strCommand = "git checkout """ & strFile & """"
        
        Set cmd = New CommandLine
        
        exitcode = cmd.Run(strPath, strCommand, strSysout)
    
        Workbooks.Open strBook, ReadOnly:=blnReadOnly
        Application.ScreenUpdating = True
        If exitcode <> 0 Then
            strSysout = "処理中にエラーが発生しました。ExitCode : " & exitcode & vbLf & vbLf & strSysout
            frmGitResult.Start strSysout
        Else
            MsgBox "前回のコミットの状態に戻しました。", vbOKOnly + vbInformation, C_TITLE
        End If
        
        Set cmd = Nothing
        
    End If

End Sub
Sub GitCommit()

    Dim strCommand As String
    Dim cmd As CommandLine
    Dim strSysout As String
    
    Dim strPath As String
    Dim strFile As String
    Dim exitcode As Long
    Dim strBook As String
    Dim blnReadOnly As Boolean
    Dim strMessage As String
    
    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    
    strBook = ActiveWorkbook.FullName
    
    If rlxIsFileExists(strBook) Then
    
        If MsgBox("コミットします。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    
        Application.ScreenUpdating = False
        
        strMessage = "commit"
    
        strPath = rlxGetFullpathFromPathName(strBook)
        strFile = rlxGetFullpathFromFileName(strBook)
        
        strCommand = "git commit -m """ & strMessage & """ """ & strFile & """"
        
        Set cmd = New CommandLine
        
        exitcode = cmd.Run(strPath, strCommand, strSysout)
    
        Application.ScreenUpdating = True
        If exitcode <> 0 Then
            strSysout = "処理中にエラーが発生しました。ExitCode : " & exitcode & vbLf & vbLf & strSysout
            frmGitResult.Start strSysout
        Else
            strSysout = "コミットが正常に行われました。ExitCode : " & exitcode & vbLf & vbLf & strSysout
            frmGitResult.Start strSysout
        End If
        
        Set cmd = Nothing
        
    End If

End Sub
