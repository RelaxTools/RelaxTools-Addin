Attribute VB_Name = "basGit"
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
'Option Explicit
'Option Private Module
'
'#If VBA7 And Win64 Then
'    Private Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'#Else
'    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'#End If
'
'Sub GitLog()
'
'    Dim strCommand As String
'    Dim cmd As CommandLine
'    Dim strSysout As String
'    Dim strFile As String
'    Dim strPath As String
'    Dim exitcode As Long
'
'    If ActiveWorkbook Is Nothing Then
'        Exit Sub
'    End If
'
'    strPath = rlxGetFullpathFromPathName(ActiveWorkbook.FullName)
'    strFile = rlxGetFullpathFromFileName(ActiveWorkbook.FullName)
'
'    strSysout = ""
'
'    strCommand = "git log --date=iso --pretty=format:""[%ad] %an : %s"" " & strFile
'
'    Set cmd = New CommandLine
'
'    exitcode = cmd.Run(strPath, strCommand, GetEnv, strSysout, True)
'
'    If exitcode <> 0 Then
'        strSysout = "処理中にエラーが発生しました。ExitCode : " & exitcode & vbLf & vbLf & strSysout
'    End If
'
'    frmGitResult.Start strSysout
'
'    Set cmd = Nothing
'
'End Sub
'Sub GitReset()
'
'    Dim strCommand As String
'    Dim cmd As CommandLine
'    Dim strSysout As String
'
'    Dim strPath As String
'    Dim strFile As String
'    Dim exitcode As Long
'    Dim strBook As String
'    Dim blnReadOnly As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        Exit Sub
'    End If
'
'
'    strBook = ActiveWorkbook.FullName
'
'    If rlxIsFileExists(strBook) Then
'
'        If MsgBox("変更を取り消します。" & vbCrLf & "よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
'            Exit Sub
'        End If
'
'        Application.ScreenUpdating = False
'
'        blnReadOnly = ActiveWorkbook.ReadOnly
'
'        ActiveWorkbook.Close
'
'
'        strPath = rlxGetFullpathFromPathName(strBook)
'        strFile = rlxGetFullpathFromFileName(strBook)
'
'        strCommand = "git checkout """ & strFile & """"
'
'        Set cmd = New CommandLine
'
'        exitcode = cmd.Run(strPath, strCommand, GetEnv, strSysout, True)
'
'        Workbooks.Open strBook, ReadOnly:=blnReadOnly
'        Application.ScreenUpdating = True
'        If exitcode <> 0 Then
'            strSysout = "処理中にエラーが発生しました。ExitCode : " & exitcode & vbLf & vbLf & strSysout
'            frmGitResult.Start strSysout
'        Else
'            MsgBox "変更を取り消しました。", vbOKOnly + vbInformation, C_TITLE
'        End If
'
'        Set cmd = Nothing
'
'    End If
'
'End Sub
'Sub GitCommit()
'
'    Dim strCommand As String
'    Dim cmd As CommandLine
'    Dim strSysout As String
'
'    Dim strPath As String
'    Dim strFile As String
'    Dim exitcode As Long
'    Dim strBook As String
'    Dim blnReadOnly As Boolean
'    Dim strMessage As String
'
'    If ActiveWorkbook Is Nothing Then
'        Exit Sub
'    End If
'
'
'    strBook = ActiveWorkbook.FullName
'
'    If rlxIsFileExists(strBook) Then
'
'        If MsgBox("コミットします。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
'            Exit Sub
'        End If
'
'        Application.ScreenUpdating = False
'
'        strMessage = "commit"
'
'        strPath = rlxGetFullpathFromPathName(strBook)
'        strFile = rlxGetFullpathFromFileName(strBook)
'
'        strCommand = "git commit -m """ & strMessage & """ """ & strFile & """"
'
'        Set cmd = New CommandLine
'
'        exitcode = cmd.Run(strPath, strCommand, GetEnv, strSysout, True)
'
'        Application.ScreenUpdating = True
'        If exitcode <> 0 Then
'            strSysout = "処理中にエラーが発生しました。ExitCode : " & exitcode & vbLf & vbLf & strSysout
'            frmGitResult.Start strSysout
'        Else
'            strSysout = "処理を実行しました。ExitCode : " & exitcode & vbLf & vbLf & strSysout
'            frmGitResult.Start strSysout
'        End If
'
'        Set cmd = Nothing
'
'    End If
'
'End Sub
'Private Function GetEnv() As String
'
'    Dim strKey As String
'    Dim strBuffer As String
'    Dim strHome As String
'    Dim lngPos As Long
'
'    ' 環境変数HOME を取得、無い場合には USERPROFILE をセット
'    strKey = "HOME"
'    strBuffer = String(1024, vbNullChar)
'
'    GetEnvironmentVariable strKey, strBuffer, 1024
'
'    lngPos = InStr(strBuffer, vbNullChar)
'    strHome = Left(strBuffer, lngPos - 1)
'
'    If strHome = "" Then
'        strKey = "USERPROFILE"
'        strBuffer = String(1024, vbNullChar)
'
'        GetEnvironmentVariable strKey, strBuffer, 1024
'        lngPos = InStr(strBuffer, vbNullChar)
'        strHome = Left(strBuffer, lngPos - 1)
'    End If
'
'    GetEnv = "HOME=" & strHome & vbNullChar & vbNullChar
'
'End Function
