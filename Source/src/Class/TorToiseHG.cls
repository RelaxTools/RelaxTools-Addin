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

#If VBA7 And Win64 Then

    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
    
    Private Type PROCESSENTRY32
        Size As Long
        RefCount As Long
        ProcessID As Long
        HeapID As LongPtr
        ModuleID As Long
        ThreadCount As Long
        ParentProcessID As Long
        BasePriority As LongPtr
        Flags As Long
        FileName As String * 260
    End Type
    
    Private Declare PtrSafe Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal Flags As Long, ByVal ProcessID As Long) As LongPtr
    Private Declare PtrSafe Function Process32First Lib "kernel32" (ByVal lngHandleshot As LongPtr, ByRef ProcessEntry As PROCESSENTRY32) As LongPtr
    Private Declare PtrSafe Function Process32Next Lib "kernel32" (ByVal lngHandle As LongPtr, ByRef ProcessEntry As PROCESSENTRY32) As LongPtr
    Private Const SIZEOF_PROCESSENTRY32 As Long = 304

#Else
    
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

    Private Type PROCESSENTRY32
        Size As Long
        RefCount As Long
        ProcessID As Long
        HeapID As Long
        ModuleID As Long
        ThreadCount As Long
        ParentProcessID As Long
        BasePriority As Long
        Flags As Long
        FileName As String * 260
    End Type
    
    Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal Flags As Long, ByVal ProcessID As Long) As Long
    Private Declare Function Process32First Lib "kernel32" (ByVal lngHandleshot As Long, ByRef ProcessEntry As PROCESSENTRY32) As Long
    Private Declare Function Process32Next Lib "kernel32" (ByVal lngHandle As Long, ByRef ProcessEntry As PROCESSENTRY32) As Long
    Private Const SIZEOF_PROCESSENTRY32 As Long = 296

#End If
Private Const TH32CS_SNAPPROCESS = 2
'--- Win32 API 定数の宣言 ---
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const INFINITE As Long = &HFFFF
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
    
    On Error GoTo e
    
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
    
    On Error GoTo e
    
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
    
    On Error GoTo e
    
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
    
    On Error GoTo e
    
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
    
    On Error GoTo e
    
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
    
    On Error GoTo e
    
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
    
    On Error GoTo e
    
    Set WB = ActiveWorkbook
    
    strPath = rlxGetFullpathFromPathName(WB.FullName)
    strCommand = CMD_ABOUT
    Run strPath, strCommand

    
    Exit Sub
e:
    MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
End Sub

Private Sub Run(ByVal strPath As String, ByVal strExe As String)

    Dim TaskId As Long          'タスクID
    
#If VBA7 And Win64 Then
    Dim hProc  As LongPtr          'プロセスハンドル
#Else
    Dim hProc  As Long          'プロセスハンドル
#End If
    Dim hWnd As Long
    Dim lngCount As Long

    ' 外部プログラムの実行
    On Error GoTo e
    
    ChDir strPath
    err.Clear
    TaskId = GetChildPid(Shell(EXE_NAME & " " & strExe))
    If err.Number <> 0 Then
        MsgBox "TortoiseHGの起動に失敗しました。インストールされていないか、PATHの設定を確認してください。", vbOKOnly + vbCritical, C_TITLE
        Exit Sub
    End If
   
    ' プロセスハンドルの取得
    hProc = OpenProcess(PROCESS_ALL_ACCESS, 0, TaskId)
    ' プロセスハンドルが返されたかを判定
    If hProc <> 0 Then
        ' プロセスのシグナル待ち
        Do Until WaitForSingleObject(hProc, 100) = 0
            DoEvents
        Loop
        ' プロセスクローズ
        CloseHandle hProc
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



Function GetChildPid(ByVal lngPid As Long) As Long

#If VBA7 And Win64 Then
    Dim lngResult As LongPtr
    Dim lngHandle As LongPtr
#Else
    Dim lngResult As Long
    Dim lngHandle As Long
#End If

    Dim procEntry As PROCESSENTRY32
    Dim lngRet As Long
    
    lngHandle = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

    procEntry.Size = SIZEOF_PROCESSENTRY32
    
    lngResult = Process32First(lngHandle, procEntry)

    Do While lngResult

        If procEntry.ParentProcessID = lngPid Then
            lngRet = procEntry.ProcessID
            Exit Do
        End If

        lngResult = Process32Next(lngHandle, procEntry)
    Loop

    CloseHandle lngHandle
    
    GetChildPid = lngRet

End Function






