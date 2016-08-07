VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 And Win64 Then

    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numBytes As LongPtr)
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByRef lpExitCode As Long) As Long
    
    Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpCommandLine As String, _
        ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
        ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal bInheritHandles As Long, _
        ByVal dwCreationFlags As Long, _
        ByRef lpEnvironment As Any, _
        ByVal lpCurrentDriectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    
    Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
        ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    
    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As LongPtr
        bInheritHandle As Long
    End Type
    
    Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As LongPtr
        hStdInput As LongPtr
        hStdOutput As LongPtr
        hStdError As LongPtr
    End Type
    
    Private Type PROCESS_INFORMATION
        hProcess As LongPtr
        hThread As LongPtr
        dwProcessId As Long
        dwThreadId As Long
    End Type
    
    Private Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

#Else

    Private Declare Function WaitForSingleObject Lib "KERNEL32.DLL" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numBytes As Long)
    Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
    
    Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpCommandLine As String, _
        ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
        ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal bInheritHandles As Long, _
        ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As String, _
        ByVal lpCurrentDriectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
    
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
        ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    
    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
    End Type
    
    Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
    End Type
    
    Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
    End Type

    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

#End If

Private Const HANDLE_FLAG_INHERIT = &H1
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Const SW_HIDE = 0
Private Const STILL_ACTIVE = &H103
Private Const CREATE_NO_WINDOW = &H8000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const INVALID_HANDLE_VALUE = -1
Private Const INFINITE = &HFFFFFFFF

'Public Enum RunExitCode
'    Success = 0
'    SysoutOpenError = 1
'    CreateProcessError = 2
'    NoSysout = 3
'End Enum

Public Function Run(ByVal strPath As String, ByVal strCommand As String, ByVal strEnv As String, ByRef strSysout As String, Optional ByVal IsUTF8 As Variant = False) As Long

    Dim sa As SECURITY_ATTRIBUTES
    Dim pi As PROCESS_INFORMATION
    Dim si As STARTUPINFO
    
#If VBA7 And Win64 Then
    Dim hOutWrite As LongPtr
#Else
    Dim hOutWrite As Long
#End If
    
    Dim strFileName As String
    Dim exitcode As Long
    
    Run = 0
    strSysout = ""
    
    Call ZeroMemory(sa, Len(sa))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1& 'TRUE
    sa.lpSecurityDescriptor = 0& ' NULL
    
    strFileName = rlxGetTempFolder & "sysout.txt"
    Dim fp As Integer
    fp = FreeFile()
    Open strFileName For Output As fp
    Close fp
    
    hOutWrite = CreateFile(strFileName, GENERIC_WRITE Or GENERIC_READ, 0&, sa, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hOutWrite = 0 Then
        Run = -1
        GoTo Error
    End If

    Call ZeroMemory(si, Len(si))
    si.cb = Len(si)
    si.hStdInput = 0 'GetStdHandle(STD_INPUT_HANDLE)
    si.hStdOutput = hOutWrite
    si.hStdError = hOutWrite
    si.dwFlags = STARTF_USESTDHANDLES
    si.wShowWindow = SW_HIDE

    ' 環境変数をセット
    strEnv = strEnv & vbNullChar & vbNullChar

    If CreateProcess(vbNullString, strCommand, sa, sa, 1&, CREATE_NO_WINDOW, strEnv, strPath, si, pi) = 0 Then
        Call CloseHandle(hOutWrite)
        Run = -2
        GoTo Error
    End If

    '終了するまで待機
    WaitForSingleObject pi.hProcess, INFINITE
    
    '終了ステータスコードの取得
    If GetExitCodeProcess(pi.hProcess, exitcode) = 0 Then
        Run = -3
        GoTo Error
    End If
    
    Run = exitcode
    
    Call CloseHandle(pi.hThread)
    Call CloseHandle(pi.hProcess)
    Call CloseHandle(hOutWrite)
    
    Dim lngsize As Long
    Dim bytBuf() As Byte
    
    fp = FreeFile()
    Open strFileName For Binary As fp

    lngsize = LOF(fp)
    If lngsize = 0 Then
        Close fp
        GoTo Error
    End If
    
    ReDim bytBuf(0 To lngsize - 1)
    
    Get #fp, , bytBuf

    Close fp
    
    If IsUTF8 Then
        ' UTF8に変換
        Dim utf8 As UTF8Encoding
        
        Set utf8 = New UTF8Encoding
        
        strSysout = utf8.GetString(bytBuf)
        
        Set utf8 = Nothing
    Else
        'SJISに変換
        strSysout = StrConv(bytBuf, vbUnicode)
    End If
    
    On Error Resume Next
    Kill strFileName
    
    Exit Function
Error:

End Function
