VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Function Commit(ByVal strBook As String, ByVal strMessage As String) As String
'
'    Dim strCommand As String
'    Dim cmd As CommandLine
'    Dim strSysout As String
'
'    Dim strPath As String
'    Dim strFile As String
'
'    Commit = ""
'
'    strPath = rlxGetFullpathFromPathName(strBook)
'    strFile = rlxGetFullpathFromFileName(strBook)
'
'    strCommand = "git commit -m """ & strMessage & """ """ & strFile & """"
'
'    ChDir strPath
'
'    Set cmd = New CommandLine
'
'    If cmd.Run(strPath, strCommand, strSysout) = 0 Then
'        Commit = strSysout
'    End If
'
'    Set cmd = Nothing
'
'End Function
'Function Reset(ByVal strBook As String) As String
'
'    Dim strCommand As String
'    Dim cmd As CommandLine
'    Dim strSysout As String
'
'    Dim strPath As String
'    Dim strFile As String
'
'    Reset = ""
'
'    strPath = rlxGetFullpathFromPathName(strBook)
'    strFile = rlxGetFullpathFromFileName(strBook)
'
'    strCommand = "git reset """ & strFile & """"
'
'    ChDir strPath
'
'    Set cmd = New CommandLine
'
'    If cmd.Run(strPath, strCommand, strSysout) = 0 Then
'        Reset = strSysout
'    End If
'
'    Set cmd = Nothing
'
'End Function
'
'Function Log(ByVal strPath As String)
'
'    Dim strCommand As String
'    Dim cmd As CommandLine
'    Dim strSysout As String
'
'    Log = ""
'
'    strCommand = "git log --pretty=format:""[%ad] %an : %s"""
'
'    ChDir strPath
'
'    Set cmd = New CommandLine
'
'    If cmd.Run(strPath, strCommand, strSysout) = 0 Then
'        Log = strSysout
'    End If
'
'    Set cmd = Nothing
'
'End Function
