VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitLib"
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
