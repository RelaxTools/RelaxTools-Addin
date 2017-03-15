VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogManager"
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

Private fp As Integer

Public Enum LogLevel
    Trace = 0
    Info = 1
    Warn = 2
    Fatal = 3
    None = 4
End Enum

Private mLevel As LogLevel
Private mLogType As Long
Private mOpen As Boolean

Private mLevelStr As Variant


Public Property Let Level(ByVal l As LogLevel)
    mLevel = l
End Property

Private Sub Class_Initialize()
    
    mLevel = GetSetting(C_TITLE, "Log", "Level", LogLevel.Info)
    mLevelStr = Array("Trace", "Info", "Warn", "Fatal", "None")
    mLogType = CLng(GetSetting(C_TITLE, "Log", "LogType", C_LOG_LOGFILE))
    
    If mLogType And C_LOG_LOGFILE Then
        Call OpenLog
    End If

End Sub

Private Sub Class_Terminate()
    
    If mOpen Then
        On Error Resume Next
        Close fp
    End If
    
End Sub

Public Sub Log(ByVal l As LogLevel, ByVal strMsg As String)

    Dim strLog As String
    
    If l >= mLevel Then
        strLog = C_TITLE & "," & Format$(Now, "yyyy-mm-dd,hh:nn:ss") & "." & getMSec & "," & mLevelStr(l) & "," & strMsg
        If mLogType And C_LOG_DEBUGWINDOW Then
            Debug.Print strLog
        End If
        If mLogType And C_LOG_LOGFILE Then
            If mOpen Then
                On Error Resume Next
                Print #fp, strLog
            Else
                Call OpenLog
                On Error Resume Next
                Print #fp, strLog
            End If
        End If
    End If
    
End Sub
Private Function getMSec() As String

    Dim dblTimer As Double
    
    dblTimer = CDbl(Timer)
    getMSec = Format$((dblTimer - Fix(dblTimer)) * 1000, "000")
    
End Function
Public Sub Flash()
    If mOpen Then
        Close fp
        Call OpenLog
    End If
End Sub
Private Sub OpenLog()

    Dim strFile As String
    Dim strLogFolder As String
    Dim lngHistory As Long
    
    On Error GoTo e
    
    'ログフォルダ作成
    strLogFolder = rlxGetAppDataFolder & "Log"
    rlxCreateFolder strLogFolder
    
    '保存件数を取得(最新５ファイルを残す)
    lngHistory = CLng(GetSetting(C_TITLE, "Log", "History", 5))
    
    Dim objFs As Object
    Dim objDic As Object
    Dim objFolder As Object
    Dim objFile As Object
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objDic = CreateObject("Scripting.Dictionary")
    
    'ログフォルダを Dictionaryに追加
    Set objFolder = objFs.GetFolder(strLogFolder)
    For Each objFile In objFolder.files
        objDic.Add objFile.Name, objFile.Name
    Next
    
    Dim i As Long
    Dim v As Variant
    
    'フォルダ名の昇順に処理する
    i = objDic.count
    For Each v In objDic
    
        If i > lngHistory Then
            Kill strLogFolder & "\" & v
        End If
        i = i - 1
    Next
   
    Set objDic = Nothing
    Set objFolder = Nothing
    Set objFs = Nothing
    
    
    fp = FreeFile()
    strFile = GetSetting(C_TITLE, "Log", "Folder", strLogFolder & "\" & Format(Now, "yyyymmdd") & ".log")
    Open strFile For Append As fp
    mOpen = True
    Exit Sub
e:
    mOpen = False
    
End Sub
