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
    Warn = 1
    Info = 2
    Fatal = 3
    None = 4
End Enum

Private mLevel As LogLevel

Private mLevelStr As Variant


Public Property Let Level(ByVal l As LogLevel)
    mLevel = l
End Property

Private Sub Class_Initialize()
    
    mLevel = GetSetting(C_TITLE, "Logger", "LogLevel", LogLevel.Info)
    mLevelStr = Array("Trace", "Warn", "Info", "Fatal", "None")

'    fp = FreeFile()
'    Dim strFile As String
'
'    strFile = GetSetting(C_TITLE, "Logger", "LogFile", rlxGetAppDataFolder & C_TITLE & ".log")
'    Open strFile For Append As fp

End Sub

Private Sub Class_Terminate()
    
'    Close fp
    
End Sub

Public Sub Log(ByVal l As LogLevel, ByVal strMsg As String)

    Dim strLog As String
    
    If l >= mLevel Then
        strLog = C_TITLE & "," & Format$(Now, "yyyy-mm-dd,hh:nn:ss") & "." & getMSec & "," & mLevelStr(l) & "," & strMsg
        Debug.Print strLog
'        Print #fp, strLog
    End If
    
End Sub
Private Function getMSec() As String

    Dim dblTimer As Double
    
    dblTimer = CDbl(Timer)
    getMSec = Format$((dblTimer - Fix(dblTimer)) * 1000, "000")
    
End Function


