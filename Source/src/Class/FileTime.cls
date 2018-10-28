VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileTime"
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

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
    Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
    Private Declare PtrSafe Function SetFileTime Lib "kernel32" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
#Else
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (ByRef lpLocalFileTime As FILETIME, ByRef lpFileTime As FILETIME) As Long
    Private Declare Function SystemTimeToFileTime Lib "kernel32" (ByRef lpSystemTime As SystemTime, ByRef lpFileTime As FILETIME) As Long
    Private Declare Function SetFileTime Lib "kernel32" (ByVal cFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
#End If

' SystemTime 構造体
Private Type SystemTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type

' FileTime 構造体
Private Type FILETIME
    LowDateTime As Long
    HighDateTime As Long
End Type

' 定数の定義
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const OPEN_EXISTING As Long = 3
Public Sub SetCreationTime(ByVal stFilePath As String, ByVal dtCreateTime As Date)
    
#If VBA7 And Win64 Then
    Dim cFileHandle As LongPtr
#Else
    Dim cFileHandle As Long
#End If
    Dim tFileTime As FILETIME
    Dim tNullable As FILETIME
    
    cFileHandle = GetFileHandle(stFilePath)
    If cFileHandle <> 0 Then
        tFileTime = GetFileTime(dtCreateTime)
        Call SetFileTime(cFileHandle, tFileTime, tNullable, tNullable)
        Call CloseHandle(cFileHandle)
    End If
    
End Sub
Public Sub SetLastWriteTime(ByVal stFilePath As String, ByVal dtUpdateTime As Date)
    
#If VBA7 And Win64 Then
    Dim cFileHandle As LongPtr
#Else
    Dim cFileHandle As Long
#End If
    Dim tFileTime As FILETIME
    Dim tNullable As FILETIME
    
    cFileHandle = GetFileHandle(stFilePath)
    If cFileHandle <> 0 Then
        tFileTime = GetFileTime(dtUpdateTime)
        Call SetFileTime(cFileHandle, tNullable, tNullable, tFileTime)
        Call CloseHandle(cFileHandle)
    End If

End Sub

Public Sub SetLastAccessTime(ByVal stFilePath As String, ByVal dtAccessTime As Date)

#If VBA7 And Win64 Then
    Dim cFileHandle As LongPtr
#Else
    Dim cFileHandle As Long
#End If
    Dim tNullable As FILETIME
    Dim tFileTime As FILETIME
    
    cFileHandle = GetFileHandle(stFilePath)
    If cFileHandle <> 0 Then
        tFileTime = GetFileTime(dtAccessTime)
        Call SetFileTime(cFileHandle, tNullable, tFileTime, tNullable)
        Call CloseHandle(cFileHandle)
    End If
    
End Sub


' FileTime を取得する
Private Function GetFileTime(ByVal dtSetting As Date) As FILETIME

    Dim tSystemTime As SystemTime
    
    With tSystemTime
        .Year = Year(dtSetting)
        .Month = Month(dtSetting)
        .DayOfWeek = Weekday(dtSetting)
        .Day = Day(dtSetting)
        .Hour = Hour(dtSetting)
        .Minute = Minute(dtSetting)
        .Second = Second(dtSetting)
    End With
    
    Dim tLocalTime As FILETIME
    Call SystemTimeToFileTime(tSystemTime, tLocalTime)
    
    Dim tFileTime As FILETIME
    Call LocalFileTimeToFileTime(tLocalTime, tFileTime)
    
    GetFileTime = tFileTime
    
End Function

' ファイルのハンドルを取得する
#If VBA7 And Win64 Then
    Private Function GetFileHandle(ByVal stFilePath As String) As LongPtr
    
        GetFileHandle = CreateFile(stFilePath, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    End Function
#Else
    Private Function GetFileHandle(ByVal stFilePath As String) As Long
    
        GetFileHandle = CreateFile(stFilePath, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    End Function
#End If

