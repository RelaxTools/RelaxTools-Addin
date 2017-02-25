VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinProcess"
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

#If VBA7 And Win64 Then

    Private Type PROCESSENTRY32
        Size As Long
        RefCount As Long
        ProcessID As Long
        HeapID As LongPtr
        ModuleID As Long
        ThreadCount As Long
        ParentProcessID As Long
        BasePriority As Long
        Flags As Long
        FileName As String * 260
    End Type
    
    Private Declare PtrSafe Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal Flags As Long, ByVal ProcessID As Long) As LongPtr
    Private Declare PtrSafe Function Process32First Lib "kernel32" (ByVal lngHandleshot As LongPtr, ByRef ProcessEntry As PROCESSENTRY32) As LongPtr
    Private Declare PtrSafe Function Process32Next Lib "kernel32" (ByVal lngHandle As LongPtr, ByRef ProcessEntry As PROCESSENTRY32) As LongPtr
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Private Const SIZEOF_PROCESSENTRY32 As Long = 304

#Else
    
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
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Const SIZEOF_PROCESSENTRY32 As Long = 296

#End If

Private Type ENTRY
    ProcessID As Long
    ParentProcessID As Long
    FileName As String
End Type
Private mEntry() As ENTRY
Private mEntryCount As Long

'--- Win32 API 定数の宣言 ---
Private Const TH32CS_SNAPPROCESS = 2
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Sub Class_Initialize()
    mEntryCount = 0
End Sub
'指定したPidの子プロセスのPidを取得する。
Function GetChildProcessByPid(ByVal lngPid As Long) As Long
    Dim i As Long
    GetChildProcessByPid = 0
    
    For i = 1 To mEntryCount
        If mEntry(i).ParentProcessID = lngPid Then
            GetChildProcessByPid = mEntry(i).ProcessID
            Exit For
        End If
    Next
End Function
'指定したPidにてファイル名を取得する。
Function GetFilenameByPid(ByVal lngPid As Long) As String
    Dim i As Long
    GetFilenameByPid = ""
    
    For i = 1 To mEntryCount
        If mEntry(i).ParentProcessID = lngPid Then
            GetFilenameByPid = mEntry(i).FileName
            Exit For
        End If
    Next
End Function
'指定した名前のプロセスが存在することを確認
Function IsExistProcess(ByVal strFile As String) As Boolean
    Dim i As Long
    IsExistProcess = False
    
    For i = 1 To mEntryCount
        If InStr(UCase(mEntry(i).FileName), UCase(strFile)) > 0 Then
            IsExistProcess = True
            Exit For
        End If
    Next
End Function
'プロセス情報の取得
Public Sub SnapShot()

#If VBA7 And Win64 Then
    Dim lngResult As LongPtr
    Dim lngHandle As LongPtr
#Else
    Dim lngResult As Long
    Dim lngHandle As Long
#End If

    Dim procEntry As PROCESSENTRY32
    Dim lngRet As Long
    Dim lngPos As Long
    
    procEntry.Size = SIZEOF_PROCESSENTRY32
    
    mEntryCount = 1
    
    lngHandle = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    lngResult = Process32First(lngHandle, procEntry)
    Do While lngResult
    
        ReDim Preserve mEntry(1 To mEntryCount)
    
        mEntry(mEntryCount).ProcessID = procEntry.ProcessID
        mEntry(mEntryCount).ParentProcessID = procEntry.ParentProcessID
        lngPos = InStr(procEntry.FileName, vbNullChar)
        If lngPos > 0 Then
            mEntry(mEntryCount).FileName = Mid$(procEntry.FileName, 1, lngPos - 1)
        Else
             mEntry(mEntryCount).FileName = ""
        End If
        
        mEntryCount = mEntryCount + 1
        lngResult = Process32Next(lngHandle, procEntry)
    Loop
    
    CloseHandle lngHandle

End Sub

