Attribute VB_Name = "basTimeLeap"
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
Option Private Module
Sub execOpen(ByVal blnFlg As Boolean)
    Call frmTimeLeap.execOpen(blnFlg)
End Sub
Sub execCompare()
    Call frmTimeLeap.execCompare
End Sub
Sub execCompareBefore()
    Call frmTimeLeap.execCompareBefore
End Sub
Sub execOverwrite()
    Call frmTimeLeap.execOverwrite
End Sub
Sub execExport()
    Call frmTimeLeap.execExport
End Sub

Sub timeLeapList()
    frmTimeLeap.Show
End Sub

'履歴の作成
Sub CreateHistory(ByVal strBook As String)
    'ファイル存在チェック
    
    With CreateObject("Scripting.FileSystemObject")
    
        If Not .FileExists(strBook) Then
            Exit Sub
        End If
        
        Dim strFolder As String
        strFolder = GetSetting(C_TITLE, "TimeLeap", "Folder", GetTimeLeapFolder())
        If Not .FolderExists(strFolder) Then
            .createFolder strFolder
        End If
        
        Dim lngGen As Long
        lngGen = Val(GetSetting(C_TITLE, "TimeLeap", "Gen", "99"))
        
        Dim i As Long
        
        Dim DateCreated As Date
        Dim DateLastModified As Date

        Dim a As FileTime
        Set a = New FileTime
        
        '99～1までバックアップ
        For i = lngGen To 1 Step -1
        
            Dim strFullName As String
            strFullName = .BuildPath(strFolder, rlxGetFullpathFromFileName(strBook))
        
            Dim strSourceFile As String
            Dim strDestFile As String
            
            strSourceFile = strFullName & "." & Format$(i, "000")
            If .FileExists(strSourceFile) Then
                If i = lngGen Then
                    .Deletefile strSourceFile, True
                Else
                
                    DateCreated = .GetFile(strSourceFile).DateCreated
                    DateLastModified = .GetFile(strSourceFile).DateLastModified
                
                    strDestFile = .getFileName(strFullName) & "." & Format$(i + 1, "000")
                    .GetFile(strSourceFile).Name = strDestFile
                    
                    a.SetCreationTime strSourceFile, DateCreated
                    a.SetLastWriteTime strSourceFile, DateLastModified
                    
                End If
            End If
        
        Next
        
        '001作成
        DateCreated = .GetFile(strBook).DateCreated
        DateLastModified = .GetFile(strBook).DateLastModified
        
        strFullName = .BuildPath(strFolder, rlxGetFullpathFromFileName(strBook)) & ".001"
        .CopyFile strBook, strFullName, True
        
        a.SetCreationTime strFullName, DateCreated
        a.SetLastWriteTime strFullName, DateLastModified
        
    End With

End Sub
'--------------------------------------------------------------
'　アプリケーションフォルダ取得
'--------------------------------------------------------------
Public Function GetTimeLeapFolder() As String

    On Error Resume Next
    
    Dim strFolder As String
    
    GetTimeLeapFolder = ""
    
    With CreateObject("Scripting.FileSystemObject")
    
        strFolder = .BuildPath(CreateObject("Wscript.Shell").SpecialFolders("AppData"), C_TITLE) & "TimeLeap"
        
'        If .FolderExists(strFolder) Then
'        Else
'            .createFolder strFolder
'        End If
        
        GetTimeLeapFolder = .BuildPath(strFolder, "\")
        
    End With
    

End Function

Sub openTimeLeapFolder()

    On Error Resume Next
    
    With CreateObject("Scripting.FileSystemObject")
        Dim strFolder As String
        strFolder = GetSetting(C_TITLE, "TimeLeap", "Folder", GetTimeLeapFolder())
        If Not .FolderExists(strFolder) Then
            .createFolder strFolder
        End If
    End With
    
    With CreateObject("WScript.Shell")
        .Run (strFolder)
    End With

End Sub
