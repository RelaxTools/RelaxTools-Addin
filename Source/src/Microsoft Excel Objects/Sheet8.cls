VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
Private Const C_File As String = "RELAXTOOLS.XLAM"
Private Const C_TITLE As String = "RelaxTools-Addin"

Sub reboot()

    Dim a As AddIn

    Err.Clear
    On Error Resume Next
    
    If MsgBox("RelaxToolsを再起動しますがよろしいですか？", vbYesNo + vbQuestion, C_TITLE) <> vbYes Then
        Exit Sub
    End If

    For Each a In Application.AddIns

        If UCase(a.Name) = C_File Then
            a.Installed = False
            a.Installed = True
            Exit For
        End If

    Next

    If Err.Number <> 0 Then
        MsgBox "再起動中にエラーがありました。Excel自体の再起動が必要な場合があります。", vbExclamation + vbOKOnly, C_TITLE
    Else
        MsgBox "RelaxToolsの再起動が完了しました。", vbInformation + vbOKOnly, C_TITLE
    End If

End Sub
Sub RelaxInit()

    Dim a As AddIn

    Err.Clear
    On Error Resume Next

    If MsgBox("RelaxToolsの設定を初期化しますがよろしいですか？", vbYesNo + vbQuestion, C_TITLE) <> vbYes Then
        Exit Sub
    End If
    
    For Each a In Application.AddIns

        If UCase(a.Name) = C_File Then
        
            a.Installed = False
            
            DeleteSetting C_TITLE
            
            a.Installed = True
            Exit For
        End If

    Next

    If Err.Number <> 0 Then
        MsgBox "初期化中にエラーがありました。Excel自体の再起動が必要な場合があります。", vbExclamation + vbOKOnly, C_TITLE
    Else
        MsgBox "RelaxToolsの初期化が完了しました。", vbInformation + vbOKOnly, C_TITLE
    End If

End Sub
Sub RelaxUninstall()

    Dim a As AddIn
    Dim strFolder As String
    
    On Error GoTo e
    
    If MsgBox("RelaxToolsをアンインストールしますがよろしいですか？", vbYesNo + vbQuestion, C_TITLE) <> vbYes Then
        Exit Sub
    End If

    For Each a In Application.AddIns

        If UCase(a.Name) = C_File Then
            a.Installed = False
            Kill a.FullName
            Exit For
        End If

    Next
    
    With CreateObject("Scripting.FileSystemObject")
    
        strFolder = .BuildPath(CreateObject("Wscript.Shell").SpecialFolders("AppData"), C_TITLE)
        
        If .FolderExists(strFolder) Then
            .DeleteFolder strFolder, True
        End If
        
    End With

    DeleteSetting C_TITLE
    MsgBox "正常にアンインストールされました。", vbInformation + vbOKOnly, C_TITLE
    Exit Sub
e:
    MsgBox "アンインストールが正常に行われませんでした。", vbCritical + vbOKOnly, C_TITLE
    
End Sub
