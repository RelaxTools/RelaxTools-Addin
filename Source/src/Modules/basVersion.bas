Attribute VB_Name = "basVersion"
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
Public Sub TortoiseSVNAdd()

    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Add
    Set Ver = Nothing

End Sub
Public Sub TortoiseSVNCommit()

    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Commit
    Set Ver = Nothing

End Sub

Public Sub TortoiseSVNRevert()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Revert
    
    Set Ver = Nothing

End Sub
Public Sub TortoiseSVNLog()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Log
    Set Ver = Nothing

End Sub

Public Sub TortoiseSVNLock()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Locked
    Set Ver = Nothing

End Sub

Public Sub TortoiseSVNUnlock()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Unlocked
    Set Ver = Nothing

End Sub

Public Sub TortoiseSVNUpdate()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Update
    Set Ver = Nothing

End Sub

Public Sub TortoiseSVNDiff()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Diff
    Set Ver = Nothing

End Sub
Public Sub TortoiseSVNBrouser()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Brouser
    Set Ver = Nothing

End Sub
Public Sub TortoiseSVNCleanup()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Cleanup
    Set Ver = Nothing

End Sub
Public Sub TortoiseSVNVer()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Ver
    Set Ver = Nothing

End Sub
Public Sub TortoiseSVNHelp()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseSVN
    Ver.Help
    Set Ver = Nothing

End Sub


Public Sub TortoiseGitAdd()

    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Add
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitCommit()

    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Commit
    Set Ver = Nothing

End Sub

Public Sub TortoiseGitRevert()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Revert
    
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitLog()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Log
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitCleanup()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Cleanup
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitVer()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Ver
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitHelp()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Help
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitTag()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Tag
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitPush()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Push
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitRevisionGraph()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.RevisionGraph
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitDiff()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Diff
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitBrouser()
    Dim Ver As IVersion
    
    Set Ver = New TortoiseGit
    Ver.Brouser
    Set Ver = Nothing

End Sub
Public Sub TortoiseGitWeb()
    If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    Call GoURL("https://git-scm.com/")
End Sub
Public Sub TortoiseSVNWeb()
    If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    Call GoURL("https://subversion.apache.org/")
End Sub




Public Sub TortoiseHGAdd()

    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Add
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGCommit()

    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Commit
    Set Ver = Nothing

End Sub

Public Sub TortoiseHGRevert()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Revert
    
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGLog()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Log
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGCleanup()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Cleanup
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGVer()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Ver
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGHelp()

    If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    Call GoURL("https://tortoisehg-ja.readthedocs.io/ja/latest/")

End Sub
Public Sub TortoiseHGTag()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Tag
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGPush()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Push
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGRevisionGraph()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.RevisionGraph
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGDiff()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Diff
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGBrouser()
    Dim Ver As IVersion
    
    Set Ver = New TorToiseHG
    Ver.Brouser
    Set Ver = Nothing

End Sub
Public Sub TortoiseHGWeb()
    If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    Call GoURL("https://tortoisehg.bitbucket.io/ja/?")
End Sub
