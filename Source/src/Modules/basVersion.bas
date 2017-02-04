Attribute VB_Name = "basVersion"
Option Explicit
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
