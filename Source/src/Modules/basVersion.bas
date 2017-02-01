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
