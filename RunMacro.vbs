' -------------------------------------------------------------------------------
' RelaxTools-Addin 別プロセス実行 Ver.1.0.0
' -------------------------------------------------------------------------------
' 修正
'   1.0.0 新規作成
' -------------------------------------------------------------------------------
On Error Resume Next
With CreateObject("Excel.Application")
    .Workbooks.Open CreateObject("WScript.Shell").SpecialFolders("Appdata") & "\Microsoft\Addins\Relaxtools.xlam"
    IF Wscript.Arguments.Count > 0 THEN
        .Run Wscript.Arguments(0)
    ELSE
        .Quit
    END IF
End With

