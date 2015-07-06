'-------------------------------------------------------------------------------
' Excelファイルの右クリック「読み取り専用で開く」を有効にするスクリプト
' 
' ExcelReadOnly.vbs
' 
' Copyright (c) 2015 Y.Watanabe
' 
' This software is released under the MIT License.
' http://opensource.org/licenses/mit-license.php
'-------------------------------------------------------------------------------
' 動作確認 : Windows 7 + Excel 2010 / Windows 8 + Excel 2013
'-------------------------------------------------------------------------------
' 以下参考サイト

' 無題 - 右クリックメニュー「読み取り専用で開く」を表示する(Excel&Word) 
' https://sites.google.com/site/universeof/tips/openasreadonly'
'-------------------------------------------------------------------------------
Option Explicit

On Error Resume Next

If WScript.Arguments.Count = 0 Then

    '自分自身を管理者権限で実行
    With CreateObject("Shell.Application")
        .ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ dummy", "", "runas"
    End With
    
    WScript.Quit
    
End If

If MsgBox("エクスプローラ右クリック(Excelの読み取り専用)を有効にしますか？", vbYesNo + vbQuestion, "読み取り専用有効化") = vbNo Then 
    WScript.Quit 
End IF

With WScript.CreateObject("WScript.Shell")

    'シフトを押さなくてもメニューが表示されるようにするように「Extended」キーを削除
    .RegDelete "HKCR\Excel.Sheet.8\shell\OpenAsReadOnly\Extended"
    .RegDelete "HKCR\Excel.Sheet.12\shell\OpenAsReadOnly\Extended"
    .RegDelete "HKCR\Excel.SheetMacroEnabled.12\shell\OpenAsReadOnly\Extended"

    Err.Clear

    '読み取り専用を有効にする
    .RegWrite "HKCR\Excel.Sheet.8\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.Sheet.12\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.SheetMacroEnabled.12\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"

End With

If Err.Number = 0 Then
    MsgBox "正常に関連付けを変更しました。", vbInformation + vbOkOnly, "読み取り専用有効化"
Else
    MsgBox "エラーが発生しました。", vbCritical + vbOkOnly, "読み取り専用有効化"
End IF

 