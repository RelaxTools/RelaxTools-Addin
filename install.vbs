On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin

'アドイン情報を設定 
addInName = "RelaxTools Addin" 
addInFileName = "Relaxtools.xlam"

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

IF Not objFileSys.FileExists(addInFileName) THEN
   MsgBox "Zipファイルを解凍してから実行してください。", vbExclamation, addInName 
   WScript.Quit 
END IF

IF MsgBox(addInName & " をインストールしますか？", vbYesNo + vbQuestion, addInName) = vbNo Then 
  WScript.Quit 
End IF


'インストール先パスの作成 
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName] 
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'ファイルコピー(上書き) 
objFileSys.CopyFile  addInFileName ,installPath , True

Set objFileSys = Nothing

'Excel インスタンス化 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Workbooks.Add

'アドイン登録 
Set objAddin = objExcel.AddIns.Add(installPath, True) 
objAddin.Installed = True

'Excel 終了 
objExcel.Quit
Set objAddin = Nothing 
Set objExcel = Nothing

IF Err.Number = 0 THEN 
   MsgBox "アドインのインストールが終了しました。", vbInformation, addInName 
ELSE 
   MsgBox "エラーが発生しました。" & vbCrLF & "Excelが起動している場合は終了してください。", vbExclamation, addInName 
    WScript.Quit 
End IF

If MsgBox("エクスプローラ右クリック(Excelの読み取り専用)を有効にしますか？" & vbCrLf & "実行には管理者権限が必要です。", vbYesNo + vbQuestion, "読み取り専用有効化") = vbNo Then 
    WScript.Quit 
End IF

objWshShell.Run "ExcelReadOnly.vbs"

Set objWshShell = Nothing 

