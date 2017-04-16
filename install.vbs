' -------------------------------------------------------------------------------
' RelaxTools-Addin インストールスクリプト Ver.1.0.3
' -------------------------------------------------------------------------------
' 参考サイト
' ある SE のつぶやき
' VBScript で Excel にアドインを自動でインストール/アンインストールする方法
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
' 修正
'   1.0.3 images フォルダをコピーするように修正。
'   1.0.2 Windows Update にて インターネットより取得したアドインファイルが Excel にて読み込まれない場合に対応。
'         警告とプロパティウィンドウを表示して「ブロック解除」をお願いするようにした。
' -------------------------------------------------------------------------------
On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin
Dim imageFolder

'アドイン情報を設定 
addInName = "RelaxTools Addin" 
addInFileName = "Relaxtools.xlam"

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

IF Not objFileSys.FileExists(addInFileName) THEN
   MsgBox "Zipファイルを展開してから実行してください。", vbExclamation, addInName 
   WScript.Quit 
END IF

'インストール先パスの作成 
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName] 
strPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\"
installPath = strPath  & addInFileName
imageFolder = objWshShell.SpecialFolders("Appdata") & "\RelaxTools-Addin\"

IF MsgBox(addInName & " をインストールしますか？" & vbCrLf &  "Version 4.0.0 以降とそれ以前では設定が引き継がれませんのでご了承ください。", vbYesNo + vbQuestion, addInName) = vbNo Then 
  WScript.Quit 
End IF

'ファイルコピー(上書き) 
objFileSys.CopyFile  addInFileName ,installPath , True

'イメージフォルダをコピー(上書き) 
objFileSys.CopyFolder  "Source\customUI\images" ,imageFolder , True

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

  Set objFolder = objShell.NameSpace(strPath)
  Set objFile = objFolder.ParseName(addInFileName)
  objFile.InvokeVerb("properties")
  MsgBox "インターネットから取得したファイルはExcelよりブロックされる場合があります。" & vbCrlf & "プロパティウィンドウを開きますので「ブロックの解除」を行ってください。" & vbCrLf & vbCrLf & "プロパティに「ブロックの解除」が表示されない場合は特に操作の必要はありません。", vbExclamation, addInName 

ELSE 
   MsgBox "エラーが発生しました。" & vbCrLF & "Excelが起動している場合は終了してください。", vbExclamation, addInName 
    WScript.Quit 
End IF

If MsgBox("エクスプローラ右クリック(Excelの読み取り専用)を有効にしますか？" & vbCrLf & "実行には管理者権限が必要です。", vbYesNo + vbQuestion, "読み取り専用有効化") = vbNo Then 
    WScript.Quit 
End IF

objWshShell.Run "ExcelReadOnly.vbs"

Set objWshShell = Nothing 

