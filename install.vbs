' -------------------------------------------------------------------------------
' RelaxTools-Addin インストールスクリプト Ver.1.0.6
' -------------------------------------------------------------------------------
' 参考サイト
' ある SE のつぶやき
' VBScript で Excel にアドインを自動でインストール/アンインストールする方法
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
' 修正
'   1.0.6 インストールパスを Application.UserLibraryPath を利用するように修正。
'   1.0.5 同名ブックを参照用に開くVBSをインストールするよう修正。
'   1.0.4 マルチプロセス用VBSが不要になったので削除。
'   1.0.3 マルチプロセス用VBSをコピーするよう修正。
'   1.0.3 images フォルダをコピーするように修正。
'   1.0.2 Windows Update にて インターネットより取得したアドインファイルが Excel にて読み込まれない場合に対応。
'         警告とプロパティウィンドウを表示して「ブロック解除」をお願いするようにした。
' -------------------------------------------------------------------------------
Option Explicit
On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin
Dim imageFolder
Dim appFile
Dim objWshShell
Dim objFileSys
Dim strPath
Dim objFolder
Dim objFile

'アドイン情報を設定 
addInName = "RelaxTools Addin" 
addInFileName = "Relaxtools.xlam"
appFile = "rlxAliasOpen.vbs"

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

IF Not objFileSys.FileExists(addInFileName) THEN
    MsgBox "Zipファイルを展開してから実行してください。", vbExclamation, addInName 
    WScript.Quit 
END IF

IF MsgBox(addInName & " をインストールしますか？" & vbCrLf &  "Version 4.0.0 以降とそれ以前では設定が引き継がれませんのでご了承ください。", vbYesNo + vbQuestion, addInName) = vbNo Then 
    WScript.Quit 
End IF

'Excel インスタンス化 
With CreateObject("Excel.Application") 

    'インストール先パスの作成 
    strPath = .UserLibraryPath
    imageFolder = objWshShell.SpecialFolders("Appdata") & "\RelaxTools-Addin\"

    'インストールフォルダがない場合は作成
    IF Not objFileSys.FolderExists(strPath) THEN
        objFileSys.CreateFolder(strPath)
    END IF

    installPath = strPath & addInFileName

    'ファイルコピー(上書き) 
    objFileSys.CopyFile  addInFileName ,installPath , True

    'イメージフォルダがない場合は作成
    IF Not objFileSys.FolderExists(imageFolder) THEN
        objFileSys.CreateFolder(imageFolder)
    END IF

    'イメージフォルダをコピー(上書き) 
    objFileSys.CopyFolder  "Source\customUI\images" ,imageFolder , True

    'ファイルをコピー(上書き) 
    objFileSys.CopyFile  appFile, imageFolder & appFile, True

    'アドイン登録 
    .Workbooks.Add
    Set objAddin = .AddIns.Add(installPath, True) 
    objAddin.Installed = True

    'Excel 終了 
    .Quit

End WIth

IF Err.Number = 0 THEN 
    MsgBox "アドインのインストールが終了しました。", vbInformation, addInName 

    'プロパティファイル表示
    CreateObject("Shell.Application").NameSpace(strPath).ParseName(addInFileName).InvokeVerb("properties")
    MsgBox "インターネットから取得したファイルはExcelよりブロックされる場合があります。" & vbCrlf & "プロパティウィンドウを開きますので「ブロックの解除」を行ってください。" & vbCrLf & vbCrLf & "プロパティに「ブロックの解除」が表示されない場合は特に操作の必要はありません。", vbExclamation, addInName 

ELSE 
    MsgBox "エラーが発生しました。" & vbCrLF & "Excelが起動している場合は終了してください。", vbExclamation, addInName 
    WScript.Quit 
End IF

If MsgBox("エクスプローラ右クリック(同名ブックを参照用に開く)を有効にしますか？" & vbCrLf & "実行には管理者権限が必要です。", vbYesNo + vbQuestion, addInName) <> vbNo Then 
    objWshShell.Run "rlxAliasOpen.vbs /install", 1, true
End IF

If MsgBox("エクスプローラ右クリック(Excelの読み取り専用)を有効にしますか？" & vbCrLf & "実行には管理者権限が必要です。", vbYesNo + vbQuestion, addInName) = vbNo Then 
    WScript.Quit 
End IF

objWshShell.Run "ExcelReadOnly.vbs", 1, true

Set objFileSys = Nothing
Set objWshShell = Nothing 
