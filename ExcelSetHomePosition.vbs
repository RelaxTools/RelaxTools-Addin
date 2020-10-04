'-------------------------------------------------------------------------------
' Excelファイルのカーソルをホームポジションに設定
' 
' ExcelSetHomePosition.vbs
' Version 1.0.0
' 
' Copyright (c) 2015 Y.Watanabe
' 
' This software is released under the MIT License.
' http://opensource.org/licenses/mit-license.php
'-------------------------------------------------------------------------------
' 動作確認 : Windows 7 + Excel 2010 / Windows 8 + Excel 2013
'-------------------------------------------------------------------------------
' for Used
' (1) ホームポジション設定するExcelファイルのフォルダにこのスクリプトを配置する。
' (2) スクリプトの「拡張子」「読み取りパスワード」を必要に応じて書き換える。
' (3) スクリプトを実行する。
' (4) 結果をテキストファイルで表示する。
' 
'-------------------------------------------------------------------------------
    Option Explicit

    Dim objFs, strMsg, SH
    Dim objDic, XL, WB, FL, LogName
    dim varPatterns, strKey, varPass, p
    Dim IE
    Dim strTitle
    
    strTitle = "ホームポジション設定"
    
    If MsgBox("同フォルダ以下のExcelファイルをホームポジション設定します。" & vbCrLf & "よろしいですか？" & VbCrLf & VbCrLf & "☆お約束☆" & vbCrLf & "Excelファイルは事前にバックアップしてください。", vbYesNo + vbQuestion, strTitle) = vbNo Then 
        WScript.Quit 
    End IF

    Set IE = WScript.CreateObject("InternetExplorer.Application")
 
    IE.Navigate "about:blank"
    Do While IE.busy
        WScript.Sleep(100)
    Loop
    Do While IE.Document.readyState <> "complete"
        WScript.Sleep(100)
    Loop
    IE.Document.body.innerHTML = "<b id=""msg"">ホームポジション設定中です<br>しばらくお待ち下さい...</b>"
    IE.AddressBar = False
    IE.ToolBar = False
    IE.StatusBar = False
    IE.Height = 120
    IE.Width = 300
    IE.Left = 0
    IE.Top = 0
    IE.Document.Title = strTitle
    IE.Visible = True
    
    On Error Resume Next

    Set objFs =  WScript.CreateObject("Scripting.FileSystemObject")
    Set objDic = WScript.CreateObject("Scripting.Dictionary")
    
    '--------------------------------------------------------------
    ' 処理を行う拡張子を正規表現で記述
    '--------------------------------------------------------------
    varPatterns = Array("\.xls$", "\.xlsx$", "\.xlsm$")
    
    '--------------------------------------------------------------
    ' 読み取りパスワードがある場合はここに記述(複数指定可)
    '--------------------------------------------------------------
    varPass = Array("", "", "")
    
    FileSearch objFs, objFs.GetParentFolderName(WScript.ScriptFullName), varPatterns, objDic

    LogName = objFs.GetBaseName(WScript.ScriptFullName) & ".txt"
    Set FL = objFs.CreateTextFile(LogName)

    FL.WriteLine "☆=ホームポジション設定 開始(" & Now() & ")☆="
    FL.WriteLine "処理ファイル数:" & objDic.Count

    If objDic.Count > 0 Then
        
        Set XL = WScript.CreateObject("Excel.Application")

        For Each strKey In objDic.Keys
        
            'パスワード指定の場合
            For Each p In varPass
                Err.Clear
                Set WB = XL.WorkBooks.Open(objDic(strKey),,False,,p,"",True,,,False)
                If Err.Number = 0 Then
                    Exit For
                End If
            Next
            
            Select Case True
                Case Err.Number <> 0
                    FL.WriteLine "エラー => " & objDic(strKey)
                    FL.WriteLine "          " & Err.Description
                    
                Case WB.ReadOnly 
  	                FL.WriteLine "エラー => " & objDic(strKey)
                    FL.WriteLine "          ブックが読み取り専用です"
                    
                Case Else
                    setAllA1 WB

                    XL.DisplayAlerts = False
                    WB.Save
                
                    If Err.Number <> 0 Or WB.Saved = False Then
                        FL.WriteLine "エラー => " & objDic(strKey)
                        FL.WriteLine "          " & Err.Description
                    Else
                        FL.WriteLine "処理済 => " & objDic(strKey)
                    End If
                
                    XL.DisplayAlerts = True
            End Select
            
            'インスタンスがあれば Close
            If Not IsNothing(WB) Then
                WB.Close
                Set WB = Nothing
            End If
        Next

        XL.Quit

        Set XL = Nothing

    End If

    FL.WriteLine "☆=ホームポジション設定 終了(" & Now() & ")☆="
    FL.Close
    Set FL = Nothing

    Set objDic = Nothing
    Set objFs =  Nothing

    With CreateObject("Shell.Application")
        .ShellExecute(LogName)
    End With

    IE.Quit
    'MsgBox "処理が完了しました。", vbInformation + VbOkOnly, strTitle

'--------------------------------------------------------------
'　すべてのシートの選択位置をＡ１にセット
'--------------------------------------------------------------
Sub setAllA1(WB)

    Dim WS
    Dim WD

    For Each WS In WB.Worksheets
        If WS.visible Then
            WS.Activate
'A1セットの際に複数セルの選択が解除されない問題 #65 #66
            WS.Range("A1").Select
            WB.Windows(1).ScrollRow = 1
            WB.Windows(1).ScrollColumn = 1
            WB.Windows(1).Zoom = 100
        End If
    Next

    '表示中の１枚目にする。
    For Each WS In WB.Worksheets
        If WS.visible  Then
            WS.Select
            Exit For
        End If
    Next

End Sub

'--------------------------------------------------------------
'　サブフォルダ検索
'--------------------------------------------------------------
Private Sub FileSearch(objFs, strPath, varPatterns, objDic)

    Dim objfld
    Dim objfl
    Dim objSub
    Dim f, objRegx
    
    Set objfld = objFs.GetFolder(strPath)

    'ファイル名取得
    For Each objfl In objfld.files
    
        Dim blnFind
        blnFind = False

	    Set objRegx = CreateObject("VBScript.RegExp")
        For Each f In varPatterns
            objRegx.Pattern = f
            If objRegx.Test(objfl.name) Then
                blnFind = True
                Exit For
            End If
        Next
	    Set objRegx = Nothing
        
        If blnFind Then
            objDic.Add objFs.BuildPath(objfl.ParentFolder.Path, objfl.name), objFs.BuildPath(objfl.ParentFolder.Path, objfl.name)
        End If
    Next
    
    'サブフォルダ検索あり
    For Each objSub In objfld.SubFolders
        FileSearch objFs, objSub.Path, varPatterns, objDic
    Next

End Sub
