'-------------------------------------------------------------------------------
' 同名ブックを参照用に開くスクリプト
' 
' rlxAliasOpen.vbs
' 
' Copyright (c) 2018 Y.Watanabe
' 
' This software is released under the MIT License.
' http://opensource.org/licenses/mit-license.php
'-------------------------------------------------------------------------------
' 動作確認 : Windows 7 + Excel 2016 / Windows 10 + Excel 2016
' コマンドライン
'   /install   ：インストールします。
'   /uninstall ：アンインストールします。
'-------------------------------------------------------------------------------
' 修正履歴
'   1.1.0 同名ブックを参照用に開き比較(&F)コマンドを追加。
'   1.0.0 新規作成
'-------------------------------------------------------------------------------
Option Explicit

    Const C_TITLE = "RelaxTools-Addin"
    Const C_REF = "（参照用）"
    Const C_COMPARE = "/C"
    Const C_INSTALL = "/RUNINSTALL"
    Const C_UNINSTALL = "/RUNUNINSTALL"

    Dim strActBook
    Dim strTmpBook
    Dim strFile
    Dim FS, v, varExt, k

    Set FS = CreateObject("Scripting.FileSystemObject")

    If WScript.Arguments.Count > 0 Then

        v = WScript.Arguments(0)

        Select Case UCase(v)
            Case "/INSTALL"
                '自分自身を管理者権限で実行
                With CreateObject("Shell.Application")
                    .ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ " & C_INSTALL, "", "runas"
                End With
                WScript.Quit

            Case "/UNINSTALL"
                '自分自身を管理者権限で実行
                With CreateObject("Shell.Application")
                    .ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ " & C_UNINSTALL, "", "runas"
                End With
                WScript.Quit
                
            Case C_INSTALL
                On Error Resume Next
                Err.Clear
                With WScript.CreateObject("WScript.Shell")
                    'ブック名を変更して開く
                    varExt = Array("Excel.Sheet.8", "Excel.Sheet.12", "Excel.SheetMacroEnabled.12")
                    For Each k In varExt
                       .RegWrite "HKCR\" & k & "\shell\rlxAliasOpen\","同名ブックを参照用に開く(&B)", "REG_SZ"
                       .RegWrite "HKCR\" & k & "\shell\rlxAliasOpen\command\","""" & FS.GetSpecialFolder(1) & "\wscript.exe"" """ & .SpecialFolders("AppData") & "\RelaxTools-Addin\rlxAliasOpen.vbs"" ""%1""", "REG_SZ"
                       .RegWrite "HKCR\" & k & "\shell\rlxAliasOpenDiff\","同名ブックを参照用に開き比較(&F)", "REG_SZ"
                       .RegWrite "HKCR\" & k & "\shell\rlxAliasOpenDiff\command\","""" & FS.GetSpecialFolder(1) & "\wscript.exe"" """ & .SpecialFolders("AppData") & "\RelaxTools-Addin\rlxAliasOpen.vbs"" """ & C_COMPARE & """ ""%1""", "REG_SZ"
                    Next            
                End With
                If Err.Number = 0 Then
                    MsgBox "レジストリを更新しました。", vbInformation + vbOkOnly, C_TITLE
                Else
                    MsgBox "エラーが発生しました。", vbCritical + vbOkOnly, C_TITLE
                End IF                
            Case C_UNINSTALL
                On Error Resume Next
                Err.Clear
                With WScript.CreateObject("WScript.Shell")
                    'ブック名を変更して開く
                    varExt = Array("Excel.Sheet.8", "Excel.Sheet.12", "Excel.SheetMacroEnabled.12")
                    For Each k In varExt
                       .RegDelete "HKCR\" & k & "\shell\rlxAliasOpen\command\"
                       .RegDelete "HKCR\" & k & "\shell\rlxAliasOpen\"
                       .RegDelete "HKCR\" & k & "\shell\rlxAliasOpenDiff\command\"
                       .RegDelete "HKCR\" & k & "\shell\rlxAliasOpenDiff\"
                    Next            
                End With
                'MsgBox "アンインストールしました。", vbInformation + vbOkOnly, C_TITLE
                
            Case C_COMPARE
                '比較モード
                If WScript.Arguments.Count > 1 Then
                    v = WScript.Arguments(1)
                    ExecExcel v, True
                Else
                    MsgBox "ファイル名が設定されていません。", vbInformation + vbOkOnly, C_TITLE 
                End If
            Case Else
                '通常モード
                ExecExcel v, False
        End Select

    End If
    
    Set FS = Nothing

'--------------------------------------------------------------
'　同名ブックを開く
'--------------------------------------------------------------
Sub ExecExcel(v, c)

    Dim XL, WB, W2, blnFind

    strActBook = v
    strTmpBook = rlxGetTempFolder() & C_REF & FS.GetFileName(v)
    FS.CopyFile strActBook, strTmpBook

    Err.Clear
    On Error Resume Next
    Set XL = GetObject(,"Excel.Application")
    If Err.Number = 0 Then

        Set WB = XL.Workbooks.Open(strTmpBook,,1)
        
        '比較モード
        If c Then
            blnFind = False
            For Each W2 In XL.Workbooks
                If W2.Name = FS.GetFileName(v) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If blnFind Then
                '比較のためＡ１にする
                setAllA1 WB
                setAllA1 W2

                '比較
                WB.Activate
                WB.Application.Windows.CompareSideBySideWith FS.GetFileName(v)
                W2.Activate
            Else
                MsgBox "比較先のブックが見つかりません。", vbInformation + vbOkOnly, C_TITLE 
            End If

        Else
            WB.Activate
        End If
    Else
        MsgBox "Excelを起動していないと実行できません。", vbInformation + vbOkOnly, C_TITLE 
    End If

End Sub
'--------------------------------------------------------------
'　テンポラリフォルダ取得
'--------------------------------------------------------------
Public Function rlxGetTempFolder() 

    On Error Resume Next
    
    Dim strFolder
    
    rlxGetTempFolder = ""
    
    With FS
    
        strFolder = rlxGetAppDataFolder & "Temp"
        
        If .FolderExists(strFolder) Then
        Else
            .createFolder strFolder
        End If
        
        rlxGetTempFolder = .BuildPath(strFolder, "\")
        
    End With
    

End Function

'--------------------------------------------------------------
'　アプリケーションフォルダ取得
'--------------------------------------------------------------
Function rlxGetAppDataFolder() 

    On Error Resume Next
    
    Dim strFolder
    
    rlxGetAppDataFolder = ""
    
    With FS
    
        strFolder = .BuildPath(CreateObject("Wscript.Shell").SpecialFolders("AppData"), C_TITLE)
        
        If .FolderExists(strFolder) Then
        Else
            .createFolder strFolder
        End If
        
        rlxGetAppDataFolder = .BuildPath(strFolder, "\")
        
    End With

End Function
'--------------------------------------------------------------
'　すべてのシートの選択位置をＡ１にセット
'--------------------------------------------------------------
Sub setAllA1(WB)

    On Error Resume Next

    WB.Application.ScreenUpdating = False

    Dim WS
    Dim lngPercent
 
    lngPercent = 100
  
    For Each WS In WB.Worksheets
        If WS.visible = -1 Then
            WS.Activate
            WS.Range("A1").Activate
            WB.Windows(1).ScrollRow = 1
            WB.Windows(1).ScrollColumn = 1
            
            WB.Windows(1).Zoom = lngPercent

        End If
    Next

    For Each WS In WB.Worksheets
        If WS.visible = -1 Then
            WS.Select
            Exit For
        End If
    Next
    
    WB.Application.ScreenUpdating = True
    
End Sub

