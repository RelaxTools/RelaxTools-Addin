Attribute VB_Name = "basShowForm"
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

'--------------------------------------------------------------
'　セルの簡易編集
'--------------------------------------------------------------
Sub cellEdit()

    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    If Selection.CountLarge > 1 And Selection.CountLarge <> Selection(1, 1).MergeArea.count Then
        MsgBox "複数セル選択されています。セルは１つのみ選択してください。", vbExclamation + vbOKOnly, C_TITLE
        Exit Sub
    End If
    
    frmEdit.show
    
End Sub
'--------------------------------------------------------------
'　セルの簡易編集
'--------------------------------------------------------------
Sub cellSearch()

    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    frmSearchEx.txtSearch.Text = Replace(Replace(ActiveCell.Value, vbCrLf, "\n"), vbCr, "\n")
    frmSearchEx.txtSearch.SelStart = 0
    
    frmSearchEx.show
    
    
End Sub
'--------------------------------------------------------------
'　置換
'--------------------------------------------------------------
Sub replaceEx()

    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    frmSearchEx.schTab.Value = 1
    
    frmSearchEx.show
    
    
End Sub
'--------------------------------------------------------------
'　SQLの「美」整形設定画面
'--------------------------------------------------------------
Sub FormatSqlSetting()

    frmFormatSql.show
    
End Sub
'--------------------------------------------------------------
'　XMLの「美」整形設定画面
'--------------------------------------------------------------
Sub FormatXMLSetting()

    frmFormatXml.show
    
End Sub

'--------------------------------------------------------------
'　バックアップ設定画面
'--------------------------------------------------------------
Sub backupSetting()

    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    frmBackupSetting.show
    
End Sub
'--------------------------------------------------------------
'　拡張検索画面
'--------------------------------------------------------------
Sub searchEx()

    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    frmSearchEx.show
    
End Sub

'--------------------------------------------------------------
'　シート管理
'--------------------------------------------------------------
Sub execSheetManager()

    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    If ActiveWorkbook.ProtectStructure Then
        MsgBox "このブックは保護されているためシート管理は使用できません。", vbOKOnly + vbInformation, C_TITLE
        Exit Sub
    End If
    
    Dim WS As Object
    For Each WS In ActiveWorkbook.Sheets
        If WS.Name = "履歴" Then
            MsgBox "「履歴」ワークシートが存在するためシート管理は使用できません。", vbOKOnly + vbInformation, C_TITLE
            Exit Sub
        End If
    Next

    frmSheetManager.show

End Sub
'--------------------------------------------------------------
'　JAVAパッケージ配置
'--------------------------------------------------------------
Sub setJavaPackage()

    frmSetPackage.show

End Sub
'--------------------------------------------------------------
'　ツリー一覧作成画面
'--------------------------------------------------------------
Sub createLinkTreeIn()

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    frmTreeList.show
    
End Sub
'--------------------------------------------------------------
'　バージョン表示
'--------------------------------------------------------------
Sub dispVer()
    
    frmVersion.show

End Sub

'--------------------------------------------------------------
'　カーソル位置に指定されたフォルダの一覧を挿入します。
'--------------------------------------------------------------
Sub createFileListIn()

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    frmFileList.show vbModeless

End Sub
'--------------------------------------------------------------
'　かんたん罫線
'--------------------------------------------------------------
Sub KantanLine()

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    If TypeOf Selection Is Range Then
    Else
        Exit Sub
    End If

    frmGridText.show
    
End Sub
'--------------------------------------------------------------
'　CSV読み込み
'--------------------------------------------------------------
Sub loadCSV()

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    frmLoadCSV.show
    
End Sub
'--------------------------------------------------------------
'　html画面
'--------------------------------------------------------------
Sub convertHtml()

    If TypeOf Selection Is Range Then
    Else
        Exit Sub
    End If
    
    frmHtml.show vbModal
    
End Sub


'--------------------------------------------------------------
'　セル書式設定画面
'--------------------------------------------------------------
Sub documentSetting()

    frmDoc.show
    
End Sub
'--------------------------------------------------------------
'　ExcelファイルのGrep
'--------------------------------------------------------------
Sub excelGrep()

    frmGrep.show
    
End Sub
'--------------------------------------------------------------
'　ExcelファイルのGrep(マルチプロセス版)
'--------------------------------------------------------------
Sub excelGrepMulti()

    MultiProsess "excelGrepMultiShow"
    
End Sub
'--------------------------------------------------------------
'　ExcelファイルのGrep(マルチプロセス版)
'--------------------------------------------------------------
Sub excelGrepMultiShow()

    frmGrepMulti.show
    
End Sub
'--------------------------------------------------------------
'　Excelファイルのページ数取得
'--------------------------------------------------------------
Sub excelPage()

    frmPageList.show
    
End Sub
''--------------------------------------------------------------
''　ファイルのMessageDigestを求める
''--------------------------------------------------------------
'Sub getMessageDigest()
'
'    frmMessageDigest.Show
'
'End Sub
Sub reSelect()
    frmReSelect.show
End Sub
Sub showFavorite()
    frmFavorite.show
End Sub
'--------------------------------------------------------------
'　ワークシートの比較
'--------------------------------------------------------------
Sub compWorkSheets()

    frmComp.show
    
End Sub
Sub cellEditExtSetting()
    frmEditEx.show
End Sub
Sub A1SettingShow()
    frmA1Setting.show
End Sub
Sub electoricSetting()
    frmElectoric.show
End Sub
Sub hotkey()
    frmHotKey.show
End Sub

Sub sectionSettingShow()
    frmSectionList.show
End Sub
Sub crossSetting()
    Dim obj As Object
    lineOnAction obj, False
    frmCrossLine.show
End Sub
Sub showBz()
    frmStampBz.show
End Sub
Sub createFolderShow()
    frmCreateFolder.show
End Sub
Sub VBAStepCountShow()
    frmStepCount.show
End Sub
Sub execScreenShotSetting()
    frmScreenSetting.show
End Sub
Sub execSourceExport()
    frmSourceExport.show
End Sub
Sub execComboSetting()
    frmCombo.show
End Sub
Sub execDelStyle()
    frmStyle.show
End Sub
Sub execCopyScreenSetting()
    frmCopyScreen.show
End Sub

Sub execOptionSetting()
    frmCommonOption.show
End Sub
Sub scrollSetting()
    frmScroll.show
End Sub
Sub convertTextile()

    If TypeOf Selection Is Range Then
    Else
        Exit Sub
    End If

    frmRedmine.show
End Sub
Sub convertMarkdown()

    If TypeOf Selection Is Range Then
    Else
        Exit Sub
    End If
    
    frmMarkdown.show
End Sub
Sub showGrammer()
    frmGrammer.show
End Sub
Sub showInfo()
    frmInfo.show
End Sub
'Sub showHoldBook()
'    frmHoldBook.Show
'End Sub
Sub showCheckList()
    frmCheckList.show
End Sub

Sub execBinaryView()
    frmBinary.show
End Sub
Sub KanaSetting()
    frmKana.show
End Sub

Sub PickSetting()
    frmPickSetting.show
End Sub

Sub ShowReport()
    frmReport.show
End Sub
Sub contextMenu()
    frmContextMenu.show
End Sub
Sub showStaticCheck()
    frmStaticCheck.show
End Sub
Sub showMergeFile()
    frmMergeFile.show
End Sub
