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
    
    frmEdit.Show
    
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
    
    frmSearchEx.Show
    
    
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
    
    frmSearchEx.Show
    
    
End Sub
'--------------------------------------------------------------
'　SQLの「美」整形設定画面
'--------------------------------------------------------------
Sub FormatSqlSetting()

    frmFormatSql.Show
    
End Sub
'--------------------------------------------------------------
'　XMLの「美」整形設定画面
'--------------------------------------------------------------
Sub FormatXMLSetting()

    frmFormatXml.Show
    
End Sub

'--------------------------------------------------------------
'　バックアップ設定画面
'--------------------------------------------------------------
Sub backupSetting()

    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    frmBackupSetting.Show
    
End Sub
'--------------------------------------------------------------
'　拡張検索画面
'--------------------------------------------------------------
Sub searchEx()

    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    frmSearchEx.Show
    
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

    frmSheetManager.Show

End Sub
'--------------------------------------------------------------
'　JAVAパッケージ配置
'--------------------------------------------------------------
Sub setJavaPackage()

    frmSetPackage.Show

End Sub
'--------------------------------------------------------------
'　ツリー一覧作成画面
'--------------------------------------------------------------
Sub createLinkTreeIn()

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    frmTreeList.Show
    
End Sub
'--------------------------------------------------------------
'　バージョン表示
'--------------------------------------------------------------
Sub dispVer()
    
    frmVersion.Show
    
End Sub

'--------------------------------------------------------------
'　カーソル位置に指定されたフォルダの一覧を挿入します。
'--------------------------------------------------------------
Sub createFileListIn()

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    frmFileList.Show vbModeless

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

    frmGridText.Show
    
End Sub
'--------------------------------------------------------------
'　CSV読み込み
'--------------------------------------------------------------
Sub loadCSV()

    If ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    frmLoadCSV.Show
    
End Sub
'--------------------------------------------------------------
'　html画面
'--------------------------------------------------------------
Sub convertHtml()

    If TypeOf Selection Is Range Then
    Else
        Exit Sub
    End If
    
    frmHtml.Show vbModal
    
End Sub


'--------------------------------------------------------------
'　セル書式設定画面
'--------------------------------------------------------------
Sub documentSetting()

    frmDoc.Show
    
End Sub
'--------------------------------------------------------------
'　ExcelファイルのGrep
'--------------------------------------------------------------
Sub excelGrep()

    frmGrep.Show
    
End Sub
'--------------------------------------------------------------
'　Excelファイルのページ数取得
'--------------------------------------------------------------
Sub excelPage()

    frmPageList.Show
    
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
    frmReSelect.Show
End Sub
Sub showFavorite()
    frmFavorite.Show
End Sub
'--------------------------------------------------------------
'　ワークシートの比較
'--------------------------------------------------------------
Sub compWorkSheets()

    frmComp.Show
    
End Sub
Sub cellEditExtSetting()
    frmEditEx.Show
End Sub
Sub A1SettingShow()
    frmA1Setting.Show
End Sub
Sub electoricSetting()
    frmElectoric.Show
End Sub
Sub hotkey()
    frmHotKey.Show
End Sub

Sub sectionSettingShow()
    frmSectionList.Show
End Sub
Sub crossSetting()
    Dim obj As Object
    lineOnAction obj, False
    frmCrossLine.Show
End Sub
Sub showBz()
    frmStampBz.Show
End Sub
Sub createFolderShow()
    frmCreateFolder.Show
End Sub
Sub VBAStepCountShow()
    frmStepCount.Show
End Sub
Sub execScreenShotSetting()
    frmScreenSetting.Show
End Sub
Sub execSourceExport()
    frmSourceExport.Show
End Sub
Sub execComboSetting()
    frmCombo.Show
End Sub
Sub execDelStyle()
    frmStyle.Show
End Sub
Sub execCopyScreenSetting()
    frmCopyScreen.Show
End Sub

Sub execOptionSetting()
    frmCommonOption.Show
End Sub
Sub scrollSetting()
    frmScroll.Show
End Sub
Sub convertTextile()

    If TypeOf Selection Is Range Then
    Else
        Exit Sub
    End If

    frmRedmine.Show
End Sub
Sub convertMarkdown()

    If TypeOf Selection Is Range Then
    Else
        Exit Sub
    End If
    
    frmMarkdown.Show
End Sub
Sub showGrammer()
    frmGrammer.Show
End Sub
Sub showInfo()
    frmInfo.Show
End Sub
'Sub showHoldBook()
'    frmHoldBook.Show
'End Sub
Sub showCheckList()
    frmCheckList.Show
End Sub

Sub execBinaryView()
    frmBinary.Show
End Sub
Sub KanaSetting()
    frmKana.Show
End Sub

