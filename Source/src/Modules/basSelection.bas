Attribute VB_Name = "basSelection"
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


Sub SelectionUndo()

    Selection.Value = pvarSelectionBuffer

End Sub
'--------------------------------------------------------------
' 指定範囲選択(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowSelect()
    
    Dim obj As SelectionRowSelect
    
    Set obj = New SelectionRowSelect
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' SQL整形(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionFormatSql()
    
    Dim obj As SelectionFormatSql
    
    Set obj = New SelectionFormatSql
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' XML整形(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionFormatXML()
    
    Dim obj As SelectionFormatXML
    
    Set obj = New SelectionFormatXML
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 式の再設定(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToFormula()
    
    Dim obj As SelectionToFormula
    
    Set obj = New SelectionToFormula
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 列のマージ(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowMergeCol()
    
    Dim obj As SelectionRowMergeCol
    
    Set obj = New SelectionRowMergeCol
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' INSERT文生成(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowMakeSQLInsert()
    
    Dim obj As SelectionRowMakeSQLInsert
    
    Set obj = New SelectionRowMakeSQLInsert
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' CRLF削除(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemoveCrLf()
    
    Dim obj As SelectionRemoveCrLf
    
    Set obj = New SelectionRemoveCrLf
    
    obj.Run
    
    Set obj = Nothing
    
End Sub

'--------------------------------------------------------------
' 指定範囲選択(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowSelectCell()
    
    Dim obj As SelectionRowSelectCell
    
    Set obj = New SelectionRowSelectCell
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 指定範囲シフト(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowShiftSelect()
    
    Dim obj As SelectionRowShiftSelect
    
    Set obj = New SelectionRowShiftSelect
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 空白除去(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionTrimCell()

    Dim obj As SelectionTrimCell
    
    Set obj = New SelectionTrimCell
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 左指定文字数削除(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemoveLeftString()
    
    Dim obj As SelectionRemoveLeftString
    
    Set obj = New SelectionRemoveLeftString
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 指定文字数以降削除(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemoveLeftToString()
    
    Dim obj As SelectionRemoveLeftToString
    
    Set obj = New SelectionRemoveLeftToString
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 指定文字以前降削除(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemoveRightToString()
    
    Dim obj As SelectionRemoveRightToString
    
    Set obj = New SelectionRemoveRightToString
    
    obj.Run
    
    Set obj = Nothing
    
End Sub '--------------------------------------------------------------
' 右指定文字数削除(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemoveRightString()

    Dim obj As SelectionRemoveRightString
    
    Set obj = New SelectionRemoveRightString
    
    obj.Run
    
    Set obj = Nothing
    

End Sub
'--------------------------------------------------------------
' 行頭文字追加(SelectionAllFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionAllInsertHead()

    Dim obj As SelectionInsertHead
    
    Set obj = New SelectionInsertHead
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 行末文字追加(SelectionAllFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionAllInsertBottom()

    Dim obj As SelectionInsertBottom
    
    Set obj = New SelectionInsertBottom
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 行中文字追加(SelectionAllFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionAllInsertMiddle()

    Dim obj As SelectionInsertMiddle
    
    Set obj = New SelectionInsertMiddle
    
    obj.Run
    
    Set obj = Nothing
    
End Sub '--------------------------------------------------------------
' 書式クリア(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionInitialize()

    Dim obj As SelectionInitialize
    
    Set obj = New SelectionInitialize
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 小文字変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToLower()

    Dim obj As SelectionToLower
    
    Set obj = New SelectionToLower
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 大文字変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToUpper()

    Dim obj As SelectionToUpper
    
    Set obj = New SelectionToUpper
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 半角変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToHankaku()

    Dim obj As SelectionToHankaku
    
    Set obj = New SelectionToHankaku
    
    obj.Run
    
    Set obj = Nothing


End Sub
'--------------------------------------------------------------
' 全角変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToZenkaku()

    Dim obj As SelectionToZenkaku
    
    Set obj = New SelectionToZenkaku
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 全角ひらがな変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToHiragana()

    Dim obj As SelectionToHiragana
    
    Set obj = New SelectionToHiragana
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 全角カタカナ変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToKatakana()

    Dim obj As SelectionToKatakana
    
    Set obj = New SelectionToKatakana
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 電子納品変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToElectoric()

    Dim obj As SelectionToElectoric
    
    Set obj = New SelectionToElectoric
    
    obj.Run
    
    Set obj = Nothing

End Sub '--------------------------------------------------------------
' 単語の先頭のみ大文字(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToPropercase()
  
    Dim obj As SelectionToPropercase
    
    Set obj = New SelectionToPropercase
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 8ケタ文字を日付に変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToDate8()

    Dim obj As SelectionToDate8
    
    Set obj = New SelectionToDate8
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 6ケタ文字(YYMMDD)を日付に変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToDate6()

    Dim obj As SelectionToDate6
    
    Set obj = New SelectionToDate6
    
    obj.Run
    
    Set obj = Nothing


End Sub
'--------------------------------------------------------------
' 6ケタ文字(MMDDYY)を日付に変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToDate6mdy()

    Dim obj As SelectionToDate6mdy
    
    Set obj = New SelectionToDate6mdy
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 6ケタ文字(DDMMYY)を日付に変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToDate6dmy()

    Dim obj As SelectionToDate6dmy
    
    Set obj = New SelectionToDate6dmy
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択している文字列でフォルダ作成(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCreateFolder()

    Dim obj As SelectionCreateFolder
    
    Set obj = New SelectionCreateFolder
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' ファイル名のみにする(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemovePath()

    Dim obj As SelectionRemovePath
    
    Set obj = New SelectionRemovePath
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 拡張子を削除(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemoveExt()

    Dim obj As SelectionRemoveExt
    
    Set obj = New SelectionRemoveExt
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' フォルダ名のみにする(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemoveFilename()

    Dim obj As SelectionRemoveFilename
    
    Set obj = New SelectionRemoveFilename
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' セルの背景色をＲＧＢで取得(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckCellColor()

    Dim obj As SelectionCheckCellColor
    
    Set obj = New SelectionCheckCellColor
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 一意チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckUniq()

    Dim obj As SelectionCheckUniq
    
    Set obj = New SelectionCheckUniq
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 一意チェック（行対応）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowCheckUniq()

    Dim obj As SelectionRowCheckUniq
    
    Set obj = New SelectionRowCheckUniq
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 一意チェック（範囲）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowCheckFromTo()

    Dim obj As SelectionRowCheckFromTo
    
    Set obj = New SelectionRowCheckFromTo
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 重複した項目を削除(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionSetUniq()
    
    Dim obj As SelectionSetUniq
    
    Set obj = New SelectionSetUniq
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 重複した項目を削除行対応(SelectionFrameWorkBox使用)
'--------------------------------------------------------------
Sub execSelectionRowSetUniq()
    
    Dim obj As SelectionRowSetUniq
    
    Set obj = New SelectionRowSetUniq
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' DBフィールド名からJAVA名(get)(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToJavaStringGet()

    'Java文字列へ
    Dim obj As SelectionToJavaString
    
    Set obj = New SelectionToJavaString
    
    obj.setType = "get"
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' DBフィールド名からJAVA名(set)(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToJavaStringSet()

    'Java文字列へ
    Dim obj As SelectionToJavaString
    
    Set obj = New SelectionToJavaString
    
    obj.setType = "set"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' DBフィールド名からJAVA名(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToJavaString()

    'Java文字列へ
    Dim obj As SelectionToJavaString
    
    Set obj = New SelectionToJavaString
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' DBフィールド名からJAVA名(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToJavaStringU()

    'Java文字列へ
    Dim obj As SelectionToJavaStringU
    
    Set obj = New SelectionToJavaStringU
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' JAVA名からDBフィールド名(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToDBString()

    'Java文字列へ
    Dim obj As SelectionToDBString
    
    Set obj = New SelectionToDBString
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' かんたん表（方眼紙用）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowDrawGridHogan()

    Dim obj As SelectionRowDrawGrid
    
    Set obj = New SelectionRowDrawGrid
    obj.HeadColor = 16764057
    obj.EvenColor = -1
    obj.Custom = False
    
    obj.HoganMode = True
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' かんたん表（シンプル）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowDrawGrid()

    Dim obj As SelectionRowDrawGrid
    
    Set obj = New SelectionRowDrawGrid
    obj.HeadColor = 16764057
    obj.EvenColor = -1
    obj.Custom = False
    
'    obj.HoganMode = True
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' かんたん表（シンプル）行１列１(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowDrawGrid1And1()

    Dim obj As SelectionRowDrawGrid

    Set obj = New SelectionRowDrawGrid
    obj.HeadColor = 16764057
    obj.EvenColor = -1
    obj.Custom = False

    obj.HeadLine = 1
    obj.ColLine = 1

'    obj.HoganMode = True

    obj.Run

    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' かんたん表（標準・ヘッダ２行）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowDrawGrid2()

    Dim obj As SelectionRowDrawGrid
    
    Set obj = New SelectionRowDrawGrid
    obj.HeadColor = 16764057
    obj.EvenColor = -1
    obj.Custom = False
    obj.HeadLine = 2
    
'    obj.HoganMode = True
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' かんたん表（標準・ヘッダ３行）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowDrawGrid3()

    Dim obj As SelectionRowDrawGrid
    
    Set obj = New SelectionRowDrawGrid
    obj.HeadColor = 16764057
    obj.EvenColor = -1
    obj.Custom = False
    obj.HeadLine = 3
    
'    obj.HoganMode = True
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' かんたん表（しましまブルー）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowDrawGridBlue()

    Dim obj As SelectionRowDrawGrid
    
    Set obj = New SelectionRowDrawGrid
    
    obj.HeadColor = 16764057
    obj.EvenColor = 16777164
    obj.Custom = False
    
'    obj.HoganMode = True
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' かんたん表（しましまベージュ）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowDrawGridBeige()

    Dim obj As SelectionRowDrawGrid
    
    Set obj = New SelectionRowDrawGrid
    
    obj.HeadColor = 10079487
    obj.EvenColor = 10092543
    obj.Custom = False
    
'    obj.HoganMode = True
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' かんたん表（カスタム）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowDrawGridCustom()

    Dim obj As SelectionRowDrawGrid
    
    Set obj = New SelectionRowDrawGrid
    
    obj.HeadColor = 16764057
    obj.EvenColor = 16777164
    obj.Custom = True
    
'    obj.HoganMode = True
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択（奇数行）(SelectionRowFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRowSelectOdd()

    Dim obj As SelectionRowSelectOdd
    
    Set obj = New SelectionRowSelectOdd
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択（奇数列）(SelectionColFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionColSelectOdd()

    Dim obj As SelectionColSelectOdd
    
    Set obj = New SelectionColSelectOdd
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' ハイパーリンクの削除(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionRemoveHyperlink()

    Dim obj As SelectionRemoveHyperlink
    
    Set obj = New SelectionRemoveHyperlink
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択しているセルの数(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckCount()

    Dim obj As SelectionCheckCount
    
    Set obj = New SelectionCheckCount
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択しているセルの文字数(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckSize()

    Dim obj As SelectionCheckSize
    
    Set obj = New SelectionCheckSize
    
    obj.CountType = SelectionCheckSizeConstants.CountTypeSJIS
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択しているセルの文字数(SelectionFrameWork使用)UTF8
'--------------------------------------------------------------
Sub execSelectionCheckSizeUTF8()

    Dim obj As SelectionCheckSize
    
    Set obj = New SelectionCheckSize
    
    obj.CountType = SelectionCheckSizeConstants.CountTypeUTF8
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択しているセルの文字数(SelectionFrameWork使用)UTF16
'--------------------------------------------------------------
Sub execSelectionCheckSizeUTF16()

    Dim obj As SelectionCheckSize
    
    Set obj = New SelectionCheckSize
    
    obj.CountType = SelectionCheckSizeConstants.CountTypeUTF16
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択範囲で値のあるセルを選択(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionSelectValueCell()

    Dim obj As SelectionSelectValueCell
    
    Set obj = New SelectionSelectValueCell
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 選択範囲で値のないセルを選択(SelectionAllFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionAllSelectEmptyCell()

    Dim obj As SelectionAllSelectEmptyCell
    
    Set obj = New SelectionAllSelectEmptyCell
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 値で更新(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToValue()

    Dim obj As SelectionToValue
    
    Set obj = New SelectionToValue
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 指数表記を文字列に変換(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionStringFormura()

    Dim obj As SelectionStringFormura
    
    Set obj = New SelectionStringFormura
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' ゼロ埋め(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionStringZeroPadding()

    Dim obj As SelectionStringZeroPadding
    
    Set obj = New SelectionStringZeroPadding
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' Luhnチェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckLuhn()

    Dim obj As SelectionCheckLuhn
    
    Set obj = New SelectionCheckLuhn
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' モジュラス１０チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckModulus10()

    Dim obj As SelectionCheckModulus10
    
    Set obj = New SelectionCheckModulus10
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' モジュラス１１チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckModulus11_10_2()

    Dim obj As SelectionCheckModulus11_10_2
    
    Set obj = New SelectionCheckModulus11_10_2
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' モジュラス１１チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckModulus11_2_7()

    Dim obj As SelectionCheckModulus11_2_7
    
    Set obj = New SelectionCheckModulus11_2_7
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' モジュラス１１チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckModulus11_Pref()

    Dim obj As SelectionCheckModulus11_Pref
    
    Set obj = New SelectionCheckModulus11_Pref
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' 数字チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckNumber()

    Dim obj As SelectionCheckNumber
    
    Set obj = New SelectionCheckNumber
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 英字チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckAlphabet()

    Dim obj As SelectionCheckAlphabet
    
    Set obj = New SelectionCheckAlphabet
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 英数字チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckAlphaAndNum()

    Dim obj As SelectionCheckAlphaAndNum
    
    Set obj = New SelectionCheckAlphaAndNum
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' 数値妥当性チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckNumeric()

    Dim obj As SelectionCheckNumeric
    
    Set obj = New SelectionCheckNumeric
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' 日付妥当性チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckDate()

    Dim obj As SelectionCheckDate
    
    Set obj = New SelectionCheckDate
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 半角文字列存在チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckAsc()

    Dim obj As SelectionCheckAsc
    
    Set obj = New SelectionCheckAsc
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 全角文字列存在チェック(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionCheckSJIS()

    Dim obj As SelectionCheckSJIS
    
    Set obj = New SelectionCheckSJIS
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' ネットワークドライブ→UNC(SelectionFrameWork使用)
'--------------------------------------------------------------
Sub execSelectionToUNC()

    Dim obj As SelectionToUNC
    
    Set obj = New SelectionToUNC
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 縦方向にマージ
'--------------------------------------------------------------
Sub execSelectionColMerge()

    Dim obj As SelectionColMerge
    
    Set obj = New SelectionColMerge
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' セルの最後に１行改行を付加する。
'--------------------------------------------------------------
Sub execSelectionLineFeedInsert()

    Dim obj As SelectionLineFeedInsert
    
    Set obj = New SelectionLineFeedInsert
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' セルの最後に１行改行を削除する。
'--------------------------------------------------------------
Sub execSelectionLineFeedDelete()

    Dim obj As SelectionLineFeedDelete
    
    Set obj = New SelectionLineFeedDelete
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' セルにハイフンを設定する。
'--------------------------------------------------------------
Sub execSelectionInsertHyphen()

    Dim obj As SelectionInsertStrInCell
    
    Set obj = New SelectionInsertStrInCell
    
    obj.InsertStr = "-"
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' セルにハイフンを設定する。
'--------------------------------------------------------------
Sub execSelectionInsertHyphenZen()

    Dim obj As SelectionInsertStrInCell
    
    Set obj = New SelectionInsertStrInCell
    
    obj.InsertStr = "－"
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 段落番号を振る
'--------------------------------------------------------------
Sub execSelectionSectionNumber()

    Dim obj As SelectionSectionNumber
    
    Set obj = New SelectionSectionNumber
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 段落番号を振る(左インデント)
'--------------------------------------------------------------
Sub execSelectionSectionNumberIndentL()

    Dim obj As SelectionSectionNumber
    
    Set obj = New SelectionSectionNumber
    
    obj.Indent = -1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 段落番号を振る(右インデント)
'--------------------------------------------------------------
Sub execSelectionSectionNumberIndentR()

    Dim obj As SelectionSectionNumber
    
    Set obj = New SelectionSectionNumber
    
    obj.Indent = 1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 右インデント
'--------------------------------------------------------------
Sub execSelectionAllIndentR()

    Dim obj As SelectionAllIndent
    
    Set obj = New SelectionAllIndent
    
    obj.Indent = 1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 左インデント
'--------------------------------------------------------------
Sub execSelectionAllIndentL()

    Dim obj As SelectionAllIndent
    
    Set obj = New SelectionAllIndent
    
    obj.Indent = -1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 左上がり罫線
'--------------------------------------------------------------
Sub execSelectionAllDiagonalUp()

'    Dim obj As SelectionAllDiagonalUp
'
'    Set obj = New SelectionAllDiagonalUp
'
'    obj.Run
'
'    Set obj = Nothing


    Dim e As SelectionFormatBoader
    
    Set e = New SelectionFormatBoader
    
    e.BoadersIndex = xlDiagonalUp
    e.LineStyle = xlContinuous
    e.Weight = xlThin
    
    e.Run
    
    Set e = Nothing

End Sub
'--------------------------------------------------------------
' 右上がり罫線
'--------------------------------------------------------------
Sub execSelectionAllDiagonalDown()

'    Dim obj As SelectionAllDiagonalDown
'
'    Set obj = New SelectionAllDiagonalDown
'
'    obj.Run
'
'    Set obj = Nothing

    Dim e As SelectionFormatBoader
    
    Set e = New SelectionFormatBoader
    
    e.BoadersIndex = xlDiagonalDown
    e.LineStyle = xlContinuous
    e.Weight = xlThin
    
    e.Run
    
    Set e = Nothing

End Sub
'--------------------------------------------------------------
' 左１桁削除
'--------------------------------------------------------------
Sub execSelectionDelete1Char()

    Dim obj As SelectionDelete1Char
    
    Set obj = New SelectionDelete1Char
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 右１桁削除
'--------------------------------------------------------------
Sub execSelectionDelete1RightChar()

    Dim obj As SelectionDelete1RightChar
    
    Set obj = New SelectionDelete1RightChar
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' ふり仮名クリア
'--------------------------------------------------------------
Sub execSelectionClearPhonetic()

    Dim obj As SelectionClearPhonetic
    
    Set obj = New SelectionClearPhonetic
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' 段落番号削除
'--------------------------------------------------------------
Sub execSelectionDelSectionNo()

    Dim obj As SelectionDelSectionNo
    
    Set obj = New SelectionDelSectionNo
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 段落番号追加
'--------------------------------------------------------------
Sub execSelectionAllAddSectionNo()

    Dim obj As SelectionAddSectionNo
    
    Set obj = New SelectionAddSectionNo
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 段落番号コピー
'--------------------------------------------------------------
Sub execSelectionCopySectionNo()

    Dim obj As SelectionCopySectionNo
    
    Set obj = New SelectionCopySectionNo
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 箇条書き削除
'--------------------------------------------------------------
Sub execSelectionDelItemNo()

    Dim obj As SelectionDelItemNo
    
    Set obj = New SelectionDelItemNo
    
    obj.Run
    
    Set obj = Nothing

End Sub '--------------------------------------------------------------
' 「・」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemPoint()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemPoint"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「●」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemCircleB()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemCircleB"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「○」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemCircleW()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemCircleW"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「◆」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemDiaB()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemDiaB"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「◇」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemDiaW()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemDiaW"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「▼」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemRevTriB()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemRevTriB"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「▽」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemRevTriW()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemRevTriW"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「■」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemSquareB()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemSquareB"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「□」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemSquareW()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemSquareW"
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' 「1)」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNum1()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemNum1"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「(a)」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumA2()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemNumA2"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「①」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumC()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemNumC"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「a)」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumA()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemNumA"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「例１)」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumExp()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemNumExp"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「1.」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumPoint2()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemNumPoint2"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「※」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumImp()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemImp"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「◎」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumDouble()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemDouble"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「★」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumStarB()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemStarB"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「☆」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumStarW()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemStarW"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「〆」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumSime()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemSime"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 「⇒」追加
'--------------------------------------------------------------
Sub execSelectionAllAddItemNumDblR()

    Dim obj As SelectionAddItemNo
    
    Set obj = New SelectionAddItemNo
    
    obj.ItemName = "itemDblR"
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' おなじ値をマージ
'--------------------------------------------------------------
Sub execSelectionMerge()

    Dim obj As SelectionMerge
    
    Set obj = New SelectionMerge
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' おなじ値をマージ
'--------------------------------------------------------------
Sub execSelectionMergeLine()

    Dim obj As SelectionMergeLine
    
    Set obj = New SelectionMergeLine
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' マージの逆
'--------------------------------------------------------------
Sub execSelectionNoMerge()

    Dim obj As SelectionNoMerge
    
    Set obj = New SelectionNoMerge
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' クリップボードファイルコピー
'--------------------------------------------------------------
Sub execSelectionSetClipboardCopy()

    Dim obj As SelectionSetClipboardCopy
    
    Set obj = New SelectionSetClipboardCopy
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' ある文字から左側を削除
'--------------------------------------------------------------
Sub execSelectionRemoveLeftToChar()

    Dim obj As SelectionRemoveLeftToChar
    
    Set obj = New SelectionRemoveLeftToChar
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' ある文字から右側を削除
'--------------------------------------------------------------
Sub execSelectionRemoveRightToChar()

    Dim obj As SelectionRemoveRightToChar
    
    Set obj = New SelectionRemoveRightToChar
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 絶対参照変換
'--------------------------------------------------------------
Sub execSelectionToAbsolute()

    Dim obj As SelectionToAbsolute
    
    Set obj = New SelectionToAbsolute
    
    obj.RefType = xlAbsolute
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 相対参照変換
'--------------------------------------------------------------
Sub execSelectionToRelative()

    Dim obj As SelectionToAbsolute
    
    Set obj = New SelectionToAbsolute
    
    obj.RefType = xlRelative
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 絶対参照変換(列)
'--------------------------------------------------------------
Sub execSelectionToRelRowAbsColumn()

    Dim obj As SelectionToAbsolute
    
    Set obj = New SelectionToAbsolute
    
    obj.RefType = xlRelRowAbsColumn
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 絶対参照変換(行)
'--------------------------------------------------------------
Sub execSelectionToAbsRowRelColumn()

    Dim obj As SelectionToAbsolute
    
    Set obj = New SelectionToAbsolute
    
    obj.RefType = xlAbsRowRelColumn
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignTopLeft()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlTop
    obj.HorizontalAlignment = xlLeft
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignTopCenter()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlTop
    obj.HorizontalAlignment = xlCenter
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignTopRight()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlTop
    obj.HorizontalAlignment = xlRight
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignCenterLeft()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlCenter
    obj.HorizontalAlignment = xlLeft
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignCenterCenter()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlCenter
    obj.HorizontalAlignment = xlCenter
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignCenterRight()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlCenter
    obj.HorizontalAlignment = xlRight
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignBottomLeft()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlBottom
    obj.HorizontalAlignment = xlLeft
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignBottomCenter()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlBottom
    obj.HorizontalAlignment = xlCenter
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 文字詰め
'--------------------------------------------------------------
Sub execSelectionAlignBottomRight()

    Dim obj As SelectionFormatAlign
    
    Set obj = New SelectionFormatAlign
    
    obj.VerticalAlignment = xlBottom
    obj.HorizontalAlignment = xlRight
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' インデントクリア
'--------------------------------------------------------------
Sub execSelectionSectionIndentClear()

    Dim obj As SelectionAllIndentClear
    
    Set obj = New SelectionAllIndentClear
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 次番号
'--------------------------------------------------------------
Sub getNextNumber()
    
    Dim obj As SelectionAllNextNo
    
    Set obj = New SelectionAllNextNo
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 次番号(左)
'--------------------------------------------------------------
Sub getNextNumberLeft()
    
    Dim obj As SelectionAllLeftNo
    
    Set obj = New SelectionAllLeftNo
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' Relaxフィル(＋１)
'--------------------------------------------------------------
Sub execSelectionAllPlus()
    
    Dim obj As SelectionAllPlus
    
    Set obj = New SelectionAllPlus
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' Relaxフィル(－１)
'--------------------------------------------------------------
Sub execSelectionAllMinus()
    
    Dim obj As SelectionAllMinus
    
    Set obj = New SelectionAllMinus
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' マイナンバー（個人）チェックデジット
'--------------------------------------------------------------
Sub execSelectionCheckMyNumber()
    
    Dim obj As SelectionCheckMyNumber
    
    Set obj = New SelectionCheckMyNumber
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' マイナンバー（企業）チェックデジット
'--------------------------------------------------------------
Sub execSelectionCheckCorpNumber()
    
    Dim obj As SelectionCheckCorpNumber
    
    Set obj = New SelectionCheckCorpNumber
    
    obj.Run
    
    Set obj = Nothing
    
End Sub
'--------------------------------------------------------------
' 2003互換色(文字色)
'--------------------------------------------------------------
Sub execSelectionFormatFontColor()
    
    Dim obj As SelectionFormatFontColor
    
    Set obj = New SelectionFormatFontColor
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 2003互換色(背景色)
'--------------------------------------------------------------
Sub execSelectionFormatBackColor()
    
    Dim obj As SelectionFormatBackColor
    
    Set obj = New SelectionFormatBackColor
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 2003互換色(枠線色)
'--------------------------------------------------------------
Sub execSelectionFormatLineColor()
    
    Dim obj As SelectionFormatLineColor
    
    Set obj = New SelectionFormatLineColor
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' チェックリスト罫線
'--------------------------------------------------------------
Sub execSelectionFormatCheckList()
    
    Dim obj As SelectionFormatCheckList
    
    Set obj = New SelectionFormatCheckList
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' チェックリスト作成(○)
'--------------------------------------------------------------
Sub execSelectionFormatCL()
    
    Dim obj As SelectionFormatCL
    
    Set obj = New SelectionFormatCL
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' チェックリスト作成(YN)
'--------------------------------------------------------------
Sub execSelectionFormatDec()
    
    Dim obj As SelectionFormatDec
    
    Set obj = New SelectionFormatDec
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 書式（標準）
'--------------------------------------------------------------
Sub execSelectionFormatStd()
    
    Dim obj As SelectionFormatStd
    
    Set obj = New SelectionFormatStd
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 書式（文字列）
'--------------------------------------------------------------
Sub execSelectionFormatStr()
    
    Dim obj As SelectionFormatStr
    
    Set obj = New SelectionFormatStr
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 書式（数値エラー解除）
'--------------------------------------------------------------
Sub execSelectionNumberAsText()
    
    Dim obj As SelectionNumberAsText
    
    Set obj = New SelectionNumberAsText
    
    obj.Run
    
    Set obj = Nothing

End Sub

'--------------------------------------------------------------
' 月＋１
'--------------------------------------------------------------
Sub execSelectioDateMonthAdd()
    
    Dim obj As SelectionDateAdd
    
    Set obj = New SelectionDateAdd
    
    obj.DateType = "m"
    obj.DateValue = 1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 月－１
'--------------------------------------------------------------
Sub execSelectioDateMonthMinus()
    
    Dim obj As SelectionDateAdd
    
    Set obj = New SelectionDateAdd
    
    obj.DateType = "m"
    obj.DateValue = -1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 年＋１
'--------------------------------------------------------------
Sub execSelectioDateYearAdd()
    
    Dim obj As SelectionDateAdd
    
    Set obj = New SelectionDateAdd
    
    obj.DateType = "yyyy"
    obj.DateValue = 1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 年－１
'--------------------------------------------------------------
Sub execSelectioDateYearMinus()
    
    Dim obj As SelectionDateAdd
    
    Set obj = New SelectionDateAdd
    
    obj.DateType = "yyyy"
    obj.DateValue = -1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 日＋１
'--------------------------------------------------------------
Sub execSelectioDateDayAdd()
    
    Dim obj As SelectionDateAdd
    
    Set obj = New SelectionDateAdd
    
    obj.DateType = "d"
    obj.DateValue = 1
    
    obj.Run
    
    Set obj = Nothing

End Sub
'--------------------------------------------------------------
' 日－１
'--------------------------------------------------------------
Sub execSelectioDateDayMinus()
    
    Dim obj As SelectionDateAdd
    
    Set obj = New SelectionDateAdd
    
    obj.DateType = "d"
    obj.DateValue = -1
    
    obj.Run
    
    Set obj = Nothing

End Sub
