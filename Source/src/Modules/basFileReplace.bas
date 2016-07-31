Attribute VB_Name = "basFileReplace"
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
'　ファイル置換
'--------------------------------------------------------------
Sub replaceFiles()

    Dim strFolder As String
    
    Dim strKey As String
    
    Dim strBackup As String
    
    Dim strTmpl As String
    Dim lngCnt As Long
    Dim lngFileCount As Long

    
    Const C_COL_NUM As Long = 1
    Const C_COL_SEARCH As Long = 2
    Const C_COL_REPLACE As Long = 3
    Const C_COL_COMPARE As Long = 4
    Const C_COL_RESULT As Long = 5
    
    Const C_ROW_VERSION As Long = 1
    Const C_ROW_HEAD As Long = 3
    Const C_ROW_DETAIL As Long = 4

    Dim rp As ReplaceParamDTO
    Dim colParam As New Collection
    
    Dim colResult As New Collection
   
    Dim blnStatusBar As Boolean ''進捗バー状態を記憶
   
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Worksheets("ReplaceFormat")
    
    '自ブックのフォーマットとアクティブなシートのフォーマットを比較
    If WS.Cells(C_ROW_VERSION, C_COL_NUM).Value <> _
        ActiveSheet.Cells(C_ROW_VERSION, C_COL_NUM).Value Then
        If MsgBox("定義シートが異なる可能性がありますが、続行しますか？", vbOKCancel, "ファイル内部置換") = vbCancel Then
            Exit Sub
        End If
    End If
    
    lngCnt = 0

    ''設定ファイルの読み込み
    Do Until Cells(lngCnt + C_ROW_DETAIL, C_COL_NUM).Value = ""
        
        strKey = Cells(lngCnt + C_ROW_DETAIL, C_COL_SEARCH).Value
        If strKey <> "" Then
            Set rp = New ReplaceParamDTO
            
            rp.SearchString = strKey
            rp.ReplaceString = Cells(lngCnt + C_ROW_DETAIL, C_COL_REPLACE).Value
            rp.CompareMode = Cells(lngCnt + C_ROW_DETAIL, C_COL_COMPARE).Value

            colParam.Add rp
        
            Set rp = Nothing
        End If
   
        lngCnt = lngCnt + 1
    Loop
        
        
    'フォルダ名取得
    strFolder = rlxSelectFolder()
    If strFolder = "" Then
        'キャンセル
        Exit Sub
    End If
    
    '''ステータスバーの状態を記憶
    blnStatusBar = Application.DisplayStatusBar
    ''ステータスバーを表示
    Application.DisplayStatusBar = True
    
    Call reDir(strFolder, colParam, colResult)

    '''==================================================
    '''ステータスバーの表示内容をExcelの既定値に戻す
    Application.StatusBar = False
    '''ステータスバーをマクロの実行前の状態に戻す
    Application.DisplayStatusBar = blnStatusBar
    '''==================================================

    MsgBox colResult.count & "ファイル置換しました。", vbInformation, C_TITLE
    
End Sub
'--------------------------------------------------------------
'　フォルダの再帰検索
'--------------------------------------------------------------
Private Sub reDir(ByVal strDir As String, colParam As Collection, colResult As Collection)

    Dim strFile As String
    Dim fc As New Collection
    Dim FS As FileParamDTO
    Dim strParent As String
    
    strParent = rlxAddFileSeparator(strDir)

    ''ファイルの読み込み処理
    strFile = Dir(strParent & "*.*", vbNormal + vbDirectory)

    Do Until strFile = ""
    
        Select Case strFile
            Case ".", ".."
            Case Else
                Set FS = New FileParamDTO
                
                FS.FileName = strFile
                FS.Directory = strDir
                FS.Attrib = getAttr(rlxAddFileSeparator(strDir) & strFile)
                
                fc.Add FS
        End Select
    
        ''次のファイルを取得
        strFile = Dir()
    Loop

    For Each FS In fc
    
        If (FS.Attrib And vbDirectory) <> 0 Then
            ''再帰呼び出し
            Call reDir(rlxAddFileSeparator(FS.Directory) & FS.FileName, colParam, colResult)
        Else
            Call repFiles(FS, colParam, colResult)
        End If
    Next

End Sub
'--------------------------------------------------------------
'　ファイルの再帰検索
'--------------------------------------------------------------
Private Sub repFiles(FS As FileParamDTO, colParam As Collection, colResult As Collection)

    Dim strSourceFile As String
    
    Dim strBuf As String
    Dim strBody As String
    Dim strWrite As String
    Dim bytBuf() As Byte
    
    
    Dim intfp As Integer
    Dim lngMaxCol As Long
    Dim lngSize As Long
    
    Dim lngSearchChar As Long
    Dim lngSearchCount As Long
    Dim lngFileCount As Long
    
    Dim lngFind As Long
    
    Dim rp As ReplaceParamDTO
    Dim rr As ReplaceResultDTO
    
    
    ''''==================================================
    ''''ステータスバーにメッセージを表示
    Application.StatusBar = FS.FileName & "を処理中です"
    ''''==================================================

    'ファイル名の作成
    strSourceFile = rlxAddFileSeparator(FS.Directory) & FS.FileName

    'ファイルを全部読む。
    intfp = FreeFile()
    Open strSourceFile For Binary As intfp
    lngSize = LOF(intfp)
    
    ReDim bytBuf(0 To lngSize - 1)

    Get intfp, , bytBuf

    Close intfp

    ''Unicodeに変換
    strBody = StrConv(bytBuf, vbUnicode)
    lngSearchCount = 0

    For Each rp In colParam
    
        lngSearchChar = 0
        
        '発見された文字列の数を取得
        lngFind = InStr(1, strBody, rp.SearchString, rp.CompareMode)
        Do Until lngFind = 0
            lngSearchChar = lngSearchChar + 1
            lngFind = InStr(lngFind + 1, strBody, rp.SearchString, rp.CompareMode)
        Loop
        
        ''１件でも検索結果があればファイル内を置換
        If lngSearchChar > 0 Then
            strBody = Replace(strBody, rp.SearchString, rp.ReplaceString, 1, -1, rp.CompareMode)
            lngSearchCount = lngSearchCount + lngSearchChar
            
            Set rr = New ReplaceResultDTO
            
            rr.FileName = FS.FileName
            rr.SearchString = rp.SearchString
            rr.ReplaceString = rp.ReplaceString
            rr.ReplaceStrCount = lngSearchCount
            
            colResult.Add rr
        End If

    Next

    If lngSearchCount > 0 Then
    
   
        'ファイルが存在する場合、一度クリアする。
        intfp = FreeFile()
        Open strSourceFile For Output As intfp
        Close intfp
        
        ''１件でも置換があれば出力
        intfp = FreeFile()
        Open strSourceFile For Binary As intfp
        Put intfp, , strBody
        Close intfp
        
        '修正されたファイルの数をカウント
        lngFileCount = lngFileCount + 1
        
    End If

End Sub


