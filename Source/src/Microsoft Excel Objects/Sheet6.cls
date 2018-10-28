VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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

'定義ファイル書式
Const C_NO As Long = 1
Const C_COLUMN As Long = 2
Const C_POS_COL As Long = 3
Const C_POS_LEN As Long = 4

'定義スタート行
Const C_DEF_START_ROW = 7

'「読込を行うテキストファイル」の場所
Const C_FILE_ROW As Long = 3
Const C_FILE_COL As Long = 3

'「１件のデータを識別する文字列」の場所
Const C_START_ROW As Long = 4
Const C_START_COL As Long = 3

'作成シート
Const C_CRE_HEADER_ROW As Long = 1
Const C_CRE_START_ROW As Long = 2

Private ColLine As Collection
        

Private Sub cmdRun_Click()

    Dim strFile As String
    Dim strNewSheet As String
    Dim fp As Integer

    Dim defCount As Long
    Dim creRow As Long
    Dim creCol As Long

    Dim strBuf As String
    Dim strValue As String

    Dim strColumn As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngLen As Long

    Dim newSheet As Worksheet
    Dim defSheet As Worksheet

    Dim strStart As String
    Dim i As Long

    Set defSheet = ActiveSheet

    strFile = defSheet.Cells(C_FILE_ROW, C_FILE_COL).Value
    If strFile = "" Then
        'ファイル名が指定されなかった場合
        Exit Sub
    End If

    strNewSheet = getFileName(strFile)

    Set newSheet = Worksheets.Add


    fp = FreeFile()

    strStart = defSheet.Cells(C_START_ROW, C_START_COL).Value


    Open strFile For Input As fp

    '定義シートのカウンタ
    defCount = C_DEF_START_ROW
    creCol = 1

    '定義シートを読みヘッダを作成する。
    Do Until defSheet.Cells(defCount, C_NO).Value = ""

        strColumn = defSheet.Cells(defCount, C_COLUMN).Value
        newSheet.Cells(C_CRE_HEADER_ROW, creCol).Value = strColumn

        defCount = defCount + 1
        creCol = creCol + 1
    Loop

    '結果シートのカウンタ
    creRow = C_CRE_START_ROW

    Do Until Eof(fp)

        '--------------------------------------------------
        ' １レコード分のデータをコレクションに設定する。
        '--------------------------------------------------
        Set ColLine = New Collection
        Dim blnStart As Boolean
        
        Do Until Eof(fp)
            
            Line Input #fp, strBuf
            If InStr(strBuf, strStart) > 0 Then
                If blnStart Then
                    Exit Do
                Else
                    blnStart = True
                End If
            End If
        
            ColLine.Add strBuf
        Loop
        
        '--------------------------------------------------
        ' １レコード分のデータより定義シートの内容を検索する。
        '--------------------------------------------------
        '結果シートのカウンタ
        creCol = 1

        For i = 1 To ColLine.Count
        
            strBuf = ColLine.Item(i)
            '定義シートのカウンタ
            defCount = C_DEF_START_ROW

            '定義シートを読む
            Do Until defSheet.Cells(defCount, C_NO).Value = ""
    
                '項目名
                strColumn = defSheet.Cells(defCount, C_COLUMN).Value
                '桁
                lngCol = defSheet.Cells(defCount, C_POS_COL).Value
                '長さ
                lngLen = defSheet.Cells(defCount, C_POS_LEN).Value
    
                'キーが見つかった場合
                If InStr(strBuf, strColumn) > 0 Then
                
                    '文字列長に０が指定されていたら、行末まで取得
                    If lngLen = 0 Then
                        strValue = Trim(Mid(strBuf, lngCol))
                    Else
                        strValue = Trim(Mid(strBuf, lngCol, lngLen))
                    End If
        
                    'セルに値を設定
                    newSheet.Cells(creRow, creCol).Value = strValue
                
                    '作成シートの列＋１
                    creCol = creCol + 1
    
                End If
    
                '次の定義に移動
                defCount = defCount + 1
            Loop
        Next

        '結果シートを次の行に移動する。
        creRow = creRow + 1
        Set ColLine = Nothing

    Loop

pass:
    Close fp

    Dim r As Range
    Set r = newSheet.Range(newSheet.Columns(1), newSheet.Columns(creCol))
    If r Is Nothing Then
    Else
        r.AutoFit
    End If
    Set r = Nothing

End Sub
Private Function getFileName(ByVal strFile As String) As String

    Dim Length As Long
    Dim strBuf As String
    Dim cnt As Long
    
    Length = Len(strFile)

    For cnt = Length To 1 Step -1
    
        If Mid(strFile, cnt, 1) = "\" Then
            getFileName = Mid(strFile, cnt + 1)
            Exit Function
        End If
    Next

End Function


