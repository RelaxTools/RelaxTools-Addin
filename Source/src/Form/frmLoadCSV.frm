VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoadCSV 
   Caption         =   "CSVファイル取込"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   OleObjectBlob   =   "frmLoadCSV.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmLoadCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private mdblLineWidth As Double
Private mblnCancel As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFolder_Click()

    Dim strFile As String


    strFile = Application.GetOpenFilename("カンマ区切りファイル(*.csv;*.txt),(*.csv;*.txt)", , "ＣＳＶファイル読込", , False)
    If strFile = "False" Then
        'ファイル名が指定されなかった場合
        Exit Sub
    End If
    
    txtFolder.Text = strFile
    
End Sub

Private Sub cmdRun_Click()

    Dim fp As Integer
    Dim WS As Worksheet
    Dim strFile As String
    Dim strBuf As String
    Dim bytBuf() As Byte
    Dim varRow As Variant
    
    
    Dim j As Long
    Dim k As Long
    Dim arPaste() As Variant
    Dim lngsize As Long
    Dim lngRead As Long
    Dim lngMax As Long
    
    Dim mm As MacroManager
        
    Dim r As Range
    
    strFile = txtFolder.Text
    
    'ファイルの存在チェック
    If rlxIsFileExists(strFile) Then
    Else
        MsgBox "ファイルが存在しません。", vbExclamation, C_TITLE
        Exit Sub
    End If
    
    Dim lngRow As Long
    Dim lngCol As Long
    
    fp = FreeFile()
    Open strFile For Binary As fp
    
    lngsize = LOF(fp)
    If lngsize <> 0 Then
        ReDim bytBuf(0 To LOF(fp) - 1)
        Get fp, , bytBuf
    End If
    
    Close fp
    
    If lngsize = 0 Then
        Exit Sub
    End If
    
    Set mm = New MacroManager
    Set mm.Form = Me
    mm.Disable
    mm.DispGuidance "ＣＳＶファイルの行数をカウントしています..."
    
    Set WS = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
    
    If chkUTF8.Value Then
        'UTF8からUNICODE
        Dim utf8 As UTF8Encoding
        Set utf8 = New UTF8Encoding
        strBuf = utf8.GetString(bytBuf())
        Set utf8 = Nothing
    Else
        strBuf = StrConv(bytBuf, vbUnicode)
    End If
    
    lngRow = 1
    lngCol = 1
    
    Dim lngPos As Long
    Dim i As Long
    Dim strLine() As String
    
    lngPos = InStr(strBuf, vbCrLf)
    If lngPos <> 0 Then
        strLine = Split(strBuf, vbCrLf)
    Else
        lngPos = InStr(strBuf, vbLf)
        If lngPos <> 0 Then
            strLine = Split(strBuf, vbLf)
        Else
            strLine = Split(strBuf, vbCr)
        End If
    End If
    
'    'CRLF or CR の場合 LF に変換
'    strBuf = Replace(strBuf, vbCrLf, vbLf)
'    strBuf = Replace(strBuf, vbCr, vbLf)
'    strLine = Split(strBuf, vbLf)
    
    'カンマ区切りで分割を行う（ダブルコーテーション内カンマ対応）
    varRow = rlxCsvPart(strLine(1))
    
    '項目数の分、列の選択をし、文字列形式にする。
    Set r = Range(WS.Columns(lngCol), WS.Columns(lngCol + UBound(varRow) - 1))
    r.NumberFormatLocal = "@"

    Const BASE_LINE As Long = 20000
    
    lngsize = UBound(strLine) + 1
    
    lngMax = UBound(strLine) + 1
    mm.StartGauge lngMax
    
    lngRow = 1
    
    i = 0
    Do While lngsize > 0
    
        If lngsize < BASE_LINE Then
            lngRead = lngsize
        Else
            lngRead = BASE_LINE
        End If
        
        ReDim arPaste(1 To lngRead, LBound(varRow) To UBound(varRow))
        
        For k = 1 To lngRead
            'カンマ区切りで分割を行う（ダブルコーテーション内カンマ対応）
            varRow = rlxCsvPart(strLine(i))
            For j = LBound(varRow) To UBound(varRow)
                arPaste(k, j) = varRow(j)
            Next
            i = i + 1
            'ゲージの表示
            mm.DisplayGauge i
            If mblnCancel Then
                Exit Do
            End If
        Next
           
        Range(WS.Cells(lngRow, 1), WS.Cells(lngRow + UBound(arPaste, 1) - 1, UBound(arPaste, 2))).Value = arPaste
       lngsize = lngsize - lngRead
       lngRow = lngRow + lngRead
       
    Loop
    
    
'    WS.Name = rlxGetFullpathFromFileName(strFile)
       
    'すべて貼り付けたら列間隔を調整
    If r Is Nothing Then
    Else
        r.AutoFit
        Set r = Nothing
    End If
    
    Set mm = Nothing
    Unload Me
    MsgBox "処理が完了しました。", vbInformation, C_TITLE

End Sub

Private Sub UserForm_Initialize()
    lblGauge.visible = False
    mblnCancel = False
End Sub

Private Sub UserForm_Terminate()
    mblnCancel = True
End Sub
