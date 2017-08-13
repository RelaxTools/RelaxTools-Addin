Attribute VB_Name = "basBinaryView"
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
'　バイナリービュー
'--------------------------------------------------------------
Private Const C_START_ROW As Long = 2
Private Const C_NO As Long = 1
Private Const C_STR As Long = 18
Sub BinaryViewer()

    Dim strFile As String
    Dim intIn As Integer
    Dim lngsize As Long
    Dim i As Long
    Dim bytBuf() As Byte
    
    Dim lngRead As Long
    
    Const key As Byte = &H44
    Const C_BUFFER_SIZE = 1048576 '1MB

    On Error GoTo ErrHandle
    
    strFile = Application.GetOpenFilename(, , "バイナリービュー", , False)
    If strFile = "False" Then
        'ファイル名が指定されなかった場合
        Exit Sub
    End If
    
    'ファイルの存在チェック
    If rlxIsFileExists(strFile) Then
    Else
        MsgBox "ファイルが存在しません。", vbExclamation, C_TITLE
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim WS As Worksheet
    
    ThisWorkbook.Worksheets("BinaryView").Copy
    Set WS = Application.Workbooks(Application.Workbooks.count).Worksheets(1)
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngAddress As Long
    
    lngRow = 1
    lngCol = C_NO
    Dim varBuf() As String

    intIn = FreeFile()
    Open strFile For Binary As intIn
    
    lngsize = LOF(intIn)
    
    If lngsize < C_BUFFER_SIZE Then
        lngRead = lngsize
    Else
        lngRead = C_BUFFER_SIZE
    End If
    
    '最大で1MBのメモリを確保。
    ReDim bytBuf(0 To lngRead - 1)

    '確保したバイト数分読み込み
    Get intIn, , bytBuf
    
    Close intIn

    Dim strChar As String
    Dim bytChar() As Byte
    ReDim bytChar(0 To 31)
    Dim lngPos As Long
    Dim lngLast As Long
    
    lngLast = (lngRead / 16) + 2
    
    ReDim varBuf(1 To lngLast, C_NO To C_STR)
    
    For i = 0 To lngRead - 1
    
        If lngCol = C_NO Then
            varBuf(lngRow, lngCol) = FixHex(lngAddress, 8)
            lngPos = i
            lngCol = lngCol + 1
        End If
        
        varBuf(lngRow, lngCol) = FixHex(bytBuf(i), 2)
        
        lngCol = lngCol + 1
        If lngCol >= C_STR Then
            bytCopy bytBuf, bytChar, lngPos, 32
            varBuf(lngRow, C_STR) = AscLeft(ReplaceStr(StrConv(bytChar, vbUnicode)), 16)
            lngPos = 0
            lngCol = 1
            lngRow = lngRow + 1
        End If
        
        lngAddress = lngAddress + 1
        
    Next
    
    Dim r As Range
    Dim s As Range
    Set r = WS.Range(WS.Cells(C_START_ROW, C_NO), WS.Cells(lngRow + C_START_ROW, C_STR))
    
    r.Value = varBuf
    
    For Each s In r
        s.Errors.Item(xlNumberAsText).Ignore = True
    Next
    
    For Each s In WS.Range(WS.Cells(1, C_NO), WS.Cells(1, C_STR))
        s.Errors.Item(xlNumberAsText).Ignore = True
    Next
    
    Application.ScreenUpdating = True

'    MsgBox "読み込みが完了しました。", vbInformation, C_TITLE
    
    Exit Sub
ErrHandle:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。", vbOKOnly, C_TITLE
End Sub
Function ReplaceStr(ByVal strBuf As String) As String

    ReplaceStr = Replace(Replace(strBuf, vbLf, "."), vbCr, ".")

End Function
Sub bytCopy(ByRef bytSource() As Byte, ByRef bytDest() As Byte, ByVal lngPos As Long, ByVal lngsize As Long)

    Dim i As Long
    Dim j As Long
    
    i = lngPos
    j = 0
    
    ReDim bytDest(0 To lngsize - 1)
    
    Do Until lngPos + lngsize <= i
        If UBound(bytSource) < i Then
            Exit Do
        End If
        bytDest(j) = bytSource(i)
        i = i + 1
        j = j + 1
    Loop


End Sub
Function FixHex(ByVal lngAddress As Long, ByVal lngLen As Long) As String
    FixHex = Right$(String$(lngLen, "0") & Hex(lngAddress), lngLen)
End Function
'----------------------------------------------------------------------------------
'　文字列の左端から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
'----------------------------------------------------------------------------------
Private Function AscLeft(ByVal var As Variant, ByVal lngsize As Long) As String

    Dim lngLen As Long
    Dim i As Long
    
    Dim strChr As String
    Dim strResult As String
    
    lngLen = Len(var)
    strResult = ""

    For i = 1 To lngLen
    
        strChr = Mid(var, i, 1)
        strResult = strResult & strChr
        If rlxAscLen(strResult) > lngsize Then
            Exit For
        End If
    
    Next

    AscLeft = strResult

End Function
Sub a()

    Dim i As Long
    Dim j As Long

    For i = 1 To 30000

        For j = 1 To 18
            ActiveSheet.Cells(i, j).Errors.Item(xlNumberAsText).Ignore = True
        Next
    Next
    MsgBox "OK"
End Sub
