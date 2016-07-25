Attribute VB_Name = "basKantan"
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
'　プレビュー用画像作成
'--------------------------------------------------------------
Public Function editKantan(ByRef s As KantanLine, ByVal lngFormat As Long) As StdPicture

    Dim strMiddle As String
    Dim WS As Worksheet
    
    Set editKantan = Nothing
    
    Set WS = ThisWorkbook.Worksheets("table")

    Dim r As Range
    Dim m As Long
    Dim i As Long
    
    Set r = WS.Range("c4:e8")
    
    '塗りつぶしなし
    r.Interior.Pattern = xlNone
    
    '中線
    r.Borders(xlInsideHorizontal).Weight = getBorderWeight(s.HInsideLine)
    r.Borders(xlInsideHorizontal).LineStyle = getLineStyle(s.HInsideLine)
    
    r.Borders(xlInsideVertical).Weight = getBorderWeight(s.VInsideLine)
    r.Borders(xlInsideVertical).LineStyle = getLineStyle(s.VInsideLine)
        
    'ヘッダ線
    For i = 1 To s.HHeadLineCount
        If i > r.Rows.count Then
            Exit For
        End If
        If s.EnableHRepeat Or i = s.HHeadLineCount Then
            r.Rows(i).Borders(xlEdgeBottom).Weight = getBorderWeight(s.HHeadBorderLine)
            r.Rows(i).Borders(xlEdgeBottom).LineStyle = getLineStyle(s.HHeadBorderLine)
        Else
            r.Rows(i).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
        End If
        r.Rows(i).Borders(xlInsideVertical).Weight = getBorderWeight(s.VHeadBorderLine)
        r.Rows(i).Borders(xlInsideVertical).LineStyle = getLineStyle(s.VHeadBorderLine)
        r.Rows(i).Interior.Color = s.HeadColor
    Next
    
    If s.EnableEvenColor Then
        For i = s.HHeadLineCount + 2 To r.Rows.count Step 2
            r.Rows(i).Interior.Color = s.EvenColor
        Next
        
    End If
    
    For i = 1 To s.VHeadLineCount
        If i > r.Columns.count Then
            Exit For
        End If
        If s.EnableVRepeat Or i = s.VHeadLineCount Then
            r.Columns(i).Borders(xlEdgeRight).Weight = getBorderWeight(s.VHeadBorderLine)
            r.Columns(i).Borders(xlEdgeRight).LineStyle = getLineStyle(s.VHeadBorderLine)
        Else
            r.Columns(i).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
        End If
        r.Columns(i).Borders(xlInsideHorizontal).Weight = getBorderWeight(s.HHeadBorderLine)
        r.Columns(i).Borders(xlInsideHorizontal).LineStyle = getLineStyle(s.HHeadBorderLine)
        r.Columns(i).Interior.Color = s.HeadColor
    Next
    
    '外周
    r.Borders(xlEdgeTop).Weight = getBorderWeight(s.OutSideLine)
    r.Borders(xlEdgeTop).LineStyle = getLineStyle(s.OutSideLine)
    
    r.Borders(xlEdgeLeft).Weight = getBorderWeight(s.OutSideLine)
    r.Borders(xlEdgeLeft).LineStyle = getLineStyle(s.OutSideLine)
    
    r.Borders(xlEdgeRight).Weight = getBorderWeight(s.OutSideLine)
    r.Borders(xlEdgeRight).LineStyle = getLineStyle(s.OutSideLine)
    
    r.Borders(xlEdgeBottom).Weight = getBorderWeight(s.OutSideLine)
    r.Borders(xlEdgeBottom).LineStyle = getLineStyle(s.OutSideLine)
    
    
    If s.EnableHogan Then
        r(2, 1).value = "方"
        r(3, 2).value = "眼"
        r(4, 3).value = "紙"
    Else
        r(2, 1).value = ""
        r(3, 2).value = ""
        r(4, 3).value = ""
    End If
    
    Set editKantan = CreatePictureFromClipboard(WS.Range("b3:f9"))
    
    Set WS = Nothing
    
End Function

'--------------------------------------------------------------
'　bzイメージファイル作成
'--------------------------------------------------------------
Function getImageKantan(ByVal Index As Long) As StdPicture

    '設定情報取得
    Dim col As Collection
    Dim k As KantanLine

    Set getImageKantan = Nothing

    Set col = getPropertyKantan()

    Set k = col(Index)

    Set getImageKantan = editKantan(k, xlBitmap)

    Set k = Nothing

End Function
Public Sub KantanPaste()

    Dim lngNo As Long

    On Error GoTo e

    lngNo = GetSetting(C_TITLE, "KantanDx", "kantanNo", 1)
    
    Application.ScreenUpdating = False
    
    Call kantanPaste2(lngNo)
    
    Application.ScreenUpdating = True
    
    Exit Sub
e:
    Logger.LogFatal "かんたん表DXの貼り付けでエラー"
End Sub

'--------------------------------------------------------------
'　データ印貼り付け
'--------------------------------------------------------------
Sub kantanPaste2(ByVal Index As Long)

    '設定情報取得
    Dim col As Collection
    Dim s As KantanLine
    Dim r As Range
    Dim i As Long
    Dim j As Long
    Dim z As Long
    
    Dim lngHLine As Long
    Dim lngVLine As Long

    On Error Resume Next

    If rlxCheckSelectRange() = False Then
        MsgBox "選択範囲が見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If

    Set col = getPropertyKantan()

    Select Case True
        Case col Is Nothing
            Exit Sub
        Case col.count = 0
            Exit Sub
        Case Else
    End Select

    Set s = col(Index)
    
    ThisWorkbook.Worksheets("Undo").Cells.Clear
    Set mUndo.sourceRange = Selection
    Set mUndo.destRange = ThisWorkbook.Worksheets("Undo").Range(Selection.Address)
    
    Dim rr As Range
    For Each rr In mUndo.sourceRange.Areas
        rr.Copy mUndo.destRange.Worksheet.Range(rr.Address)
    Next
    
    For Each r In Selection.Areas

        'ヘッダの境界行（方眼紙の際の列判定）
        If s.HHeadLineCount < r.Rows.count Then
            lngHLine = s.HHeadLineCount
        Else
            lngHLine = r.Rows.count
        End If
        
        If s.VHeadLineCount < r.Columns.count Then
            lngVLine = s.VHeadLineCount
        Else
            lngVLine = r.Columns.count
        End If
        
        If s.AuthoHogan Then
            z = lngHLine
        Else
            z = s.HoganJudgeLineCount
        End If
        
        '塗りつぶしなし
        r.Interior.Pattern = xlNone
        
        '中線
        r.Borders(xlInsideHorizontal).Weight = getBorderWeight(s.HInsideLine)
        r.Borders(xlInsideHorizontal).LineStyle = getLineStyle(s.HInsideLine)
        
        If s.EnableHogan Then
            For j = 2 To r.Columns.count
                If r(z, j).value <> "" Then
                    r.Columns(j).Borders(xlEdgeLeft).Weight = getBorderWeight(s.VHeadBorderLine)
                    r.Columns(j).Borders(xlEdgeLeft).LineStyle = getLineStyle(s.VHeadBorderLine)
                Else
                    r.Columns(j).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                End If
            Next
        Else
            r.Borders(xlInsideVertical).Weight = getBorderWeight(s.VInsideLine)
            r.Borders(xlInsideVertical).LineStyle = getLineStyle(s.VInsideLine)
        End If
        
        'ヘッダ線
        For i = 1 To s.HHeadLineCount
            If i > r.Rows.count Then
                Exit For
            End If
            
            If s.EnableHRepeat Or i = lngHLine Then
                r.Rows(i).Borders(xlEdgeBottom).Weight = getBorderWeight(s.HHeadBorderLine)
                r.Rows(i).Borders(xlEdgeBottom).LineStyle = getLineStyle(s.HHeadBorderLine)
            Else
                r.Rows(i).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
            End If
            
            If s.EnableHogan Then
                For j = 2 To r.Columns.count
                    If r(z, j).value <> "" Then
                        r(i, j).Borders(xlEdgeLeft).Weight = getBorderWeight(s.VHeadBorderLine)
                        r(i, j).Borders(xlEdgeLeft).LineStyle = getLineStyle(s.VHeadBorderLine)
                    Else
                        r(i, j).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                    End If
                Next
            Else
                r.Rows(i).Borders(xlInsideVertical).Weight = getBorderWeight(s.VHeadBorderLine)
                r.Rows(i).Borders(xlInsideVertical).LineStyle = getLineStyle(s.VHeadBorderLine)
            End If
            r.Rows(i).Interior.Color = s.HeadColor
        Next
        
        If s.EnableEvenColor Then
            For i = s.HHeadLineCount + 2 To r.Rows.count Step 2
                r.Rows(i).Interior.Color = s.EvenColor
            Next
            
        End If
        
        For i = 1 To s.VHeadLineCount
            If i > r.Columns.count Then
                Exit For
            End If
            If s.EnableVRepeat Or i = lngVLine Then
                r.Columns(i).Borders(xlEdgeRight).Weight = getBorderWeight(s.VHeadBorderLine)
                r.Columns(i).Borders(xlEdgeRight).LineStyle = getLineStyle(s.VHeadBorderLine)
            Else
                r.Columns(i).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
            End If
            r.Columns(i).Borders(xlInsideHorizontal).Weight = getBorderWeight(s.HHeadBorderLine)
            r.Columns(i).Borders(xlInsideHorizontal).LineStyle = getLineStyle(s.HHeadBorderLine)
            r.Columns(i).Interior.Color = s.HeadColor
        Next
        
        '外周
        r.Borders(xlEdgeTop).Weight = getBorderWeight(s.OutSideLine)
        r.Borders(xlEdgeTop).LineStyle = getLineStyle(s.OutSideLine)
        
        r.Borders(xlEdgeLeft).Weight = getBorderWeight(s.OutSideLine)
        r.Borders(xlEdgeLeft).LineStyle = getLineStyle(s.OutSideLine)
        
        r.Borders(xlEdgeRight).Weight = getBorderWeight(s.OutSideLine)
        r.Borders(xlEdgeRight).LineStyle = getLineStyle(s.OutSideLine)
        
        r.Borders(xlEdgeBottom).Weight = getBorderWeight(s.OutSideLine)
        r.Borders(xlEdgeBottom).LineStyle = getLineStyle(s.OutSideLine)

    Next
    
    Application.OnUndo "Undo", "execUndo"

End Sub

'--------------------------------------------------------------
'　レジストリ設定値取得
'--------------------------------------------------------------
Public Function getPropertyKantan() As Collection

    Dim strBuf As String
    Dim k As KantanLine
    Dim lngMax As Long
    Dim i As Long

    Dim col As Collection

    Set col = New Collection
    
    lngMax = Val(GetSetting(C_TITLE, "KantanDx", "Count", -1))
    
    If lngMax = -1 Then

        Set k = New KantanLine
        
        k.Text = "標準"
        
        k.HHeadLineCount = 1
        k.VHeadLineCount = 0
        
        k.HeadColor = 16764057
        
        k.OutSideLine = 7
        k.HHeadBorderLine = 7
        k.VHeadBorderLine = 7
        k.HInsideLine = 2
        k.VInsideLine = 7
        
        k.EnableEvenColor = False
        k.EvenColor = 16777164
        
        k.EnableHogan = False
        k.EnableHRepeat = False
        k.EnableVRepeat = False
        
        k.AuthoHogan = True
        k.HoganJudgeLineCount = 1

        col.Add k

        Set k = Nothing

        Set k = New KantanLine
        
        k.Text = "標準ヘッダ２行"
        
        k.HHeadLineCount = 2
        k.VHeadLineCount = 0
        
        k.HeadColor = 16764057
        
        k.OutSideLine = 7
        k.HHeadBorderLine = 7
        k.VHeadBorderLine = 7
        k.HInsideLine = 2
        k.VInsideLine = 7
        
        k.EnableEvenColor = False
        k.EvenColor = 16777164
        
        k.EnableHogan = False
        k.EnableHRepeat = False
        k.EnableVRepeat = False
        
        k.AuthoHogan = True
        k.HoganJudgeLineCount = 1

        col.Add k

        Set k = Nothing
        
        
        Set k = New KantanLine
        
        k.Text = "行ヘッダ１列ヘッダ列１"
        
        k.HHeadLineCount = 1
        k.VHeadLineCount = 1
        
        k.HeadColor = 16764057
        
        k.OutSideLine = 7
        k.HHeadBorderLine = 7
        k.VHeadBorderLine = 7
        k.HInsideLine = 2
        k.VInsideLine = 7
        
        k.EnableEvenColor = False
        k.EvenColor = 16777164
        
        k.EnableHogan = False
        k.EnableHRepeat = False
        k.EnableVRepeat = False
        
        k.AuthoHogan = True
        k.HoganJudgeLineCount = 1

        col.Add k

        Set k = Nothing
        
        Set k = New KantanLine
        
        k.Text = "方眼紙"
        
        k.HHeadLineCount = 1
        k.VHeadLineCount = 0
        
        k.HeadColor = 16764057
        
        k.OutSideLine = 7
        k.HHeadBorderLine = 7
        k.VHeadBorderLine = 7
        k.HInsideLine = 2
        k.VInsideLine = 7
        
        k.EnableEvenColor = False
        k.EvenColor = 16777164
        
        k.EnableHogan = True
        k.EnableHRepeat = False
        k.EnableVRepeat = False
        
        k.AuthoHogan = True
        k.HoganJudgeLineCount = 1

        col.Add k

        Set k = Nothing

        Set k = New KantanLine
        
        k.Text = "しましまブルー"
        
        k.HHeadLineCount = 1
        k.VHeadLineCount = 0
        
        k.HeadColor = 16764057
        
        k.OutSideLine = 7
        k.HHeadBorderLine = 7
        k.VHeadBorderLine = 7
        k.HInsideLine = 2
        k.VInsideLine = 7
        
        k.EnableEvenColor = True
        k.EvenColor = 16777164
        
        k.EnableHogan = False
        k.EnableHRepeat = False
        k.EnableVRepeat = False

        k.AuthoHogan = True
        k.HoganJudgeLineCount = 1
        
        col.Add k

        Set k = Nothing
    
        Set k = New KantanLine
        
        k.Text = "しましまベージュ"
        
        k.HHeadLineCount = 1
        k.VHeadLineCount = 0
        
        k.HeadColor = 10079487
        
        k.OutSideLine = 7
        k.HHeadBorderLine = 7
        k.VHeadBorderLine = 7
        k.HInsideLine = 2
        k.VInsideLine = 7
        
        k.EnableEvenColor = True
        k.EvenColor = 10092543
        
        k.EnableHogan = False
        k.EnableHRepeat = False
        k.EnableVRepeat = False

        k.AuthoHogan = True
        k.HoganJudgeLineCount = 1
        
        col.Add k

        Set k = Nothing
        
        Set k = New KantanLine
        
        k.Text = "しましまグリーン"
        
        k.HHeadLineCount = 1
        k.VHeadLineCount = 0
        
        k.HeadColor = 52377
        
        k.OutSideLine = 7
        k.HHeadBorderLine = 7
        k.VHeadBorderLine = 7
        k.HInsideLine = 2
        k.VInsideLine = 7
        
        k.EnableEvenColor = True
        k.EvenColor = 13434828
        
        k.EnableHogan = False
        k.EnableHRepeat = False
        k.EnableVRepeat = False

        k.AuthoHogan = True
        k.HoganJudgeLineCount = 1
        
        col.Add k

        Set k = Nothing
        
        Set k = New KantanLine
        
        k.Text = "しましまピンク"
        
        k.HHeadLineCount = 1
        k.VHeadLineCount = 0
        
        k.HeadColor = 13408767
        
        k.OutSideLine = 7
        k.HHeadBorderLine = 7
        k.VHeadBorderLine = 7
        k.HInsideLine = 2
        k.VInsideLine = 7
        
        k.EnableEvenColor = True
        k.EvenColor = 16764159
        
        k.EnableHogan = False
        k.EnableHRepeat = False
        k.EnableVRepeat = False

        k.AuthoHogan = True
        k.HoganJudgeLineCount = 1
        
        col.Add k

        Set k = Nothing
    
    
    Else
        For i = 0 To lngMax - 1

            strBuf = GetSetting(C_TITLE, "KantanDx", Format(i, "000"), "")
            
            Set k = deserialize(strBuf)
        
            col.Add k

            Set k = Nothing
        Next
    End If

    Set getPropertyKantan = col

End Function
'--------------------------------------------------------------
'　レジストリ設定値取得
'--------------------------------------------------------------
Public Sub setPropertyKantan(ByRef col As Collection)

    Dim strBuf As String
    Dim s As KantanLine
    Dim lngMax As Long
    Dim i As Long

    On Error Resume Next
    
    DeleteSetting C_TITLE, "KantanDx"

    For i = 0 To col.count - 1

        Set s = col(i + 1)
        
        strBuf = serialize(s)
        Call SaveSetting(C_TITLE, "KantanDx", Format$(i, "000"), strBuf)

        Set s = Nothing
    Next

    Call SaveSetting(C_TITLE, "KantanDx", "Count", col.count)

End Sub
Function serialize(ByRef k As KantanLine) As String

    Dim strBuf As String
    
    'テキスト
    strBuf = strBuf & k.Text & vbVerticalTab
    
    strBuf = strBuf & k.OutSideLine & vbVerticalTab

    '行ヘッダの行数
    strBuf = strBuf & k.HHeadLineCount & vbVerticalTab

    ''行ヘッダと行の間の線
    strBuf = strBuf & k.HHeadBorderLine & vbVerticalTab
    
    '行ヘッダの色
    strBuf = strBuf & k.HeadColor & vbVerticalTab

    ''明細色設定
    strBuf = strBuf & k.EnableEvenColor & vbVerticalTab

    '明細色
    strBuf = strBuf & k.EvenColor & vbVerticalTab

    '列ヘッダの行数
    strBuf = strBuf & k.VHeadLineCount & vbVerticalTab

    '列ヘッダの列の間の線
    strBuf = strBuf & k.VHeadBorderLine & vbVerticalTab
    
    '縦中線
    strBuf = strBuf & k.VInsideLine & vbVerticalTab
    
    '横中線
    strBuf = strBuf & k.HInsideLine & vbVerticalTab

    '方眼紙設定
    strBuf = strBuf & k.EnableHogan & vbVerticalTab

    '行ヘッダ線繰り返し
    strBuf = strBuf & k.EnableHRepeat & vbVerticalTab
    
    '列ヘッダ線繰り返し
    strBuf = strBuf & k.EnableVRepeat & vbVerticalTab
    
    '方眼紙判定行自動判定
    strBuf = strBuf & k.AuthoHogan & vbVerticalTab
    
    '方眼紙判定行
    strBuf = strBuf & k.HoganJudgeLineCount & vbVerticalTab
    
    serialize = strBuf

End Function
Function deserialize(ByVal strBuf As String) As KantanLine

    Dim k As KantanLine
    Dim varBuf As Variant
    Dim i As Long
    
    varBuf = Split(strBuf, vbVerticalTab)
    
    Set k = New KantanLine

    i = 0
    
    k.Text = varBuf(i)
    i = i + 1
    
    '外周
    k.OutSideLine = varBuf(i)
    i = i + 1

    '行ヘッダの行数
    k.HHeadLineCount = varBuf(i)
    i = i + 1

    ''行ヘッダと行の間の線
    k.HHeadBorderLine = varBuf(i)
    i = i + 1
    
    '行ヘッダの色
    k.HeadColor = varBuf(i)
    i = i + 1

    ''明細色設定
    k.EnableEvenColor = varBuf(i)
    i = i + 1

    '明細色
    k.EvenColor = varBuf(i)
    i = i + 1

    '列ヘッダの行数
    k.VHeadLineCount = varBuf(i)
    i = i + 1

    '列ヘッダの列の間の線
    k.VHeadBorderLine = varBuf(i)
    i = i + 1
    
    '縦中線
    k.VInsideLine = varBuf(i)
    i = i + 1
    
    '横中線
    k.HInsideLine = varBuf(i)
    i = i + 1

    '方眼紙設定
    k.EnableHogan = varBuf(i)
    i = i + 1
    
    '行ヘッダ線繰り返し
    k.EnableHRepeat = varBuf(i)
    i = i + 1
    
    '列ヘッダ線繰り返し
    k.EnableVRepeat = varBuf(i)
    i = i + 1
    
    '方眼紙判定行自動判定
    k.AuthoHogan = varBuf(i)
    i = i + 1

    '方眼紙判定行
    k.HoganJudgeLineCount = varBuf(i)
    i = i + 1

    Set deserialize = k

End Function

Function getLineStyle(ByVal tag As String) As XlLineStyle
    
    Dim bStyle As XlLineStyle
    
    Select Case Val(tag)
        Case 1
            bStyle = xlLineStyleNone
        Case 2
            bStyle = xlContinuous
        Case 3
            bStyle = xlDot
        Case 4
            bStyle = xlDashDotDot
        Case 5
            bStyle = xlDashDot
        Case 6
            bStyle = xlDash
        Case 7
            bStyle = xlContinuous
        Case 8
            bStyle = xlDashDotDot
        Case 9
            bStyle = xlSlantDashDot
        Case 10
            bStyle = xlDashDot
        Case 11
            bStyle = xlDash
        Case 12
            bStyle = xlContinuous
        Case 13
            bStyle = xlContinuous
        Case 14
            bStyle = xlDouble
    End Select
    
    getLineStyle = bStyle

End Function
Function getBorderWeight(ByVal tag As String) As XlBorderWeight
    
    Dim bLine As XlBorderWeight
    
    Select Case Val(tag)
        Case 1
            bLine = xlHairline
        Case 2
            bLine = xlHairline
        Case 3
            bLine = xlThin
        Case 4
            bLine = xlThin
        Case 5
            bLine = xlThin
        Case 6
            bLine = xlThin
        Case 7
            bLine = xlThin
        Case 8
            bLine = xlMedium
        Case 9
            bLine = xlMedium
        Case 10
            bLine = xlMedium
        Case 11
            bLine = xlMedium
        Case 12
            bLine = xlMedium
        Case 13
            bLine = xlThick
        Case 14
            bLine = xlThick
    End Select
    
    getBorderWeight = bLine

End Function

Sub showKantanDx()
    frmKantanDx.Show
End Sub
'--------------------------------------------------------------
'　選択範囲の表を取り込む
'--------------------------------------------------------------
Sub kantanScan()

'    '設定情報取得
'    Dim col As Collection
'    Dim s As KantanLine
'    Dim i As Long
'    Dim j As Long
'    Dim z As Long
'
'    Dim lngHLine As Long
'    Dim lngVLine As Long
'
'    Dim r As Range
'
'    Set col = getPropertyKantan()
'    If col Is Nothing Then
'        Exit Sub
'    End If
'
'
'#If VBA7 Then
'    Set r = Selection.DisplayFormat
'#Else
'    Set r = Selection
'#End If
'
'    If r.Rows.count < 3 Or r.Columns.count <= 3 Then
'        MsgBox "選択範囲が小さいため判定できません。", vbExclamation + vbOKOnly, C_TITLE
'        Exit Sub
'    End If
'
'    Set s = New KantanLine
'
'    '名前
'    s.Text = Format(Now, "yyyy-mm-dd hh:nn:ss")
'
'    '外周
'    s.OutSideLine = getLineNo(r.Borders(xlEdgeTop))
'
'
'    Dim i As Long
'    Dim vatPattern As Variant
'
'    '１行目に色がついている場合列ヘッダ有とみなす
'
'
'    For i = 1 To r.Rows.count - 1
'
'        If r(1, r.Columns.count).Interior.Color.Pattern <> xlNone And r(1, r.Columns.count).Interior.Color.RGB <> vbWhite Then
'
'        End If
'
'        If r(i, r.Columns.count).boders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then
'
'
'        End If
'
'
'
'        'パターン無の場合、色はないと思われる
'        If vatPattern <> r(i, r.Columns.count).Interior.Color.Pattern Then
'
'        lngNo = getLineNo(r(i, r.Columns.count).boders(xlEdgeBottom))
'
'
'
'    Next
'
''        r(i, r.Columns.count).boders (xlEdgeBottom)
'
'    For i = 1 To r.Columns.count
'
'
'    Next
'
'    '行ヘッダの行数
'    s.HHeadLineCount
'
'    '行ヘッダと行の間の線
'    s.HHeadBorderLine getLineNo(r.Borders(xlEdgeTop))
'
'    '行ヘッダの色
'    s.HeadColor
'
'    '明細色設定
'    s.EnableEvenColor
'
'    '明細色
'    s.EvenColor
'    '
'    '列ヘッダの行数
'    s.VHeadLineCount
'
'    '列ヘッダの列の間の線
'    s.VHeadBorderLine = getLineNo(r.Borders(xlEdgeTop))
'
'    '縦中線
'    s.VInsideLine = getLineNo(r.Borders(xlEdgeTop))
'
'    '横中線
'    s.HInsideLine = getLineNo(r.Borders(xlEdgeTop))
'
'    '方眼紙設定
'    s.EnableHogan
'
'    '行ヘッダ線の繰り返し
'    s.EnableHRepeat
'
'    '列ヘッダ線の繰り返し
'    s.EnableVRepeat
'
'    '方眼紙行自動設定
'    s.AuthoHogan
'
'    '方眼紙判定行
'    s.HoganJudgeLineCount
'
'    col.Add s
'
'    Call setPropertyKantan(col)

End Sub
Private Function getLineNo(ByRef b As Border) As Long

    Dim lngRet As Long
    
    Select Case True
        Case b.LineStyle = xlLineStyleNone
            lngRet = 1
        
        Case b.Weight = xlHairline And b.LineStyle = xlContinuous
            lngRet = 2
        
        Case b.Weight = xlThin And b.LineStyle = xlDot
            lngRet = 3
        
        Case b.Weight = xlThin And b.LineStyle = xlDashDotDot
            lngRet = 4
        
        Case b.Weight = xlThin And b.LineStyle = xlDashDot
            lngRet = 5
        
        Case b.Weight = xlThin And b.LineStyle = xlDash
            lngRet = 6
        
        Case b.Weight = xlThin And b.LineStyle = xlContinuous
            lngRet = 7
        
        Case b.Weight = xlMedium And b.LineStyle = xlDashDotDot
            lngRet = 8
        
        Case b.Weight = xlMedium And b.LineStyle = xlSlantDashDot
            lngRet = 9
        
        Case b.Weight = xlMedium And b.LineStyle = xlDashDot
            lngRet = 10
        
        Case b.Weight = xlMedium And b.LineStyle = xlDash
            lngRet = 11
        
        Case b.Weight = xlMedium And b.LineStyle = xlContinuous
            lngRet = 12
            
        Case b.Weight = xlThick And b.LineStyle = xlContinuous
            lngRet = 13
        
        Case b.Weight = xlThick And b.LineStyle = xlDouble
            lngRet = 14
            
        Case Else
            lngRet = 0
    End Select

    lngRet = getLineNo

End Function

