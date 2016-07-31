VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGridText 
   Caption         =   "表のテキスト化"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   OleObjectBlob   =   "frmGridText.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmGridText"
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

Private Const C_LINE_MAX As Long = 1000
Private Const C_DEFAULT_COL As Long = 40
Private mlngMaxKeta As Long
Private mlngMinKeta As Long

Private mlngWidth() As Long
Private mlngHeight() As Long
Private mlngMaxColWidth() As Long
Private mblnFixColumn() As Boolean
    
Private Type typGrid
    Text() As String
    TextCount As Long
    TextMaxLength As Long
    Align As Long
    vAlign As Long
    ColSpan As Long
    RowSpan As Long
    WrapText As Boolean
    NoWrapField As Boolean
End Type

Private mudtGrid() As typGrid

Private mlngRow As Long
Private mlngCol As Long

Private mlngWidthMax As Long
Private mlngHeightMax As Long

'罫線の太さ条件
Private mlngBorderWeight1 As Long
Private mlngBorderWeight2 As Long

Private Const C_LINE_WIDTH As Long = 2
    
Private Const C_BORDER_NONE     As Long = 0     'なし
Private Const C_BORDER_TOP      As Long = 1     '上
Private Const C_BORDER_BOTTOM   As Long = 2     '下
Private Const C_BORDER_LEFT     As Long = 4     '左
Private Const C_BORDER_RIGHT    As Long = 8     '右

Private Const C_BORDER_TOP_BOLD      As Long = 16     '上
Private Const C_BORDER_BOTTOM_BOLD   As Long = 32     '下
Private Const C_BORDER_LEFT_BOLD     As Long = 64     '左
Private Const C_BORDER_RIGHT_BOLD    As Long = 128    '右

Private Const C_BORDER_LR As Long = C_BORDER_LEFT + C_BORDER_RIGHT      '─(よこ)
Private Const C_BORDER_TB As Long = C_BORDER_TOP + C_BORDER_BOTTOM      '│(たて)
Private Const C_BORDER_TL As Long = C_BORDER_TOP + C_BORDER_LEFT        '┘(右下)
Private Const C_BORDER_TR As Long = C_BORDER_TOP + C_BORDER_RIGHT       '└(左下)
Private Const C_BORDER_BR As Long = C_BORDER_BOTTOM + C_BORDER_RIGHT    '┌(左上)
Private Const C_BORDER_BL As Long = C_BORDER_BOTTOM + C_BORDER_LEFT     '┐(右上)
Private Const C_BORDER_TBR As Long = C_BORDER_TOP + C_BORDER_BOTTOM + C_BORDER_RIGHT    '├(縦右)
Private Const C_BORDER_BLT As Long = C_BORDER_BOTTOM + C_BORDER_LEFT + C_BORDER_RIGHT   '┬(横下)
Private Const C_BORDER_TBL As Long = C_BORDER_TOP + C_BORDER_BOTTOM + C_BORDER_LEFT     '┤(縦左)
Private Const C_BORDER_TLR As Long = C_BORDER_TOP + C_BORDER_LEFT + C_BORDER_RIGHT      '┴(横上)
Private Const C_BORDER_CROSS As Long = C_BORDER_TOP + C_BORDER_BOTTOM + C_BORDER_LEFT + C_BORDER_RIGHT  '┼(真中)

Private Const C_BORDER_LR_BOLD As Long = C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD      '━(よこ)
Private Const C_BORDER_TB_BOLD As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD      '┃(たて)
Private Const C_BORDER_TL_BOLD As Long = C_BORDER_TOP_BOLD + C_BORDER_LEFT_BOLD        '┛(右下)
Private Const C_BORDER_TR_BOLD As Long = C_BORDER_TOP_BOLD + C_BORDER_RIGHT_BOLD       '┗(左下)
Private Const C_BORDER_BR_BOLD As Long = C_BORDER_BOTTOM_BOLD + C_BORDER_RIGHT_BOLD    '┏(左上)
Private Const C_BORDER_BL_BOLD As Long = C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT_BOLD     '┓(右上)
Private Const C_BORDER_TBR_BOLD As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD + C_BORDER_RIGHT_BOLD    '┣(縦右)
Private Const C_BORDER_BLT_BOLD As Long = C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD   '┳(横下)
Private Const C_BORDER_TBL_BOLD As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT_BOLD     '┫(縦左)
Private Const C_BORDER_TLR_BOLD As Long = C_BORDER_TOP_BOLD + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD      '┻(横上)
Private Const C_BORDER_CROSS_BOLD As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD  '╋(真中)

Private Const C_BORDER_TBR_BH As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD + C_BORDER_RIGHT    '┠(縦右)
Private Const C_BORDER_BLT_BH As Long = C_BORDER_BOTTOM + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD   '┯(横下)
Private Const C_BORDER_TBL_BH As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT     '┨(縦左)
Private Const C_BORDER_TLR_BH As Long = C_BORDER_TOP + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD      '┷(横上)
Private Const C_BORDER_CROSS_BH As Long = C_BORDER_TOP + C_BORDER_BOTTOM + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD  '┿(真中)

Private Const C_BORDER_TBR_HB As Long = C_BORDER_TOP + C_BORDER_BOTTOM + C_BORDER_RIGHT_BOLD    '┝(縦右)
Private Const C_BORDER_BLT_HB As Long = C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT + C_BORDER_RIGHT   '┰(横下)
Private Const C_BORDER_TBL_HB As Long = C_BORDER_TOP + C_BORDER_BOTTOM + C_BORDER_LEFT_BOLD     '┥(縦左)
Private Const C_BORDER_TLR_HB As Long = C_BORDER_TOP_BOLD + C_BORDER_LEFT + C_BORDER_RIGHT      '┸(横上)
Private Const C_BORDER_CROSS_HB As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT + C_BORDER_RIGHT  '╂(真中)

Private Const C_BORDER_TL_BH As Long = C_BORDER_TOP_BOLD + C_BORDER_LEFT             '┛(右下)
Private Const C_BORDER_TR_BH As Long = C_BORDER_TOP_BOLD + C_BORDER_RIGHT            '┗(左下)
Private Const C_BORDER_BR_BH As Long = C_BORDER_BOTTOM_BOLD + C_BORDER_RIGHT         '┏(左上)
Private Const C_BORDER_BL_BH As Long = C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT          '┓(右上)



'罫線だけでは表現できないもの
Private Const C_BORDER_TL_HB As Long = C_BORDER_TOP + C_BORDER_LEFT_BOLD             '┛(右下)
Private Const C_BORDER_TR_HB As Long = C_BORDER_TOP + C_BORDER_RIGHT_BOLD            '┗(左下)
Private Const C_BORDER_BR_HB As Long = C_BORDER_BOTTOM + C_BORDER_RIGHT_BOLD         '┏(左上)
Private Const C_BORDER_BL_HB As Long = C_BORDER_BOTTOM + C_BORDER_LEFT_BOLD          '┓(右上)

'下だけ細い十字
Private Const C_BORDER_CROSS_BOLD_UL As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD  '╋(真中)
Private Const C_BORDER_CROSS_UB As Long = C_BORDER_TOP + C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT_BOLD  '╂(右だけ細い十字)
Private Const C_BORDER_CROSS_RB As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT  '╂(右だけ細い十字)
Private Const C_BORDER_CROSS_LB As Long = C_BORDER_TOP_BOLD + C_BORDER_BOTTOM_BOLD + C_BORDER_LEFT + C_BORDER_RIGHT_BOLD  '╂(左だけ細い十字)

Private Const C_BORDER_TLR_TLBOLD As Long = C_BORDER_TOP_BOLD + C_BORDER_LEFT_BOLD + C_BORDER_RIGHT      '┻(横上)

Private Const C_SQUARE_TOP_LEFT As Long = 1
Private Const C_SQUARE_TOP_MIDDLE As Long = 2
Private Const C_SQUARE_TOP_RIGHT As Long = 3
Private Const C_SQUARE_LEFT_MIDDLE As Long = 4
Private Const C_SQUARE_RIGHT_MIDDLE As Long = 5
Private Const C_SQUARE_BOTTOM_LEFT As Long = 6
Private Const C_SQUARE_BOTTOM_MIDDLE As Long = 7
Private Const C_SQUARE_BOTTOM_RIGHT As Long = 8
Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1
'--------------------------------------------------------------
'　かんたん罫線
'--------------------------------------------------------------
Private Sub kantanLineRun()

    Dim lngJitsuLineMax As Long
    Dim lngIdxRow As Long
    Dim lngIdxCol As Long
    Dim i As Long
    
    mlngRow = Selection.Rows.count
    mlngCol = Selection.Columns.count

    'メモリの確保
    ReDim mlngWidth(1 To Selection.Columns.count)
    ReDim mlngMaxColWidth(1 To Selection.Columns.count)
    ReDim mlngHeight(1 To Selection.Rows.count)
    ReDim mudtGrid(1 To mlngRow, 1 To mlngCol)
    ReDim mblnFixColumn(1 To Selection.Columns.count)

    '最大幅より、罫線幅を引いた実際の幅
    lngJitsuLineMax = mlngMaxKeta - (mlngCol + 1) * 2

    If Selection.Areas.count > 1 Then
        Exit Sub
    End If
    
    '表データのセット(一次）
    Call setGridData

    '各列の最大桁数を求める
    For lngIdxRow = 1 To mlngRow
        For lngIdxCol = 1 To mlngCol
            
           Dim lngLen As Long
           
            'マージセル以外
            If mudtGrid(lngIdxRow, lngIdxCol).ColSpan = 1 Then
                lngLen = mudtGrid(lngIdxRow, lngIdxCol).TextMaxLength
                
                '奇数だった場合＋１
                If lngLen Mod 2 = 1 Then
                    lngLen = lngLen + 1
                End If
            
                If lngLen > mlngMaxColWidth(lngIdxCol) Then
                    mlngMaxColWidth(lngIdxCol) = lngLen
                End If
               
                If mudtGrid(lngIdxRow, lngIdxCol).NoWrapField Then
                    mblnFixColumn(lngIdxCol) = True
                End If
            End If
        Next
    Next
    
    '結合セルの幅を設定する。
    For lngIdxRow = 1 To mlngRow
        For lngIdxCol = 1 To mlngCol
            
            If mudtGrid(lngIdxRow, lngIdxCol).ColSpan > 1 Then
            
                lngLen = mudtGrid(lngIdxRow, lngIdxCol).TextMaxLength
                
                '奇数だった場合＋１
                If lngLen Mod 2 = 1 Then
                    lngLen = lngLen + 1
                End If
            
                Dim lngSize As Long
                lngSize = 0
                For i = 1 To mudtGrid(lngIdxRow, lngIdxCol).ColSpan
                    lngSize = lngSize + mlngMaxColWidth(lngIdxCol + i - 1)
                Next
                
                '結合しているセルの内容の方が大きい場合
                If lngLen > lngSize Then
                
                    Dim lngSa As Long
                    lngSa = (lngLen - lngSize) \ mudtGrid(lngIdxRow, lngIdxCol).ColSpan
                
                    '奇数だった場合＋１
                    If lngSa Mod 2 = 1 Then
                        lngSa = lngSa + 1
                    End If
                
                    '結合セルが表示されるように各セルに割り振る
                    For i = 1 To mudtGrid(lngIdxRow, lngIdxCol).ColSpan
                        mlngMaxColWidth(lngIdxCol + i - 1) = mlngMaxColWidth(lngIdxCol + i - 1) + lngSa
                    Next
                
                End If
                
            End If
               
        Next
    Next
    
    
    
    '表全体の幅を求める。
    mlngWidthMax = 0

    For lngIdxCol = 1 To mlngCol
        mlngWidthMax = mlngWidthMax + mlngMaxColWidth(lngIdxCol)
    Next
    
    '表の幅が実幅より超えていた場合
    If mlngWidthMax > lngJitsuLineMax Then
        '現在のセル幅の割合に応じて最大桁数分割り振る。

        Dim lngWk As Long
        Dim lngMaxPos As Long
        Dim lngMaxWidth As Long
        Dim lngAmari As Long
        
        '固定フィールドの数を数える
        Dim lngFix As Long
        Dim lngFixSize As Long
        lngFix = 0
        lngFixSize = 0
        For lngIdxCol = 1 To mlngCol
            If mblnFixColumn(lngIdxCol) Then
                lngFix = lngFix + 1
                lngFixSize = lngFixSize + mlngMaxColWidth(lngIdxCol)
            End If
        Next
    
        Dim lngNoFix As Long
        Dim lngNoFixSize As Long
    
        lngNoFix = mlngCol - lngNoFix
        lngNoFixSize = mlngWidthMax - lngFixSize
    
        Dim lngDelSize As Long
        
        lngDelSize = (lngJitsuLineMax - lngFixSize)
        If lngDelSize < 0 Then
            lngDelSize = 0
        End If
    
        For lngIdxCol = 1 To mlngCol
            '割合で各列の幅を計算する。
            If mblnFixColumn(lngIdxCol) Then
                mlngWidth(lngIdxCol) = mlngMaxColWidth(lngIdxCol)
            Else
                mlngWidth(lngIdxCol) = Fix(lngDelSize * (mlngMaxColWidth(lngIdxCol) / lngNoFixSize))
                '奇数になってしまった場合－１
                If mlngWidth(lngIdxCol) Mod 2 = 1 Then
                    mlngWidth(lngIdxCol) = mlngWidth(lngIdxCol) - 1
                End If
                If mlngWidth(lngIdxCol) <= 1 Then
                    mlngWidth(lngIdxCol) = 2
                End If
            End If
            
            If lngMaxWidth < mlngWidth(lngIdxCol) Then
                '最大の列を求める。
                lngMaxWidth = mlngWidth(lngIdxCol)
                lngMaxPos = lngIdxCol
            End If
            lngWk = lngWk + mlngWidth(lngIdxCol)
        Next

        '列で最大のものに余りを寄せる
        lngAmari = lngJitsuLineMax - lngWk
        If lngAmari < 0 Then
            lngAmari = 0
        End If
        mlngWidth(lngMaxPos) = mlngWidth(lngMaxPos) + lngAmari
        
    Else
        For lngIdxCol = 1 To mlngCol
            mlngWidth(lngIdxCol) = mlngMaxColWidth(lngIdxCol)
        Next
        
    End If
    
    '表データのセット(二次）
    '決定した幅で内容をつめなおす。
    Call setGridData

    '各行の最大の高さを求める。
    For lngIdxRow = 1 To mlngRow
        For lngIdxCol = 1 To mlngCol
            
            If mudtGrid(lngIdxRow, lngIdxCol).RowSpan = 1 Then
                lngLen = mudtGrid(lngIdxRow, lngIdxCol).TextCount
                If lngLen > mlngHeight(lngIdxRow) Then
                   mlngHeight(lngIdxRow) = lngLen
                End If
            End If
        Next
    Next
    
    '各行の最大の高さを求める。
    For lngIdxRow = 1 To mlngRow
        For lngIdxCol = 1 To mlngCol
        
            If mudtGrid(lngIdxRow, lngIdxCol).RowSpan > 1 Then
                
                lngLen = mudtGrid(lngIdxRow, lngIdxCol).TextCount
                
                lngSize = 0
                For i = 1 To mudtGrid(lngIdxRow, lngIdxCol).RowSpan
                    lngSize = lngSize + mlngHeight(lngIdxRow + i - 1)
                Next
                
                '結合しているセルの内容の方が大きい場合
                If lngLen > lngSize Then
                
                    lngSa = (lngLen - lngSize) \ mudtGrid(lngIdxRow, lngIdxCol).ColSpan
                
                    '結合セルが表示されるように各セルに割り振る
                    For i = 1 To mudtGrid(lngIdxRow, lngIdxCol).RowSpan
                        mlngHeight(lngIdxRow + i - 1) = mlngHeight(lngIdxRow + i - 1) + lngSa
                    Next
                
                End If
            
            End If
        Next
    Next
    
'                Dim lngMod As Long
'                Dim lngAns As Long
'                lngLen = mudtGrid(lngIdxRow, lngIdxCol).TextCount
'                lngAns = lngLen \ mudtGrid(lngIdxRow, lngIdxCol).RowSpan
'                lngMod = lngLen Mod mudtGrid(lngIdxRow, lngIdxCol).RowSpan
'                lngLen = lngAns + lngMod
    
    
    
    '表の高さを求める。
    mlngHeightMax = 0

    For lngIdxRow = 1 To mlngRow
        mlngHeightMax = mlngHeightMax + mlngHeight(lngIdxRow)
    Next
    
    '罫線の作成
    Dim strGrid As String
    strGrid = drawGridData()
    
    putClipboard strGrid
    
    txtPreview.Text = strGrid
    txtPreview.SelLength = 0
    txtPreview.SelStart = 0
    
End Sub
'--------------------------------------------------------------
'　罫線の作成
'--------------------------------------------------------------
Private Function drawGridData() As String

    Dim bytGrid() As Byte
    Dim strGrid As String
    Dim strClipBoard As String
    
    strGrid = ""

    Dim lngLineCount As Long
    Dim lngLine As Long
    Dim lngColCount As Long
    Dim lngIdxRow As Long
    Dim lngIdxCol As Long
    Dim lngIdxHeight As Long
    Dim lngPos As Long
    Dim i As Long
    Dim lngTextCount As Long
    
    lngLineCount = 0
    
    'すべての行数を求める。
    For lngIdxRow = 1 To mlngRow
        lngLineCount = lngLineCount + mlngHeight(lngIdxRow)
    Next
    lngLineCount = lngLineCount + (mlngRow + 1)

    For lngIdxCol = 1 To mlngCol
        lngColCount = lngColCount + mlngWidth(lngIdxCol)
    Next
    lngColCount = lngColCount + (mlngCol + 1) * 2

    ReDim bytGrid(1 To lngLineCount, 1 To lngColCount)

    'スペースで初期化
    For lngIdxRow = 1 To lngLineCount
        For lngIdxCol = 1 To lngColCount
            bytGrid(lngIdxRow, lngIdxCol) = &H20
        Next
    Next

    '結合セルの内側の線をクリアする。
    clearInsideLine
    
    '線種を決定する。
    checkWeight

    '--------------------------------------------------------------
    '　罫線の描画
    '--------------------------------------------------------------
    lngLine = 1
    For lngIdxRow = 1 To mlngRow

        '上線
        strGrid = ""
        For lngIdxCol = 1 To mlngCol

            strGrid = strGrid & getLineData(lngIdxRow, lngIdxCol, C_SQUARE_TOP_LEFT)

            Dim strParts As String
            strParts = getLineData(lngIdxRow, lngIdxCol, C_SQUARE_TOP_MIDDLE)
            Dim strMiddle As String
            strMiddle = ""

            Do Until rlxAscLen(strMiddle) >= mlngWidth(lngIdxCol)
                strMiddle = strMiddle & strParts
            Loop
            strGrid = strGrid & strMiddle

            If lngIdxCol = mlngCol Then
                strGrid = strGrid & getLineData(lngIdxRow, lngIdxCol, C_SQUARE_TOP_RIGHT)
            End If

        Next
        Call setByte(strGrid, bytGrid(), lngLine, 1)
        lngLine = lngLine + 1

        '中身
        For lngIdxHeight = 1 To mlngHeight(lngIdxRow)
            strGrid = ""
            For lngIdxCol = 1 To mlngCol

                strGrid = strGrid & getLineData(lngIdxRow, lngIdxCol, C_SQUARE_LEFT_MIDDLE)
                
                strGrid = strGrid & Space$(mlngWidth(lngIdxCol))

                If lngIdxCol = mlngCol Then
                    strGrid = strGrid & getLineData(lngIdxRow, lngIdxCol, C_SQUARE_RIGHT_MIDDLE)
                End If

            Next
            Call setByte(strGrid, bytGrid, lngLine, 1)
            lngLine = lngLine + 1
        Next
        
        '下線
        If lngIdxRow = mlngRow Then
            strGrid = ""
            For lngIdxCol = 1 To mlngCol

                strGrid = strGrid & getLineData(lngIdxRow, lngIdxCol, C_SQUARE_BOTTOM_LEFT)

                strParts = getLineData(lngIdxRow, lngIdxCol, C_SQUARE_BOTTOM_MIDDLE)

                strMiddle = ""

                Do Until rlxAscLen(strMiddle) >= mlngWidth(lngIdxCol)
                    strMiddle = strMiddle & strParts
                Loop
                strGrid = strGrid & strMiddle


                If lngIdxCol = mlngCol Then
                    strGrid = strGrid & getLineData(lngIdxRow, lngIdxCol, C_SQUARE_BOTTOM_RIGHT)
                End If

            Next
            Call setByte(strGrid, bytGrid, lngLine, 1)
            lngLine = lngLine + 1
        End If
    Next

    '--------------------------------------------------------------
    '　値の描画
    '--------------------------------------------------------------
'    lngLine = 1
'    For lngIdxRow = 1 To mlngRow
'
'        '罫線分
'        lngLine = lngLine + 1
'
'        For lngIdxHeight = 1 To mlngHeight(lngIdxRow)
'
'            strGrid = ""
'            For lngIdxCol = 1 To mlngCol
'
'                Dim lngShift As Long
'                '連結セルの場合
'                If mudtGrid(lngIdxRow, lngIdxCol).RowSpan > 1 Then
'
'                    Dim lngHeight As Long
'                    lngHeight = 0
'                    For i = 0 To mudtGrid(lngIdxRow, lngIdxCol).RowSpan - 1
'                        lngHeight = lngHeight + mlngHeight(lngIdxRow + i)
'                    Next
'                    lngHeight = lngHeight + mudtGrid(lngIdxRow, lngIdxCol).RowSpan - 1
'                    Select Case mudtGrid(lngIdxRow, lngIdxCol).vAlign
'                        Case xlGeneral, xlTop
'                            lngShift = 0
'                        Case xlCenter
'                            lngShift = (lngHeight - mudtGrid(lngIdxRow, lngIdxCol).TextCount) \ 2
'                        Case xlBottom
'                            lngShift = lngHeight - mudtGrid(lngIdxRow, lngIdxCol).TextCount
'                    End Select
'
'                Else
'                    lngShift = 0
'                End If
'
'                If lngIdxHeight <= mudtGrid(lngIdxRow, lngIdxCol).TextCount Then
'                    If mudtGrid(lngIdxRow, lngIdxCol).ColSpan > 1 Then
'
'                        Dim lngSize As Long
'
'                        lngSize = 0
'                        For i = 0 To mudtGrid(lngIdxRow, lngIdxCol).ColSpan - 1
'                            lngSize = lngSize + mlngWidth(lngIdxCol + i)
'                        Next
'                        '罫線分をプラス
'                        lngSize = lngSize + (mudtGrid(lngIdxRow, lngIdxCol).ColSpan - 1) * 2
'                        strGrid = setAlign(mudtGrid(lngIdxRow, lngIdxCol).Text(lngIdxHeight), lngSize, mudtGrid(lngIdxRow, lngIdxCol).Align)
'
'                    Else
'                        strGrid = setAlign(mudtGrid(lngIdxRow, lngIdxCol).Text(lngIdxHeight), mlngWidth(lngIdxCol), mudtGrid(lngIdxRow, lngIdxCol).Align)
'                    End If
'
'                    lngPos = getPos(lngIdxCol)
'                    Call setByte(strGrid, bytGrid, lngLine + lngShift, lngPos)
'
'                End If
'
'            Next
'            lngLine = lngLine + 1
'
'        Next
'
'    Next


    lngLine = 1
    For lngIdxRow = 1 To mlngRow

        '罫線分
        lngLine = lngLine + 1


        strGrid = ""
        For lngIdxCol = 1 To mlngCol


            For lngTextCount = 1 To mudtGrid(lngIdxRow, lngIdxCol).TextCount

                Dim lngShift As Long
                '連結セルの場合
                If mudtGrid(lngIdxRow, lngIdxCol).RowSpan > 1 Then

                    Dim lngHeight As Long

                    lngHeight = lngTextCount - 1
                    For i = 0 To mudtGrid(lngIdxRow, lngIdxCol).RowSpan - 1
                        lngHeight = lngHeight + mlngHeight(lngIdxRow + i)
                    Next

                    lngHeight = lngHeight + mudtGrid(lngIdxRow, lngIdxCol).RowSpan - 1
                    Select Case mudtGrid(lngIdxRow, lngIdxCol).vAlign
                        Case xlGeneral, xlTop
                            lngShift = 0
                        Case xlCenter
                            lngShift = (lngHeight - mudtGrid(lngIdxRow, lngIdxCol).TextCount) \ 2
                        Case xlBottom
                            lngShift = lngHeight - mudtGrid(lngIdxRow, lngIdxCol).TextCount
                    End Select


                Else
                    lngShift = 0
                End If


                If mudtGrid(lngIdxRow, lngIdxCol).ColSpan > 1 Then

                    Dim lngSize As Long

                    lngSize = 0
                    For i = 0 To mudtGrid(lngIdxRow, lngIdxCol).ColSpan - 1
                        lngSize = lngSize + mlngWidth(lngIdxCol + i)
                    Next
                    '罫線分をプラス
                    lngSize = lngSize + (mudtGrid(lngIdxRow, lngIdxCol).ColSpan - 1) * 2
                    strGrid = setAlign(mudtGrid(lngIdxRow, lngIdxCol).Text(lngTextCount), lngSize, mudtGrid(lngIdxRow, lngIdxCol).Align)

                Else
                
                    Select Case mudtGrid(lngIdxRow, lngIdxCol).vAlign
                        Case xlGeneral, xlTop
                            lngShift = 0
                        Case xlCenter
                            lngShift = (mlngHeight(lngIdxRow) - mudtGrid(lngIdxRow, lngIdxCol).TextCount) \ 2
                        Case xlBottom
                            lngShift = mlngHeight(lngIdxRow) - mudtGrid(lngIdxRow, lngIdxCol).TextCount
                    End Select
                    
                    strGrid = setAlign(mudtGrid(lngIdxRow, lngIdxCol).Text(lngTextCount), mlngWidth(lngIdxCol), mudtGrid(lngIdxRow, lngIdxCol).Align)
                
                End If
                
                If lngShift < 0 Then
                    lngShift = 0
                End If
                lngShift = lngShift + lngTextCount - 1

                lngPos = getPos(lngIdxCol)
                Call setByte(strGrid, bytGrid, lngLine + lngShift, lngPos)

           Next

        Next
        lngLine = lngLine + mlngHeight(lngIdxRow)

    Next



    '--------------------------------------------------------------
    '　文字列の組み立て
    '--------------------------------------------------------------
    Dim strBuf As String
    strClipBoard = ""
    For lngIdxRow = 1 To lngLineCount
        strBuf = ""
        For lngIdxCol = 1 To lngColCount
            strBuf = strBuf & ChrB(bytGrid(lngIdxRow, lngIdxCol))
        Next
        strClipBoard = strClipBoard & StrConv(strBuf, vbUnicode) & vbCrLf
    Next

    drawGridData = strClipBoard

End Function
'--------------------------------------------------------------
'　描画桁の取得
'--------------------------------------------------------------
Private Function getPos(ByVal lngIdxCol As Long) As Long

    Dim i As Long
    Dim lngPos As Long

    lngPos = 1 + C_LINE_WIDTH
    For i = 1 To lngIdxCol - 1
        lngPos = lngPos + mlngWidth(i) + C_LINE_WIDTH
    Next

    getPos = lngPos

End Function
'--------------------------------------------------------------
'　文字列をバイト型配列に設定
'--------------------------------------------------------------
Private Sub setByte(ByVal strBuf As String, ByRef bytBuf() As Byte, ByVal lngLine As Long, ByVal lngPos As Long)

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngIdx As Long
    Dim lngTerm As Long
    
    Dim strSjis As String
    
    strSjis = StrConv(strBuf, vbFromUnicode)
    
    lngStart = 1
    lngEnd = LenB(strSjis)
    
    '必ず２バイトずつ設定
    If lngEnd Mod 2 = 1 Then
        strSjis = strSjis & ChrB(&H20)
        lngEnd = lngEnd + 1
    End If
    
    lngTerm = UBound(bytBuf, 2)
    
    For lngIdx = lngStart To lngEnd
    
        '配列以上の場合設定しない
        If lngPos > lngTerm Then
            Exit For
        End If
        
        bytBuf(lngLine, lngPos) = AscB(MidB$(strSjis, lngIdx, 1))
        lngPos = lngPos + 1
        
    Next

End Sub
'--------------------------------------------------------------
'　文字の配置
'--------------------------------------------------------------
Private Function setAlign(ByVal strValue As String, ByVal lngSize As Long, ByVal lngAlign As Long) As String

    Dim strResult As String
    Dim lngLen As Long
    Dim lngLeft As Long

    lngLen = rlxAscLen(strValue)
    
    If lngLen = 0 Then
        setAlign = ""
        Exit Function
    End If
    
    If lngSize - lngLen < 0 Then
        setAlign = strValue
        Exit Function
    End If

    Select Case lngAlign
        Case xlGeneral, xlLeft
            strResult = strValue
            
        Case xlRight
            strResult = Space(lngSize - lngLen) & strValue
        
        Case xlCenter
            lngLeft = (lngSize - lngLen) \ 2
            strResult = Space(lngLeft) & strValue
            
    End Select

    setAlign = strResult

End Function
'--------------------------------------------------------------
'　各セルの内容をワークエリアに保持
'--------------------------------------------------------------
Private Sub setGridData()

    Dim strSrc As String
    Dim lngNewSize As Long
    Dim strBuf As String
    Dim strChr As String
    Dim lngLine As Long
    Dim strLine() As String
    Dim i As Long
    Dim lngLineLen As Long

    Dim lngLen As Long
    
    Dim lngIdxRow As Long
    Dim lngIdxCol As Long
    'Dim i As Long

    
    
    For lngIdxRow = 1 To mlngRow
        For lngIdxCol = 1 To mlngCol

            strSrc = Selection(lngIdxRow, lngIdxCol).Text
            lngNewSize = mlngWidth(lngIdxCol)
            
            'マージセルの場合
            If Selection(lngIdxRow, lngIdxCol).MergeCells Then
                If Selection(lngIdxRow, lngIdxCol).MergeArea(1, 1).Address = Selection(lngIdxRow, lngIdxCol).Address Then
                    lngNewSize = 0
                    For i = 0 To Selection(lngIdxRow, lngIdxCol).MergeArea.Columns.count - 1
                        lngNewSize = lngNewSize + mlngWidth(lngIdxCol + i)
                    Next
                End If
            End If
            
            If lngNewSize = 0 Then
                lngNewSize = C_DEFAULT_COL
            End If
            
            '各種属性設定
'暫定 とりあえず全部折り返す。
'            mudtGrid(lngIdxRow, lngIdxCol).WrapText = Selection(lngIdxRow, lngIdxCol).WrapText
            mudtGrid(lngIdxRow, lngIdxCol).WrapText = True
            
            '横位置が標準以外であればそれにあわせる
            Select Case Selection(lngIdxRow, lngIdxCol).HorizontalAlignment
                Case xlGeneral
                    '書式が文字列なら左寄せ
                    Select Case True
                        Case Selection(lngIdxRow, lngIdxCol).NumberFormatLocal = "@"
                            mudtGrid(lngIdxRow, lngIdxCol).Align = xlLeft
                            
                        Case IsNumeric(Selection(lngIdxRow, lngIdxCol).Value)
                            '数値の場合、右寄せ
                            mudtGrid(lngIdxRow, lngIdxCol).Align = xlRight
                        
                        Case IsDate(Selection(lngIdxRow, lngIdxCol).Value)
                            '日付の場合、右寄せ
                            mudtGrid(lngIdxRow, lngIdxCol).Align = xlRight
                            
                        Case Else
                            mudtGrid(lngIdxRow, lngIdxCol).Align = xlLeft
                    End Select
                    
                Case xlFill, xlJustify
                    '繰り返し,両端揃え
                    mudtGrid(lngIdxRow, lngIdxCol).Align = xlLeft
                
                Case xlCenterAcrossSelection, xlDistributed
                    '選択範囲内で中央, 均等割り付け
                    mudtGrid(lngIdxRow, lngIdxCol).Align = xlCenter
                
                Case Else
                    mudtGrid(lngIdxRow, lngIdxCol).Align = Selection(lngIdxRow, lngIdxCol).HorizontalAlignment
            End Select
            
            '縦位置
            Select Case Selection(lngIdxRow, lngIdxCol).VerticalAlignment
                Case xlJustify
                    mudtGrid(lngIdxRow, lngIdxCol).vAlign = xlTop
                Case xlDistributed
                    mudtGrid(lngIdxRow, lngIdxCol).vAlign = xlCenter
                Case Else
                    mudtGrid(lngIdxRow, lngIdxCol).vAlign = Selection(lngIdxRow, lngIdxCol).VerticalAlignment
            End Select
            
            Select Case True
                Case IsNumeric(Selection(lngIdxRow, lngIdxCol).Text) And mudtGrid(lngIdxRow, lngIdxCol).Align = xlRight
                    '数値の場合、word wrapしない
                    mudtGrid(lngIdxRow, lngIdxCol).NoWrapField = True
                
                Case IsDate(Selection(lngIdxRow, lngIdxCol).Value) And mudtGrid(lngIdxRow, lngIdxCol).Align = xlRight
                    '日付の場合、word wrapしない
                    mudtGrid(lngIdxRow, lngIdxCol).NoWrapField = True
            
                Case Else
                    mudtGrid(lngIdxRow, lngIdxCol).NoWrapField = False
                    
            End Select
            
            lngLen = Len(strSrc)
            lngLine = 0
            strBuf = ""
            
            Erase strLine
            For i = 1 To lngLen
            
                strChr = Mid(strSrc, i, 1)
                
                Select Case True
                    Case strChr = vbCrLf
                        '改行コードの場合
                        lngLine = lngLine + 1
                        ReDim Preserve strLine(1 To lngLine)
                        strLine(lngLine) = strBuf
                
                        '改行コードを捨てる
                        strBuf = ""
                
                    Case strChr = vbLf Or strChr = vbCr
                        '改行コードの場合
                        lngLine = lngLine + 1
                        ReDim Preserve strLine(1 To lngLine)
                        strLine(lngLine) = strBuf
                
                        '改行コードを捨てる
                        strBuf = ""
                
                    Case rlxAscLen(strBuf & strChr) > lngNewSize And mudtGrid(lngIdxRow, lngIdxCol).WrapText = True And mudtGrid(lngIdxRow, lngIdxCol).NoWrapField = False
'                    Case rlxAscLen(strBuf & strChr)
                        '幅を超える場合
                        lngLine = lngLine + 1
                        ReDim Preserve strLine(1 To lngLine)
                        strLine(lngLine) = strBuf
                        
                        'バッファを初期化
                        strBuf = strChr
                
                    Case Else
                        strBuf = strBuf & strChr
                        
                End Select
            
            Next
            
            lngLine = lngLine + 1
            ReDim Preserve strLine(1 To lngLine)
            strLine(lngLine) = strBuf
            
            mudtGrid(lngIdxRow, lngIdxCol).Text = strLine
            mudtGrid(lngIdxRow, lngIdxCol).TextCount = lngLine
            
            mudtGrid(lngIdxRow, lngIdxCol).ColSpan = 1
            mudtGrid(lngIdxRow, lngIdxCol).RowSpan = 1
            
            If Selection(lngIdxRow, lngIdxCol).MergeCells Then
                If Selection(lngIdxRow, lngIdxCol).MergeArea(1, 1).Address = Selection(lngIdxRow, lngIdxCol).Address Then
                    mudtGrid(lngIdxRow, lngIdxCol).ColSpan = Selection(lngIdxRow, lngIdxCol).MergeArea.Columns.count
                    mudtGrid(lngIdxRow, lngIdxCol).RowSpan = Selection(lngIdxRow, lngIdxCol).MergeArea.Rows.count
                End If
            End If
            
            Dim lngMaxWork As Long
            lngMaxWork = 0
            For i = 1 To lngLine

                Dim lngWk As Long
                lngWk = rlxAscLen(mudtGrid(lngIdxRow, lngIdxCol).Text(i))
            
                If lngWk > lngMaxWork Then
                    lngMaxWork = lngWk
                End If
                
            Next
            mudtGrid(lngIdxRow, lngIdxCol).TextMaxLength = lngMaxWork
            
'            mudtGrid(lngIdxRow, lngIdxCol).WrapText = Selection(lngIdxRow, lngIdxCol).WrapText
'
'            '横位置が標準以外であればそれにあわせる
'            Select Case Selection(lngIdxRow, lngIdxCol).HorizontalAlignment
'                Case xlGeneral
'                    '書式が文字列なら左寄せ
'                    Select Case True
'                        Case Selection(lngIdxRow, lngIdxCol).NumberFormatLocal = "@"
'                            mudtGrid(lngIdxRow, lngIdxCol).Align = xlLeft
'
'                        Case IsNumeric(Selection(lngIdxRow, lngIdxCol).Value)
'                            '数値の場合、右寄せ
'                            mudtGrid(lngIdxRow, lngIdxCol).Align = xlRight
'
'                        Case IsDate(Selection(lngIdxRow, lngIdxCol).Value)
'                            '日付の場合、右寄せ
'                            mudtGrid(lngIdxRow, lngIdxCol).Align = xlRight
'
'                        Case Else
'                            mudtGrid(lngIdxRow, lngIdxCol).Align = xlLeft
'                    End Select
'
'                Case xlFill, xlJustify
'                    '繰り返し,両端揃え
'                    mudtGrid(lngIdxRow, lngIdxCol).Align = xlLeft
'
'                Case xlCenterAcrossSelection, xlDistributed
'                    '選択範囲内で中央, 均等割り付け
'                    mudtGrid(lngIdxRow, lngIdxCol).Align = xlCenter
'
'                Case Else
'                    mudtGrid(lngIdxRow, lngIdxCol).Align = Selection(lngIdxRow, lngIdxCol).HorizontalAlignment
'            End Select
'
'            '縦位置
'            Select Case Selection(lngIdxRow, lngIdxCol).VerticalAlignment
'                Case xlJustify
'                    mudtGrid(lngIdxRow, lngIdxCol).vAlign = xlTop
'                Case xlDistributed
'                    mudtGrid(lngIdxRow, lngIdxCol).vAlign = xlCenter
'                Case Else
'                    mudtGrid(lngIdxRow, lngIdxCol).vAlign = Selection(lngIdxRow, lngIdxCol).VerticalAlignment
'            End Select

        Next
    Next
    
End Sub
'--------------------------------------------------------------
'　線種（形）の判定
'--------------------------------------------------------------
Private Function getLineData(ByVal lngIdxRow As Long, ByVal lngIdxCol As Long, ByVal lngSquare As Long) As String

    Dim lngResult As Long
    Dim strResult As String

    '以下、コメントに番号があるものは表の場所を表す

    '１│２
    '─┼─
    '３│４


    Select Case lngSquare
        Case C_SQUARE_TOP_LEFT
            '左上

            '左上│上
            '　─┼─
            '　左│○


            If lngIdxRow <> 1 And lngIdxCol <> 1 Then
                With Selection(lngIdxRow, lngIdxCol).Offset(-1, -1)  '選択セルの左上（１）
    
                    If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_LEFT
                            Case Else
                                lngResult = lngResult Or C_BORDER_LEFT_BOLD
                        End Select
                        
                    End If
    
                    If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                    End If
    
                End With
            End If

            If lngIdxRow <> 1 Then
                With Selection(lngIdxRow, lngIdxCol).Offset(-1, 0)   '選択セルの上（２）
    
                    If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                    End If
    
                End With
            End If

            If lngIdxCol <> 1 Then
                With Selection(lngIdxRow, lngIdxCol).Offset(0, -1)   '選択セルの左（３）
    
                    If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_LEFT
                            Case Else
                                lngResult = lngResult Or C_BORDER_LEFT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_BOTTOM
                            Case Else
                                lngResult = lngResult Or C_BORDER_BOTTOM_BOLD
                        End Select
                    End If
    
                End With
            End If

            With Selection(lngIdxRow, lngIdxCol).Offset(0, 0)    '選択セル（４)

                If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                End If

                If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_BOTTOM
                            Case Else
                                lngResult = lngResult Or C_BORDER_BOTTOM_BOLD
                        End Select
                End If

            End With

        Case C_SQUARE_TOP_MIDDLE
            '上中

            '上
            '─
            '○
            If lngIdxRow <> 1 Then
                With Selection(lngIdxRow, lngIdxCol).Offset(-1, 0)   '選択セルの上（２）
    
                    If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                    End If
    
                End With
            End If

            With Selection(lngIdxRow, lngIdxCol).Offset(0, 0)    '選択セル（４)

                If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                End If

            End With

        Case C_SQUARE_TOP_RIGHT
            '右上

            '　上│右上
            '　─┼─
            '　○│右
            If lngIdxRow <> 1 Then
                With Selection(lngIdxRow, lngIdxCol).Offset(-1, 0)   '上（１）
    
                    If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_LEFT
                            Case Else
                                lngResult = lngResult Or C_BORDER_LEFT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                    End If
    
                End With
            End If
            
            If lngIdxRow <> 1 And lngIdxCol <> mlngCol Then
                With Selection(lngIdxRow, lngIdxCol).Offset(-1, 1)  '右上（２）
    
                    If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                    End If
    
                End With
            End If

            With Selection(lngIdxRow, lngIdxCol).Offset(0, 0)    '選択セル（３）

                If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_LEFT
                            Case Else
                                lngResult = lngResult Or C_BORDER_LEFT_BOLD
                        End Select
                End If

                If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_BOTTOM
                            Case Else
                                lngResult = lngResult Or C_BORDER_BOTTOM_BOLD
                        End Select
                End If

            End With

            If lngIdxCol <> mlngCol Then
                With Selection(lngIdxRow, lngIdxCol).Offset(0, 1)    '右（４）
        
                    If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                    End If
        
                    If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_BOTTOM
                            Case Else
                                lngResult = lngResult Or C_BORDER_BOTTOM_BOLD
                        End Select
                    End If
        
                End With
            End If

        Case C_SQUARE_LEFT_MIDDLE

            '左中

            '　左│○
            If lngIdxCol <> 1 Then
                With Selection(lngIdxRow, lngIdxCol).Offset(0, -1)   '左（１）
    
                    If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                    End If
    
                End With
            End If

            With Selection(lngIdxRow, lngIdxCol).Offset(0, 0)    '選択セル（２）

                If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                End If

            End With

        Case C_SQUARE_RIGHT_MIDDLE

            '右中

            '　○│右
            If lngIdxCol <> mlngCol Then
                With Selection(lngIdxRow, lngIdxCol).Offset(0, 1)   '右（１）
    
                    If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                    End If
    
                End With
            End If

            With Selection(lngIdxRow, lngIdxCol).Offset(0, 0)    '選択セル（２）

                If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                End If

            End With

        Case C_SQUARE_BOTTOM_LEFT

            '左下

            '　左│○
            '　─┼─
            '左下│下
            If lngIdxCol <> 1 Then
                With Selection(lngIdxRow, lngIdxCol).Offset(0, -1)   '左（１）
    
                    If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_LEFT
                            Case Else
                                lngResult = lngResult Or C_BORDER_LEFT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                    End If
    
                End With
            End If

            With Selection(lngIdxRow, lngIdxCol).Offset(0, 0)    '選択セル（２）

                If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                End If

                If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                End If

            End With

            If lngIdxRow <> mlngRow And lngIdxCol <> 1 Then
                With Selection(lngIdxRow, lngIdxCol).Offset(1, -1)   '左下（３）
    
                    If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_LEFT
                            Case Else
                                lngResult = lngResult Or C_BORDER_LEFT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_BOTTOM
                            Case Else
                                lngResult = lngResult Or C_BORDER_BOTTOM_BOLD
                        End Select
                    End If
    
                End With
            End If

            If lngIdxRow <> mlngRow Then
                With Selection(lngIdxRow, lngIdxCol).Offset(1, 0)    '下（４）
    
                    If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_BOTTOM
                            Case Else
                                lngResult = lngResult Or C_BORDER_BOTTOM_BOLD
                        End Select
                    End If
    
                End With
            End If

        Case C_SQUARE_BOTTOM_MIDDLE

            '下中

            '○
            '─
            '下
            With Selection(lngIdxRow, lngIdxCol).Offset(0, 0)    '選択セル（２）

                If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                End If

            End With

            If lngIdxRow <> mlngRow Then
                With Selection(lngIdxRow, lngIdxCol).Offset(1, 0)    '下（４）
    
                    If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                    End If
    
                End With
            End If

        Case C_SQUARE_BOTTOM_RIGHT
            '右下

            '　○│右
            '　─┼─
            '　下│右下
            With Selection(lngIdxRow, lngIdxCol).Offset(0, 0)    '選択セル（１）

                If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_LEFT
                            Case Else
                                lngResult = lngResult Or C_BORDER_LEFT_BOLD
                        End Select
                End If

                If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                End If

            End With

            If lngIdxCol <> mlngCol Then
                With Selection(lngIdxRow, lngIdxCol).Offset(0, 1)   '右（２）
    
                    If .Borders(xlEdgeBottom).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeBottom).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_TOP
                            Case Else
                                lngResult = lngResult Or C_BORDER_TOP_BOLD
                        End Select
                    End If
    
                End With
            End If

            If lngIdxRow <> mlngRow Then
                With Selection(lngIdxRow, lngIdxCol).Offset(1, 0)    '下（３）
    
                    If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_LEFT
                            Case Else
                                lngResult = lngResult Or C_BORDER_LEFT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeRight).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeRight).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_BOTTOM
                            Case Else
                                lngResult = lngResult Or C_BORDER_BOTTOM_BOLD
                        End Select
                    End If
    
                End With
            End If

            If lngIdxRow <> mlngRow And lngIdxCol <> mlngCol Then
                With Selection(lngIdxRow, lngIdxCol).Offset(1, 1)   '右下(４)
    
                    If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeTop).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_RIGHT
                            Case Else
                                lngResult = lngResult Or C_BORDER_RIGHT_BOLD
                        End Select
                    End If
    
                    If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Select Case .Borders(xlEdgeLeft).Weight
                            Case mlngBorderWeight1, mlngBorderWeight2
                                lngResult = lngResult Or C_BORDER_BOTTOM
                            Case Else
                                lngResult = lngResult Or C_BORDER_BOTTOM_BOLD
                        End Select
                    End If
    
                End With
            End If
    End Select

    Select Case lngResult
        Case C_BORDER_LR, C_BORDER_LEFT, C_BORDER_RIGHT
            strResult = "─"

        Case C_BORDER_TB, C_BORDER_TOP, C_BORDER_BOTTOM
            strResult = "│"

        Case C_BORDER_TL
            strResult = "┘"

        Case C_BORDER_TR
            strResult = "└"

        Case C_BORDER_BR
            strResult = "┌"

        Case C_BORDER_BL
            strResult = "┐"

        Case C_BORDER_TBR
            strResult = "├"

        Case C_BORDER_BLT
            strResult = "┬"

        Case C_BORDER_TBL
            strResult = "┤"

        Case C_BORDER_TLR
            strResult = "┴"

        Case C_BORDER_CROSS
            strResult = "┼"
            
            
        Case C_BORDER_LR_BOLD, C_BORDER_LEFT_BOLD, C_BORDER_RIGHT_BOLD
            strResult = "━"

        Case C_BORDER_TB_BOLD, C_BORDER_TOP_BOLD, C_BORDER_BOTTOM_BOLD
            strResult = "┃"

        Case C_BORDER_TL_BOLD, C_BORDER_TL_BH, C_BORDER_TL_HB
            strResult = "┛"

        Case C_BORDER_TR_BOLD, C_BORDER_TR_BH, C_BORDER_TR_HB
            strResult = "┗"

        Case C_BORDER_BR_BOLD, C_BORDER_BR_BH, C_BORDER_BR_HB
            strResult = "┏"

        Case C_BORDER_BL_BOLD, C_BORDER_BL_BH, C_BORDER_BL_HB
            strResult = "┓"

        Case C_BORDER_TBR_BOLD
            strResult = "┣"

        Case C_BORDER_BLT_BOLD
            strResult = "┳"

        Case C_BORDER_TBL_BOLD
            strResult = "┫"

        Case C_BORDER_TLR_BOLD, C_BORDER_TLR_TLBOLD
            strResult = "┻"

        Case C_BORDER_CROSS_BOLD, C_BORDER_CROSS_BOLD_UL, C_BORDER_CROSS_RB, C_BORDER_CROSS_UB, C_BORDER_CROSS_LB
            strResult = "╋"
            
        
        
        Case C_BORDER_TBR_BH
            strResult = "┠"

        Case C_BORDER_BLT_BH
            strResult = "┯"

        Case C_BORDER_TBL_BH
            strResult = "┨"

        Case C_BORDER_TLR_BH
            strResult = "┷"

        Case C_BORDER_CROSS_BH
            strResult = "┿"
            
        
        
        Case C_BORDER_TBR_HB
            strResult = "┝"

        Case C_BORDER_BLT_HB
            strResult = "┰"

        Case C_BORDER_TBL_HB
            strResult = "┥"

        Case C_BORDER_TLR_HB
            strResult = "┸"

        Case C_BORDER_CROSS_HB
            strResult = "╂"
            
        
        
        Case C_BORDER_NONE
            strResult = "  "
            
        Case Else
            strResult = "  "
    
    End Select

    getLineData = strResult

End Function
'--------------------------------------------------------------
'　線種（太さ）の判定
'--------------------------------------------------------------
Private Sub checkWeight()

    Dim lngHair As Long
    Dim lngThin As Long
    Dim lngMedium As Long
    Dim lngThick As Long

    Dim r As Range
    
    
    For Each r In Selection

        Select Case r.Borders(xlEdgeTop).Weight
            Case xlHairline
                lngHair = lngHair + 1
            Case xlThin
                lngThin = lngThin + 1
            Case xlMedium
                lngMedium = lngMedium + 1
            Case xlThick
                lngThick = lngThick + 1
        End Select
        
        Select Case r.Borders(xlEdgeBottom).Weight
            Case xlHairline
                lngHair = lngHair + 1
            Case xlThin
                lngThin = lngThin + 1
            Case xlMedium
                lngMedium = lngMedium + 1
            Case xlThick
                lngThick = lngThick + 1
        End Select
        
        Select Case r.Borders(xlEdgeLeft).Weight
            Case xlHairline
                lngHair = lngHair + 1
            Case xlThin
                lngThin = lngThin + 1
            Case xlMedium
                lngMedium = lngMedium + 1
            Case xlThick
                lngThick = lngThick + 1
        End Select
        
        Select Case r.Borders(xlEdgeRight).Weight
            Case xlHairline
                lngHair = lngHair + 1
            Case xlThin
                lngThin = lngThin + 1
            Case xlMedium
                lngMedium = lngMedium + 1
            Case xlThick
                lngThick = lngThick + 1
        End Select
        
    Next

    Select Case True
        Case lngHair > 0 Or lngThin = 0 And lngMedium = 0 And lngThick = 0
            '1
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlHairline
        
        Case lngHair = 0 Or lngThin > 0 And lngMedium = 0 And lngThick = 0
            '2
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair > 0 Or lngThin > 0 And lngMedium = 0 And lngThick = 0
            '3
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair = 0 Or lngThin = 0 And lngMedium > 0 And lngThick = 0
            '4
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair > 0 Or lngThin = 0 And lngMedium > 0 And lngThick = 0
            '5
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair = 0 Or lngThin > 0 And lngMedium > 0 And lngThick = 0
            '6
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair > 0 Or lngThin > 0 And lngMedium > 0 And lngThick = 0
            '7
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair = 0 Or lngThin = 0 And lngMedium = 0 And lngThick > 0
            '8
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair > 0 Or lngThin = 0 And lngMedium = 0 And lngThick > 0
            '9
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair = 0 Or lngThin > 0 And lngMedium = 0 And lngThick > 0
            '10
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair > 0 Or lngThin > 0 And lngMedium = 0 And lngThick > 0
            '11
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair = 0 Or lngThin = 0 And lngMedium > 0 And lngThick > 0
            '12
            mlngBorderWeight1 = xlMedium
            mlngBorderWeight2 = xlThick
        
        Case lngHair > 0 Or lngThin = 0 And lngMedium > 0 And lngThick > 0
            '13
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair = 0 Or lngThin > 0 And lngMedium > 0 And lngThick > 0
            '14
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case lngHair > 0 Or lngThin > 0 And lngMedium > 0 And lngThick > 0
            '15
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlThin
        
        Case Else
            'Else
            mlngBorderWeight1 = xlHairline
            mlngBorderWeight2 = xlHairline

    End Select

End Sub
'--------------------------------------------------------------
'　結合セルの内側の線をクリアする。
'--------------------------------------------------------------
Private Sub clearInsideLine()

    Dim r As Range
    Dim strAddress As String
    Dim c As Collection
    
    Set c = New Collection

    For Each r In Selection
    
        If r.MergeCells Then
            On Error Resume Next
            strAddress = ""
            strAddress = c(r.MergeArea(1, 1).Address)
            On Error GoTo 0
            
            If strAddress = "" Then
                r.MergeArea.Borders(xlInsideVertical).LineStyle = xlNone
                r.MergeArea.Borders(xlInsideHorizontal).LineStyle = xlNone
                c.Add r.MergeArea(1, 1).Address, r.MergeArea(1, 1).Address
            End If
        End If
    Next

End Sub


Private Sub chkKetaEnabled_Click()

    If chkKetaEnabled.Value = True Then
        txtKeta.enabled = True
        spnKeta.enabled = True
        chkDate.enabled = True
        chkNum.enabled = True
    Else
        txtKeta.enabled = False
        spnKeta.enabled = False
        chkDate.enabled = False
        chkNum.enabled = False
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()

    Dim lngKeta As Long
    Dim lngMin As Long

    If chkKetaEnabled.Value Then
    
        lngKeta = Val(txtKeta.Text)
        mlngMinKeta = (Selection.Columns.count + 1) * 2 + (Selection.Columns.count * 2)
    
        If mlngMinKeta > lngKeta Then
            MsgBox CStr(mlngMinKeta) & "桁以下には指定できません。" & "これ以上表を小さくする場合は選択列を減らしてください。", vbExclamation, C_TITLE
            Exit Sub
        End If
    
        If lngKeta Mod 2 = 1 Then
            MsgBox "桁には偶数を指定してください。", vbExclamation, C_TITLE
            Exit Sub
        End If
    
        mlngMaxKeta = txtKeta.Text
    Else
        mlngMaxKeta = C_LINE_MAX
    End If
    
    Call kantanLineRun
    
    
    Call SaveSetting(C_TITLE, "EasyLine", "chkKetaEnabled", CStr(chkKetaEnabled.Value))
    Call SaveSetting(C_TITLE, "EasyLine", "txtKeta", CStr(txtKeta.Text))
    Call SaveSetting(C_TITLE, "EasyLine", "chkDate", CStr(chkDate.Value))
    Call SaveSetting(C_TITLE, "EasyLine", "chkNum", CStr(chkNum.Value))
    
    
End Sub


Private Sub spnKeta_SpinDown()
    txtKeta.Text = spinDown(txtKeta.Text)
End Sub

Private Sub spnKeta_SpinUp()
    txtKeta.Text = spinUp(txtKeta.Text)
End Sub

Private Sub txtPreview_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = txtPreview
End Sub

Private Sub UserForm_Initialize()

    chkKetaEnabled.Value = CBool(GetSetting(C_TITLE, "EasyLine", "chkKetaEnabled", "False"))
    txtKeta.Text = CLng(GetSetting(C_TITLE, "EasyLine", "txtKeta", CStr(C_LINE_MAX)))
    chkDate.Value = CBool(GetSetting(C_TITLE, "EasyLine", "chkDate", "False"))
    chkNum.Value = CBool(GetSetting(C_TITLE, "EasyLine", "chkNum", "False"))

    mlngMinKeta = (Selection.Columns.count + 1) * 2 + (Selection.Columns.count * 2)
    
    If chkKetaEnabled.Value Then
        mlngMaxKeta = CLng(txtKeta.Text)
    Else
        mlngMaxKeta = C_LINE_MAX
    End If

    Call chkKetaEnabled_Click

    Call kantanLineRun
    
    Set MW = basMouseWheel.GetInstance
    MW.Install
    
End Sub
Private Sub UserForm_Activate()
    txtPreview.SelStart = 0
    txtPreview.SelLength = Len(txtPreview.Text)
End Sub
Private Function spinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 2
    spinUp = lngValue

End Function

Private Function spinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    mlngMinKeta = (Selection.Columns.count + 1) * 2 + (Selection.Columns.count * 2)

    lngValue = Val(vntValue)
    lngValue = lngValue - 2
    If lngValue < mlngMinKeta Then
        lngValue = mlngMinKeta
    End If
    spinDown = lngValue

End Function

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = Nothing
End Sub

Private Sub UserForm_Terminate()
    MW.UnInstall
    Set MW = Nothing
End Sub
Private Sub MW_WheelDown(obj As Object)
    
    Dim lngPos As Long
    
    On Error GoTo e
    
    lngPos = obj.CurLine + 3
    If lngPos >= obj.LineCount Then
        lngPos = obj.LineCount - 1
    End If
    obj.CurLine = lngPos

e:
End Sub

Private Sub MW_WheelUp(obj As Object)

    Dim lngPos As Long
    
    On Error GoTo e
    
    lngPos = obj.CurLine - 3
    If lngPos < 0 Then
        lngPos = 0
    End If
    obj.CurLine = lngPos
e:
End Sub
