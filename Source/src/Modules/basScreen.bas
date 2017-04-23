Attribute VB_Name = "basScreen"
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


#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbsize As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

    Private Type MOUSEINPUT
        dx As Long
        dy As Long
        mouseData As Long
        dwFlags As Long
        time As LongLong
        dwExtraInfo As LongPtr
    End Type
    
    Private Type INPUT_TYPE
        dwType As Long
        dummy As Long
        mi As MOUSEINPUT
    End Type

#Else
    Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Integer) As Integer
    Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbsize As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    
    Private Type MOUSEINPUT
        dx As Long
        dy As Long
        mouseData As Long
        dwFlags As Long
        time As Long
        dwExtraInfo As Long
    End Type
    
    Private Type INPUT_TYPE
        dwType As Long
        mi As MOUSEINPUT
    End Type
    
#End If
    
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const INPUT_MOUSE As Long = 0
Private Const MOUSE_MOVED As Long = &H1               'マウスを移動する
Private Const MOUSEEVENTF_ABSOLUTE As Long = &H8000&  '移動時、絶対座標を指定
Private Const SM_CXSCREEN = 0    'ディスプレイの幅
Private Const SM_CYSCREEN = 1    'ディスプレイの高さ
Private Const DPI As Long = 96
Private Const PTUNIT As Single = 0.75  'エクセル上のポイント値は0.75の倍数
Private Const MOUSEEVENTF_LEFTDOWN As Integer = &H2      '左ボタンDown

Function PointToPixelX(pt As Double) As Long
    PointToPixelX = Int(pt * DPI / PPI)
End Function
Function PointToPixelY(pt As Double) As Long
    PointToPixelY = Int(pt * DPI / PPI)
End Function
Private Function PPI() As Double
    PPI = Application.InchesToPoints(1)
End Function
Sub PickShape(ByRef objDataSet As Object)

    Dim X As Long
    Dim Y As Long
    Dim ax As Long
    Dim ay As Long

    Dim bx As Long
    Dim by As Long
    Dim cx As Long
    Dim cy As Long
    Dim dx As Long
    Dim dy As Long
    Dim spx As Long
    Dim spy As Long

    Dim a As POINTAPI

    Dim dummy As Long
    bx = ActiveWindow.ActivePane.ScrollColumn
    by = ActiveWindow.ActivePane.ScrollRow

    ActiveWindow.ActivePane.ScrollColumn = ActiveWindow.SplitColumn + 1
    
    'Excel 2010 以前は True にして　PointsToScreenPixelsX/Y が動作するようにする
    Application.ScreenUpdating = (Val(Application.Version) <= C_EXCEL_VERSION_2010)
    
    dummy = ActiveWindow.SplitHorizontal    'PointsToScreenPixelsXの値を更新するために使用
    spx = ActiveWindow.ActivePane.PointsToScreenPixelsX(0)
    ActiveWindow.ActivePane.ScrollColumn = bx

    ActiveWindow.ActivePane.ScrollRow = ActiveWindow.SplitRow + 1
    
    dummy = ActiveWindow.SplitVertical    'PointsToScreenPixelsXの値を更新するために使用
    spy = ActiveWindow.ActivePane.PointsToScreenPixelsY(0)
    ActiveWindow.ActivePane.ScrollRow = by
    
    
    Dim dblZoomX As Double
    Dim dblZoomY As Double

    '現在のカーソル位置のスクリーン座標を取得
    GetCursorPos a
    
    If ActiveWindow.RangeFromPoint(a.X, a.Y) Is Nothing Then
        dx = 100
        dy = 100
    Else
        dx = ActiveWindow.RangeFromPoint(a.X, a.Y).Left + (ActiveWindow.RangeFromPoint(a.X, a.Y).width) / 2
        dy = ActiveWindow.RangeFromPoint(a.X, a.Y).Top + (ActiveWindow.RangeFromPoint(a.X, a.Y).Height) / 2
    End If
    
    If ActiveWindow.Zoom = 100 Then
        dblZoomX = 1
        dblZoomY = 1
        
        '100%時の分割サイズを保存
        cx = ActiveWindow.SplitHorizontal
        cy = ActiveWindow.SplitVertical
    Else
    
        'この処理の趣旨 現在のマウスポインタ位置の指定％のピクセル数と１００％時のピクセル数から本来の倍率を求める
        Dim lngZoom As Long
        
        Dim lngToX1 As Long
        Dim lngToY1 As Long
        Dim lngToX2 As Long
        Dim lngToY2 As Long

        lngZoom = ActiveWindow.Zoom

        dummy = ActiveWindow.SplitHorizontal    'PointsToScreenPixelsXの値を更新するために使用
        lngToX1 = ActiveWindow.ActivePane.PointsToScreenPixelsX(0)
        
        dummy = ActiveWindow.SplitVertical    'PointsToScreenPixelsYの値を更新するために使用
        lngToY1 = ActiveWindow.ActivePane.PointsToScreenPixelsY(0)

        dummy = ActiveWindow.SplitHorizontal    'PointsToScreenPixelsXの値を更新するために使用
        lngToX2 = ActiveWindow.ActivePane.PointsToScreenPixelsX(dx) - lngToX1
        
        dummy = ActiveWindow.SplitVertical    'PointsToScreenPixelsYの値を更新するために使用
        lngToY2 = ActiveWindow.ActivePane.PointsToScreenPixelsY(dy) - lngToY1

        Dim lngFromX2 As Long
        Dim lngFromY2 As Long
        
        lngFromX2 = Round(dx * DPI / PPI)
        lngFromY2 = Round(dy * DPI / PPI)

        '倍率を計算
        dblZoomX = CDbl(lngFromX2) / lngToX2
        dblZoomY = CDbl(lngFromY2) / lngToY2
        
        Application.ScreenUpdating = False
        
        '100%時の分割サイズを保存
        ' ズーム１００に対する割合を取得（ポイント）
        ActiveWindow.Zoom = 100
        
        cx = ActiveWindow.SplitHorizontal
        cy = ActiveWindow.SplitVertical
        
        ActiveWindow.Zoom = lngZoom

    End If

    'マウスポインタの座標(ピクセル)をポイントに変換
    ax = Round(((a.X - spx) * PPI / DPI) * dblZoomX)
    X = ActiveSheet.Cells(by, bx).Left - cx + ax

    ay = Round(((a.Y - spy) * PPI / DPI) * dblZoomY)
    Y = ActiveSheet.Cells(by, bx).Top - cy + ay
    
    Dim r As Range
    Set r = ActiveWindow.ActivePane.VisibleRange
    
    'マウスカーソルが作業ウィンドウ内にある場合
    If r(1).Top < Y And r(1).Left < X And r(r.count).Top + r(r.count).Height > Y + objDataSet.Height And r(r.count).Left + r(r.count).width > X + objDataSet.width Then
    
        'シェイプをマウスカーソル位置に移動する
        objDataSet.Left = X - (objDataSet.width / 2)
        objDataSet.Top = Y - (objDataSet.Height / 2)
        
        'シェイプを選択
        Call SetCursoleAndLeftDown(a.X, a.Y)
    
    'マウスカーソルが作業ウィンドウ外にある場合
    Else
        'カーソルをシェイプに移動する
        ax = ((objDataSet.Left + (objDataSet.width / 2) - ActiveSheet.Cells(by, bx).Left + cx) * DPI / PPI)
        X = Round(spx + ax / dblZoomX)
        
        ay = ((objDataSet.Top + (objDataSet.Height / 2) - ActiveSheet.Cells(by, bx).Top + cy) * DPI / PPI)
        Y = Round(spy + ay / dblZoomY)
        
        'シェイプを選択
        Call SetCursoleAndLeftDown(X, Y)
    
    End If

    Application.ScreenUpdating = True

End Sub
Private Sub SetCursoleAndLeftDown(ByVal X As Long, ByVal Y As Long)

    Dim inp(0 To 2) As INPUT_TYPE
    
    Dim a As POINTAPI

    GetCursorPos a
    
    With inp(0)
        .dwType = INPUT_MOUSE
        .mi.dx = (X * 65535 / (GetSystemMetrics(SM_CXSCREEN) - 1))
        .mi.dy = (Y * 65535 / (GetSystemMetrics(SM_CYSCREEN) - 1))
        .mi.mouseData = 0
        .mi.dwFlags = MOUSE_MOVED Or MOUSEEVENTF_ABSOLUTE
        .mi.time = 0
        .mi.dwExtraInfo = 0
    End With
    
    With inp(1)
        .dwType = INPUT_MOUSE
        .mi.dx = 0
        .mi.dy = 0
        .mi.mouseData = 0
        .mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        .mi.time = 0
        .mi.dwExtraInfo = 0
    End With
    
    SendInput 2, inp(0), LenB(inp(0))

End Sub
