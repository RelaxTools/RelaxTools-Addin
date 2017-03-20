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
        dX As Long
        dY As Long
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
        dX As Long
        dY As Long
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
    Dim spx As Long
    Dim spy As Long

    Application.ScreenUpdating = False

    Dim a As POINTAPI

    Dim dummy As Long
    bx = ActiveWindow.ActivePane.ScrollColumn
    by = ActiveWindow.ActivePane.ScrollRow

    ActiveWindow.ActivePane.ScrollColumn = 1
    
    dummy = ActiveWindow.SplitHorizontal
    spx = ActiveWindow.ActivePane.PointsToScreenPixelsX(0)
    ActiveWindow.ActivePane.ScrollColumn = bx

    ActiveWindow.ActivePane.ScrollRow = 1
    
    dummy = ActiveWindow.SplitHorizontal
    spy = ActiveWindow.ActivePane.PointsToScreenPixelsY(0)
    ActiveWindow.ActivePane.ScrollRow = by
    
    Dim dblZoomX As Double
    Dim dblZoomY As Double

    If ActiveWindow.Zoom = 100 Then
        dblZoomX = 1
        dblZoomY = 1
    Else
        Dim lngZoom As Long
        Dim lngUsableWidth As Long
        Dim lngUsableHeight As Long

        lngZoom = ActiveWindow.Zoom

        lngUsableWidth = ActiveWindow.ActivePane.PointsToScreenPixelsX(Cells(2, 2).Left) - ActiveWindow.ActivePane.PointsToScreenPixelsX(0)
        lngUsableHeight = ActiveWindow.ActivePane.PointsToScreenPixelsY(Cells(2, 2).Top) - ActiveWindow.ActivePane.PointsToScreenPixelsY(0)

        ' ズーム１００に対する割合を取得（ポイント）
        ActiveWindow.Zoom = 100

        dblZoomX = (ActiveWindow.ActivePane.PointsToScreenPixelsX(Cells(2, 2).Left) - ActiveWindow.ActivePane.PointsToScreenPixelsX(0)) / lngUsableWidth
        dblZoomY = (ActiveWindow.ActivePane.PointsToScreenPixelsY(Cells(2, 2).Top) - ActiveWindow.ActivePane.PointsToScreenPixelsY(0)) / lngUsableHeight

        ActiveWindow.Zoom = lngZoom
    End If

    '現在のカーソル位置のスクリーン座標を取得
    GetCursorPos a

    ax = ((a.X - spx) * PPI / DPI)
    X = ActiveSheet.Cells(by, bx).Left + ax * dblZoomX


    ay = ((a.Y - spy) * PPI / DPI)
    Y = ActiveSheet.Cells(by, bx).Top + ay * dblZoomY
    
    Dim r As Range
    Set r = ActiveWindow.ActivePane.VisibleRange
    
    'マウスカーソルが作業ウィンドウ内にある場合
    If r(1).Top < Y And r(1).Left < X And r(r.count).Top + r(r.count).Height > Y + objDataSet.Height And r(r.count).Left + r(r.count).Width > X + objDataSet.Width Then
    
        'シェイプをマウスカーソル位置に移動する
        objDataSet.Left = X - (objDataSet.Width / 2)
        objDataSet.Top = Y - (objDataSet.Height / 2)
        
        Call SetCursoleAndLeftDown(a.X, a.Y)
    
    'マウスカーソルが作業ウィンドウ外にある場合
    Else
        'カーソルをシェイプに移動する
        ax = ((objDataSet.Left + (objDataSet.Width / 2) - ActiveSheet.Cells(by, bx).Left) * DPI / PPI)
        X = spx + ax / dblZoomX
        
        ay = ((objDataSet.Top + (objDataSet.Height / 2) - ActiveSheet.Cells(by, bx).Top) * DPI / PPI)
        Y = spy + ay / dblZoomY
        
        Call SetCursoleAndLeftDown(X, Y)
    
    End If

    Application.ScreenUpdating = True

End Sub
Sub SetCursoleAndLeftDown(ByVal X As Long, ByVal Y As Long)

    Dim inp(0 To 2) As INPUT_TYPE
    
    Dim a As POINTAPI

    GetCursorPos a
    
    With inp(0)
        .dwType = INPUT_MOUSE
        .mi.dX = (X * 65535 / (GetSystemMetrics(SM_CXSCREEN) - 1))
        .mi.dY = (Y * 65535 / (GetSystemMetrics(SM_CYSCREEN) - 1))
        .mi.mouseData = 0
        .mi.dwFlags = MOUSE_MOVED Or MOUSEEVENTF_ABSOLUTE
        .mi.time = 0
        .mi.dwExtraInfo = 0
    End With
    
    With inp(1)
        .dwType = INPUT_MOUSE
        .mi.dX = 0
        .mi.dY = 0
        .mi.mouseData = 0
        .mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        .mi.time = 0
        .mi.dwExtraInfo = 0
    End With
    
'    With inp(2)
'        .dwType = INPUT_MOUSE
'        .mi.dX = (X2 * 65535 / (GetSystemMetrics(SM_CXSCREEN) - 1))
'        .mi.dY = (Y2 * 65535 / (GetSystemMetrics(SM_CYSCREEN) - 1))
'        .mi.mouseData = 0
'        .mi.dwFlags = MOUSE_MOVED Or MOUSEEVENTF_ABSOLUTE
'        .mi.time = 0
'        .mi.dwExtraInfo = 0
'    End With
    
    
    SendInput 2, inp(0), LenB(inp(0))

End Sub



'Sub SetCursoleAndLeftDownCustom(ByVal x As Long, ByVal y As Long)
'
'    Dim inp(0 To 2) As INPUT_TYPE
'
'    Dim a As POINTAPI
'
'    GetCursorPos a
'
'    With inp(0)
'        .dwType = INPUT_MOUSE
'        .mi.dX = 0
'        .mi.dY = 0
'        .mi.mouseData = 0
'        .mi.dwFlags = MOUSEEVENTF_LEFTDOWN
'        .mi.time = 0
'        .mi.dwExtraInfo = 0
'    End With
'
'    SendInput 1, inp(0), Len(inp(0))
'
'End Sub



'Sub PickShapeO(ByRef objDataSet As Object)
'
'    Dim x As Long
'    Dim y As Long
'    Dim ax As Long
'    Dim ay As Long
'
'    Dim bx As Long
'    Dim by As Long
'    Dim spx As Long
'    Dim spy As Long
'
'    Application.ScreenUpdating = False
'
'
'    bx = ActiveWindow.ActivePane.ScrollColumn
'    by = ActiveWindow.ActivePane.ScrollRow
'
'    ActiveWindow.ActivePane.ScrollColumn = 1
'    spx = ActiveWindow.ActivePane.PointsToScreenPixelsX(0)
'    ActiveWindow.ActivePane.ScrollColumn = bx
'
'    ActiveWindow.ActivePane.ScrollRow = 1
'    spy = ActiveWindow.ActivePane.PointsToScreenPixelsY(0)
'    ActiveWindow.ActivePane.ScrollRow = by
'
'    ax = ((objDataSet.Left + (objDataSet.Width / 2) - Cells(by, bx).Left) * DPI / PPI)
'
''    x = spx + ax * (ActiveWindow.Zoom / 100)
'    x = spx + ax
'    'x = (ActiveWindow.ActivePane.PointsToScreenPixelsX((objDataSet.Left + (objDataSet.Width / 2)) * DPI / PPI) - ActiveWindow.ActivePane.PointsToScreenPixelsX(0)) * (ActiveWindow.Zoom / 100)
'
'
'
'    ay = ((objDataSet.Top + (objDataSet.Height / 4) - Cells(by, bx).Top) * DPI / PPI)
'
''    y = spy + ay * (ActiveWindow.Zoom / 100)
'    y = spy + ay
'
''    y = ActiveWindow.ActivePane.PointsToScreenPixelsY(0) + (ActiveWindow.ActivePane.PointsToScreenPixelsY(0) - spy + ay) * (ActiveWindow.Zoom / 100)  '+ (ActiveWindow.Zoom - 100) * ay
''    y = ActiveWindow.ActivePane.PointsToScreenPixelsY(0) + Fix(((objDataSet.Top + (objDataSet.Height / 2)) * DPI / PPI) * (ActiveWindow.Zoom / 100))
'    'y = (ActiveWindow.ActivePane.PointsToScreenPixelsY((objDataSet.Top + (objDataSet.Height / 2)) * DPI / PPI) - ActiveWindow.ActivePane.PointsToScreenPixelsY(0)) * (ActiveWindow.Zoom / 100)
'
'
'
''    x = ActiveWindow.ActivePane.PointsToScreenPixelsX(0) + ((objDataSet.Left + (objDataSet.Width / 2)) * DPI / PPI) * (ActiveWindow.Zoom / 100)
''    y = ActiveWindow.ActivePane.PointsToScreenPixelsY(0) + ((objDataSet.Top + (objDataSet.Height / 2)) * DPI / PPI) * (ActiveWindow.Zoom / 100)
'
''    Call SetCursorPos(x, y)
''
''    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
'
'    Call SetCursoleAndLeftDown(x, y)
'
'    Application.ScreenUpdating = True
'
'
'End Sub



