VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeFramework"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'--------------------------------------------------------------
' ShapeFramework 1.00 2017/03/29 Y.Watanabe
'--------------------------------------------------------------
' マウスカーソルにシェイプを添わせるフレームワーク
'--------------------------------------------------------------
Option Explicit

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbsize As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDc As LongPtr) As Long

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
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Sub ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long)
    
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

Private Type ZOOMVALUE
    X As Double
    Y As Double
End Type

Private Type TABLE
    Row As Long
    Col As Long
End Type
Private Const LOGPIXELSX As Long = &H58&
Private Const LOGPIXELSY As Long = &H5A&

Private Const INPUT_MOUSE As Long = 0
Private Const MOUSE_MOVED As Long = &H1               'マウスを移動する
Private Const MOUSEEVENTF_ABSOLUTE As Long = &H8000&  '移動時、絶対座標を指定
Private Const SM_CXSCREEN = 0    'ディスプレイの幅
Private Const SM_CYSCREEN = 1    'ディスプレイの高さ
Private Const PTUNIT As Single = 0.75  'エクセル上のポイント値は0.75の倍数
Private Const MOUSEEVENTF_LEFTDOWN As Integer = &H2      '左ボタンDown

Private WD As Window
Private WS As Worksheet
Private dblSx As Double
Private dblSy As Double

'初期化イベント
Public Event SelectionInit(ByRef Cancel As Boolean)
'メインイベント
Public Event SelectionMain(ByRef obj As Object, ByRef pos As ClickPosition, ByRef margin As Long, ByRef Cancel As Boolean)
'終了イベント
Public Event SelectionTerm()

Enum ClickPosition
    TopLeft
    TopCenter
    TopRight
    Middleleft
    MiddleCenter
    MiddleRight
    BottomLeft
    BottomCenter
    BottomRight
End Enum

Public Sub Run()

    Dim objDataSet As Object
    Dim Cancel As Boolean
    Dim pos As ClickPosition
    Dim margin As Long
    Dim sh As Object
    Const C_MARGIN As Long = 1
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    If ActiveWindow Is Nothing Then
        MsgBox "アクティブなウィンドウが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    Set WD = ActiveWindow
    Set WS = WD.ActiveSheet
    
    'DPI取得
    Call GetDPI
    
    Cancel = False
    pos = MiddleCenter
    margin = C_MARGIN
    
    '2010以前はTrueそれ以降はFalse
    Application.ScreenUpdating = (Val(Application.Version) <= C_EXCEL_VERSION_2010)
    
    '---------------------------
    '初期化イベント
    '---------------------------
    RaiseEvent SelectionInit(Cancel)

    'キャンセルの場合
    If Cancel Then
        Exit Sub
    End If
    
    '表示シートの左上のスクリーン座標を取得
    Dim org_px As POINTAPI
    org_px = GetOriginPixel()

    '現在のカーソル位置のスクリーン座標を取得
    Dim cr_px As POINTAPI
    Dim zm As ZOOMVALUE
    
    GetCursorPos cr_px
    
    If WD.Zoom = 100 Then
        zm.X = 1
        zm.Y = 1
    Else
        '現在のカーソル位置のセルまたはシェイプから大体の位置を割り出す。
        Dim sd_pt As POINTAPI
        Dim c As POINTAPI

        If WD.RangeFromPoint(cr_px.X, cr_px.Y) Is Nothing Then
            sd_pt.X = 100
            sd_pt.Y = 100
        Else
            'カーソル下のオブジェクト(セルかシェイプ）の中心の位置を割り出す
            sd_pt.X = WD.RangeFromPoint(cr_px.X, cr_px.Y).Left + (WD.RangeFromPoint(cr_px.X, cr_px.Y).Width) / 2
            sd_pt.Y = WD.RangeFromPoint(cr_px.X, cr_px.Y).Top + (WD.RangeFromPoint(cr_px.X, cr_px.Y).Height) / 2
        End If
        'シード値から真のズームを求める
        zm = GetTrueZoom(sd_pt)
    End If
    
    Application.ScreenUpdating = False
    
    '---------------------------
    'メインイベント
    '---------------------------
    RaiseEvent SelectionMain(objDataSet, pos, margin, Cancel)
    If Cancel Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    'マウスポインタの座標(ピクセル)をポイントに変換
    Dim wk_px As POINTAPI
    Dim wk_pt As POINTAPI
    
    wk_px.X = cr_px.X - org_px.X
    wk_px.Y = cr_px.Y - org_px.Y
    
    wk_pt = PixelToPointZoom(wk_px, zm)
    
    '表示シートの左上を取得
    Dim tb As TABLE
    tb.Col = WD.ActivePane.ScrollColumn
    tb.Row = WD.ActivePane.ScrollRow
    
    Dim cr_pt As POINTAPI
    cr_pt.X = WS.Cells(tb.Row, tb.Col).Left + wk_pt.X
    cr_pt.Y = WS.Cells(tb.Row, tb.Col).Top + wk_pt.Y
    
    'マウスカーソルが作業ウィンドウ内にある場合
    Dim r As Range
    Set r = WD.ActivePane.VisibleRange
    If r(1).Top < cr_pt.Y And r(1).Left < cr_pt.X And r(r.count).Offset(-1, -1).Top + r(r.count).Offset(-1, -1).Height > cr_pt.Y And r(r.count).Offset(-1, -1).Left + r(r.count).Offset(-1, -1).Width > cr_pt.X Then
    
        'シェイプをマウスカーソル位置に移動する。ただし、かならずカーソルの中心に移動できるわけではないため、下でシェイプにあわせてカーソルを移動
        objDataSet.Top = cr_pt.Y - (objDataSet.Height / 2)
        objDataSet.Left = cr_pt.X - (objDataSet.Width / 2)
    Else
    
        objDataSet.Top = cr_pt.Y - (objDataSet.Height / 2)
        objDataSet.Left = cr_pt.X - (objDataSet.Width / 2)
    
        If cr_pt.Y < r.Top Then
            objDataSet.Top = r.Top
        End If
        
        If cr_pt.X < r.Left Then
            objDataSet.Left = r.Left
        End If
        
        If cr_pt.Y > r(r.count).Offset(-1, -1).Top + r(r.count).Offset(-1, -1).Height - objDataSet.Height Then
            objDataSet.Top = r(r.count).Offset(-1, -1).Top + r(r.count).Offset(-1, -1).Height - objDataSet.Height
        End If
        
        If cr_pt.X > r(r.count).Offset(-1, -1).Left + r(r.count).Offset(-1, -1).Width - objDataSet.Width Then
            objDataSet.Left = r(r.count).Offset(-1, -1).Left + r(r.count).Offset(-1, -1).Width - objDataSet.Width
        End If

    End If
    
    'カーソルをシェイプに移動する
    Dim dx_pt As POINTAPI
        
    Select Case pos
        Case TopLeft
            dx_pt.X = objDataSet.Left + margin
            dx_pt.Y = objDataSet.Top + margin
        Case TopCenter
            dx_pt.X = objDataSet.Left + (objDataSet.Width / 2)
            dx_pt.Y = objDataSet.Top + margin
        Case TopRight
            dx_pt.X = objDataSet.Left + objDataSet.Width - margin
            dx_pt.Y = objDataSet.Top + margin
        Case Middleleft
            dx_pt.X = objDataSet.Left + margin
            dx_pt.Y = objDataSet.Top + (objDataSet.Height / 2)
        Case MiddleCenter
            dx_pt.X = objDataSet.Left + (objDataSet.Width / 2)
            dx_pt.Y = objDataSet.Top + (objDataSet.Height / 2)
        Case MiddleRight
            dx_pt.X = objDataSet.Left + objDataSet.Width - margin
            dx_pt.Y = objDataSet.Top + (objDataSet.Height / 2)
        Case BottomLeft
            dx_pt.X = objDataSet.Left + margin
            dx_pt.Y = objDataSet.Top + objDataSet.Height - margin
        Case BottomCenter
            dx_pt.X = objDataSet.Left + (objDataSet.Width / 2)
            dx_pt.Y = objDataSet.Top + objDataSet.Height - margin
        Case BottomRight
            dx_pt.X = objDataSet.Left + objDataSet.Width - margin
            dx_pt.Y = objDataSet.Top + objDataSet.Height - margin
    End Select
            
    dx_pt.X = dx_pt.X - WS.Cells(tb.Row, tb.Col).Left
    dx_pt.Y = dx_pt.Y - WS.Cells(tb.Row, tb.Col).Top
    
    '実際の移動位置でもう一度精度を高めるためにズーム値を計算。
    Application.ScreenUpdating = (Val(Application.Version) <= C_EXCEL_VERSION_2010)
        
    zm = GetTrueZoom(dx_pt)
    
    Dim dx_px As POINTAPI
    dx_px = PointToPixelZoom(dx_pt, zm)
    
    dx_px.X = dx_px.X + org_px.X
    dx_px.Y = dx_px.Y + org_px.Y
    
    'シェイプを選択
    objDataSet.TopLeftCell.Select
    Call SetCursoleAndLeftDown(dx_px.X, dx_px.Y)

    Application.ScreenUpdating = True
    
    '---------------------------
    '終了イベント
    '---------------------------
    RaiseEvent SelectionTerm
    
    Exit Sub
ErrHandle:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。", vbOKOnly, C_TITLE

End Sub
' Point per Inch
Private Function PPI() As Double
    PPI = Application.InchesToPoints(1)
End Function
' Dot per Inch
Private Function DPIX() As Double
    DPIX = dblSx
End Function
Private Function DPIY() As Double
    DPIY = dblSy
End Function 'ポイント から スクリーン座標
Private Sub GetDPI()

#If VBA7 And Win64 Then
    Dim hWnd As LongPtr
    Dim hDc As LongPtr
#Else
    Dim hWnd As Long
    Dim hDc As Long
#End If
    hWnd = Application.hWnd
    hDc = GetDC(hWnd)
    dblSx = GetDeviceCaps(hDc, LOGPIXELSX)
    dblSy = GetDeviceCaps(hDc, LOGPIXELSY)
    ReleaseDC hWnd, hDc
    
End Sub
Private Function PointToPixel(ByRef pt As POINTAPI) As POINTAPI

    Dim px As POINTAPI
    
    px.X = Round(pt.X * DPIX / PPI)
    px.Y = Round(pt.Y * DPIY / PPI)
    
    PointToPixel = px
    
End Function
'スクリーン座標からポイント
Private Function PixelToPoint(ByRef px As POINTAPI) As POINTAPI

    Dim pt As POINTAPI
    
    pt.X = Round(px.X * PPI / DPIX)
    pt.Y = Round(px.Y * PPI / DPIY)
    
    PixelToPoint = pt
    
End Function
'スクリーン座標からポイント(ズーム対応)
Private Function PixelToPointZoom(ByRef px As POINTAPI, ByRef zm As ZOOMVALUE) As POINTAPI

    Dim pt As POINTAPI
    
    pt.X = Round((px.X * PPI / DPIX) / zm.X)
    pt.Y = Round((px.Y * PPI / DPIY) / zm.Y)
    
    PixelToPointZoom = pt
    
End Function
'ポイント から スクリーン座標(ズーム対応)
Private Function PointToPixelZoom(ByRef pt As POINTAPI, ByRef zm As ZOOMVALUE) As POINTAPI

    Dim px As POINTAPI
    
    px.X = Round((pt.X * DPIX / PPI) * zm.X)
    px.Y = Round((pt.Y * DPIY / PPI) * zm.Y)
    
    PointToPixelZoom = px
    
End Function
'シートの左上位置(Pixel)ヘッダや行番号を含むためZOOMにより変動する
'Excel 2010以前ではExcelのバグで、ScreenUpdating = False の場合に値が取れないので注意
Private Function GetOriginPixel() As POINTAPI

    Dim dummy As Long
    Dim ret As POINTAPI
    
    dummy = WD.SplitHorizontal    'PointsToScreenPixelsXの値を更新するために使用
    ret.X = WD.ActivePane.PointsToScreenPixelsX(WD.ActivePane.VisibleRange(1).Left)

    dummy = WD.SplitVertical    'PointsToScreenPixelsXの値を更新するために使用
    ret.Y = WD.ActivePane.PointsToScreenPixelsY(WD.ActivePane.VisibleRange(1).Top)
    
    GetOriginPixel = ret

End Function
'真のズームを求める
'引数：ポイント（距離によりが割合が変わるためそのシード数）
'計算しやすいように小数で返却(85% = 0.85)
Private Function GetTrueZoom(pt As POINTAPI) As ZOOMVALUE

    Dim px As POINTAPI
    Dim p1 As POINTAPI
    Dim p2 As POINTAPI
    Dim dummy As Long
    Dim ret As ZOOMVALUE
    
    If WD.Zoom = 100 Then
        ret.X = 1
        ret.Y = 1
        GetTrueZoom = ret
        Exit Function
    End If
    
    '100% 時のPixel数
    px = PointToPixel(pt)

    dummy = WD.SplitHorizontal    'PointsToScreenPixelsXの値を更新するために使用
    p1.X = WD.ActivePane.PointsToScreenPixelsX(0)
    
    dummy = WD.SplitVertical    'PointsToScreenPixelsYの値を更新するために使用
    p1.Y = WD.ActivePane.PointsToScreenPixelsY(0)

    dummy = WD.SplitHorizontal    'PointsToScreenPixelsXの値を更新するために使用
    p2.X = WD.ActivePane.PointsToScreenPixelsX(pt.X) - p1.X
    
    dummy = WD.SplitVertical    'PointsToScreenPixelsYの値を更新するために使用
    p2.Y = WD.ActivePane.PointsToScreenPixelsY(pt.Y) - p1.Y

    '100% のピクセル数と 現在の% のピクセル数で、倍率を計算
    If px.X = 0 Then
        ret.X = 1
    Else
        ret.X = CDbl(p2.X) / CDbl(px.X)
    End If
        
    If px.Y = 0 Then
        ret.Y = 1
    Else
        ret.Y = CDbl(p2.Y) / CDbl(px.Y)
    End If
    
    GetTrueZoom = ret

End Function

'指定スクリーン座標に移動してマウス左ボタンクリック
Private Sub SetCursoleAndLeftDown(ByVal X As Long, ByVal Y As Long)

    Dim inp(0 To 1) As INPUT_TYPE
    
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

'指定スクリーン座標に移動してマウス左ボタンクリック
Private Sub SetCursole(ByVal X As Long, ByVal Y As Long)

    Dim inp(0 To 0) As INPUT_TYPE
    
    With inp(0)
        .dwType = INPUT_MOUSE
        .mi.dx = (X * 65535 / (GetSystemMetrics(SM_CXSCREEN) - 1))
        .mi.dy = (Y * 65535 / (GetSystemMetrics(SM_CYSCREEN) - 1))
        .mi.mouseData = 0
        .mi.dwFlags = MOUSE_MOVED Or MOUSEEVENTF_ABSOLUTE
        .mi.time = 0
        .mi.dwExtraInfo = 0
    End With
    
    SendInput 1, inp(0), LenB(inp(0))

End Sub
