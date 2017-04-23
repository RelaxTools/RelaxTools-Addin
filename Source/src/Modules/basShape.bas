Attribute VB_Name = "basShape"
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

'１マスのサイズ（9.75×12）
Public Const C_RECT_X  As Single = 9.75
Public Const C_RECT_Y  As Single = 12

'その他の構成（４×３マス）
Public Const C_NORMAL_WIDTH As Long = 7
Public Const C_NORMAL_HEIGHT As Long = 4

'
' 線を真っ直ぐにする
'
Sub straightLine()

    Dim s As Object

    On Error GoTo e
        
    For Each s In Selection.ShapeRange
    
        Dim w As Long
        Dim h As Long
        
        w = s.width
        h = s.Height
        
        If w > h Then
            If s.VerticalFlip Then
                s.Top = s.Top + s.Height
                s.Height = 0
            Else
                s.Height = 0
            End If
        Else
            If s.HorizontalFlip Then
                s.Left = s.Left + s.width
                s.width = 0
            Else
                s.width = 0
            End If
        End If
    Next
e:

End Sub
Sub largeShape()

    On Error Resume Next
    Selection.ShapeRange.ScaleHeight 1.1, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleWidth 1.1, msoFalse, msoScaleFromTopLeft
End Sub
Sub smallShape()
    On Error Resume Next
    Selection.ShapeRange.ScaleHeight 0.9, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleWidth 0.9, msoFalse, msoScaleFromTopLeft
End Sub
'
' 記憶データ（シェイプ）描画
'
Sub drawFlowchartStoredData()

    Dim objDataSet As Shape
    Dim r As Range

    For Each r In Selection
        
        'データ記憶シェイプの作成
        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeFlowchartStoredData, r.Left + r.width, r.Top, C_RECT_X * C_NORMAL_WIDTH, C_RECT_Y * C_NORMAL_HEIGHT)
    
        With objDataSet.TextFrame
            .Characters.Text = r.Value
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        Set objDataSet = Nothing
    Next

End Sub
'
' テキストボックス（シェイプ）描画
'
Sub drawTextbox1()

    Dim objDataSet As Shape
    Dim r As Range
    Dim strBuf As String
    Dim lngCnt As Long
    

    For Each r In Selection
        
        strBuf = r.Value
        
        lngCnt = InStr(strBuf, vbCrLf) + 3
        
        'データ記憶シェイプの作成
        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangle, r.Left + r.width, r.Top, C_RECT_X * 10, C_RECT_Y * lngCnt)
    
        With objDataSet.TextFrame
            .Characters.Text = r.Value
        End With
            

        Set objDataSet = Nothing
    Next

End Sub
'
' テキストボックス（シェイプ）描画
'
Sub drawTextbox2()

    Dim objDataSet As Shape
    Dim r As Range
    Dim strBuf As String
    Dim lngCnt As Long
    

    For Each r In Selection
        
        strBuf = r.Value
        
        lngCnt = InStr(strBuf, vbCrLf) + 3
        
        'データ記憶シェイプの作成
        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangle, r.Left + r.width, r.Top, C_RECT_X * 10, C_RECT_Y * lngCnt)
    
        With objDataSet.TextFrame
            .Characters.Text = r.Value
        End With
            
        '枠なしの場合
        objDataSet.Line.visible = msoFalse
            
        Set objDataSet = Nothing
    Next

End Sub
'
' 四角形吹き出し（シェイプ）描画
'
Sub drawShapeRectangularCallout()

    Dim objDataSet As Shape
    Dim r As Range
    
    Dim strBuf As String
    Dim lngCnt As Long

    For Each r In Selection
        
        strBuf = r.Value
        
        lngCnt = InStr(strBuf, vbCrLf) + 3
        
        'データ記憶シェイプの作成
        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangularCallout, r.Left + r.width, r.Top, C_RECT_X * 10, C_RECT_Y * lngCnt)
    
        With objDataSet.TextFrame
            .Characters.Text = r.Value
        End With
        
        Set objDataSet = Nothing
    Next

End Sub
'
' 角丸四角形吹き出し（シェイプ）描画
'
Sub drawShapeRoundedRectangularCallout()

    Dim objDataSet As Shape
    Dim r As Range
    
    Dim strBuf As String
    Dim lngCnt As Long

    For Each r In Selection
        
        strBuf = r.Value
        
        lngCnt = InStr(strBuf, vbCrLf) + 3
        
        'データ記憶シェイプの作成
        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangularCallout, r.Left + r.width, r.Top, C_RECT_X * 10, C_RECT_Y * lngCnt)
    
        With objDataSet.TextFrame
            .Characters.Text = r.Value
        End With
        
        Set objDataSet = Nothing
    Next

End Sub
'
' 丸形吹き出し（シェイプ）描画
'
Sub drawShapeOvalCallout()

    Dim objDataSet As Shape
    Dim r As Range
    
    Dim strBuf As String
    Dim lngCnt As Long

    For Each r In Selection
        
        strBuf = r.Value
        
        lngCnt = InStr(strBuf, vbCrLf) + 3
        
        'データ記憶シェイプの作成
        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeOvalCallout, r.Left + r.width, r.Top, C_RECT_X * 10, C_RECT_Y * lngCnt)
    
        With objDataSet.TextFrame
            .Characters.Text = r.Value
        End With
        
        Set objDataSet = Nothing
    Next

End Sub
'
' 雲形吹き出し（シェイプ）描画
'
Sub drawShapeCloudCallout()

    Dim objDataSet As Shape
    Dim r As Range
    
    Dim strBuf As String
    Dim lngCnt As Long

    For Each r In Selection
        
        strBuf = r.Value
        
        lngCnt = InStr(strBuf, vbCrLf) + 3
        
        'データ記憶シェイプの作成
        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeCloudCallout, r.Left + r.width, r.Top, C_RECT_X * 10, C_RECT_Y * lngCnt)
    
        With objDataSet.TextFrame
            .Characters.Text = r.Value
        End With
        
        Set objDataSet = Nothing
    Next

End Sub
'
' 丸形（シェイプ）変換
'
Sub convShapeOval()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeOval
    Next

End Sub
'
' 四角形（シェイプ）変換
'
Sub convShapeRectangle()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeRectangle
    Next

End Sub
'
' 四角形吹き出し（シェイプ）変換
'
Sub convShapeRectangularCallout()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeRectangularCallout
    Next

End Sub
'
' 角丸四角形吹き出し（シェイプ）変換
'
Sub convShapeRoundedRectangularCallout()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeRoundedRectangularCallout
    Next

End Sub
'
' 丸形吹き出し（シェイプ）変換
'
Sub convShapeOvalCallout()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeOvalCallout
    Next

End Sub
'
' 雲形吹き出し（シェイプ）変換
'
Sub convShapeCloudCallout()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeCloudCallout
    Next

End Sub
'
' フローチャート：処理（シェイプ）変換
'
Sub convShapeFlowchartProcess()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartProcess
    Next

End Sub
'
' フローチャート：代替処理（シェイプ）変換
'
Sub convShapeFlowchartAlternateProcess()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartAlternateProcess
    Next

End Sub
'
' フローチャート：判断（シェイプ）変換
'
Sub convShapeFlowchartDecision()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartDecision
    Next

End Sub
'
' フローチャート：データ（シェイプ）変換
'
Sub convShapeFlowchartData()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartData
    Next

End Sub
'
' フローチャート：定義済み処理（シェイプ）変換
'
Sub convShapeFlowchartPredefinedProcess()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartPredefinedProcess
    Next

End Sub
'
' フローチャート：内部記憶（シェイプ）変換
'
Sub convShapeFlowchartInternalStorage()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartInternalStorage
    Next

End Sub
'
' フローチャート：書類（シェイプ）変換
'
Sub convShapeFlowchartDocument()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartDocument
    Next

End Sub
'
' フローチャート：複数書類（シェイプ）変換
'
Sub convShapeFlowchartMultidocument()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartMultidocument
    Next

End Sub
'
' フローチャート：端子（シェイプ）変換
'
Sub convShapeFlowchartTerminator()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartTerminator
    Next

End Sub
'
' フローチャート：準備（シェイプ）変換
'
Sub convShapeFlowchartPreparation()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartPreparation
    Next

End Sub
'
' フローチャート：手操作入力（シェイプ）変換
'
Sub convShapeFlowchartManualInput()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartManualInput
    Next

End Sub
'
' フローチャート：手作業（シェイプ）変換
'
Sub convShapeFlowchartManualOperation()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartManualOperation
    Next

End Sub
'
' フローチャート：カード（シェイプ）変換
'
Sub convShapeFlowchartCard()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartCard
    Next

End Sub
'
' フローチャート：せん孔テープ（シェイプ）変換
'
Sub convShapeFlowchartPunchedTape()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartPunchedTape
    Next

End Sub
'
' フローチャート：記憶データ（シェイプ）変換
'
Sub convShapeFlowchartStoredData()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartStoredData
    Next

End Sub
'
' フローチャート：順次アクセス記憶（シェイプ）変換
'
Sub convShapeFlowchartSequentialAccessStorage()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartSequentialAccessStorage
    Next

End Sub
'
' フローチャート：直接アクセス記憶（シェイプ）変換
'
Sub convShapeFlowchartDirectAccessStorage()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartDirectAccessStorage
    Next

End Sub
'
' フローチャート：磁気ディスク（シェイプ）変換
'
Sub convShapeFlowchartMagneticDisk()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartMagneticDisk
    Next

End Sub
'
' フローチャート：表示（シェイプ）変換
'
Sub convShapeFlowchartDisplay()

    Dim r As Shape
    
    For Each r In Selection.ShapeRange
        r.AutoShapeType = msoShapeFlowchartDisplay
    Next

End Sub

'
' エビデンス用四角（シェイプ）描画
'
Sub drawEvidenceTextbox()

    Dim obj As ShapeDrawEvidenceTextbox
    
    Set obj = New ShapeDrawEvidenceTextbox
    
    obj.Run
    
    Set obj = Nothing


'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.Left + (Selection.Width - C_RECT_X * C_NORMAL_WIDTH * 3) / 2, Selection.Top + (Selection.Height - C_RECT_Y * C_NORMAL_HEIGHT) / 2, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveWindow.ActivePane.VisibleRange.Left + 100, ActiveWindow.ActivePane.VisibleRange.Top + 100, C_RECT_X * C_NORMAL_WIDTH * 2, C_RECT_Y * C_NORMAL_HEIGHT)
'    End If
'
'    '基本のスタイルをセット
'    objDataSet.ShapeStyle = msoShapeStylePreset1
'
'    With objDataSet.Fill
'        .Solid
'        .visible = msoTrue
'        .Transparency = 1
'        .ForeColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineSingle
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a
        
End Sub
'
' エビデンス用楕円（シェイプ）描画
'
Sub drawEvidenceOval()

    Dim obj As ShapeDrawEvidenceOval
    
    Set obj = New ShapeDrawEvidenceOval
    
    obj.Run
    
    Set obj = Nothing


'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeOval, Selection.Left + (Selection.Width - C_RECT_X * C_NORMAL_WIDTH * 1.5) / 2, Selection.Top + (Selection.Height - C_RECT_Y * C_NORMAL_HEIGHT) / 2, C_RECT_X * C_NORMAL_WIDTH * 1.5, C_RECT_Y * C_NORMAL_HEIGHT)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeOval, ActiveWindow.ActivePane.VisibleRange.Left + 100, ActiveWindow.ActivePane.VisibleRange.Top + 100, C_RECT_X * C_NORMAL_WIDTH * 1.5, C_RECT_Y * C_NORMAL_HEIGHT)
'    End If
'
'    '基本のスタイルをセット
'    objDataSet.ShapeStyle = msoShapeStylePreset1
'
'    With objDataSet.Fill
'        .Solid
'        .visible = msoTrue
'        .Transparency = 1
'        .ForeColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineSingle
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a
        
End Sub
'
' エビデンス用ふきだし描画
'
Sub drawEvidenceCallout()

    Dim obj As ShapeDrawEvidenceCallout
    
    Set obj = New ShapeDrawEvidenceCallout
    
    obj.Run
    
    Set obj = Nothing

'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangularCallout, Selection.Left + (Selection.Width - C_RECT_X * C_NORMAL_WIDTH * 3) / 2, Selection.Top + (Selection.Height - C_RECT_Y * C_NORMAL_HEIGHT) / 2, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangularCallout, ActiveWindow.ActivePane.VisibleRange.Left + 100, ActiveWindow.ActivePane.VisibleRange.Top + 100, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    End If
'
'    '基本のスタイルをセット
'    objDataSet.ShapeStyle = msoShapeStylePreset1
'
'    With objDataSet.Fill
'        .Solid
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineSingle
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.TextFrame2.TextRange.Characters.Font
'        .Fill.ForeColor.RGB = RGB(0, 0, 0)
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    'シェイプをつまむ
'    PickShape objDataSet
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a

End Sub
'
' エビデンス用線ふきだし描画
'
Sub drawEvidenceLineCallout()

    Dim obj As ShapeDrawEvidenceLineCallout
    
    Set obj = New ShapeDrawEvidenceLineCallout
    
    obj.Run
    
    Set obj = Nothing


'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeLineCallout1, Selection.Left + (Selection.Width - C_RECT_X * C_NORMAL_WIDTH * 3) / 2, Selection.Top + (Selection.Height - C_RECT_Y * C_NORMAL_HEIGHT) / 2, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeLineCallout1, ActiveWindow.ActivePane.VisibleRange.Left + 100, ActiveWindow.ActivePane.VisibleRange.Top + 100, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    End If
'
'    '基本のスタイルをセット
'    objDataSet.ShapeStyle = msoShapeStylePreset1
'
'    With objDataSet.Fill
'        .Solid
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineSingle
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.TextFrame2.TextRange.Characters.Font
'        .Fill.ForeColor.RGB = RGB(0, 0, 0)
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a

End Sub
'
' エビデンス用矢印描画
'
Sub drawEvidenceArrow()

    Dim obj As ShapeDrawEvidenceArrow
    
    Set obj = New ShapeDrawEvidenceArrow
    
    obj.Run
    
    Set obj = Nothing

'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, Selection.Left + (Selection.Width / 2), Selection.Top + Selection.Height - (C_NORMAL_HEIGHT * 25), Selection.Left + (Selection.Width / 2), Selection.Top + Selection.Height)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, ActiveWindow.ActivePane.VisibleRange.Left + 100, Selection.Top + Selection.Height - (C_NORMAL_HEIGHT * 25), ActiveWindow.ActivePane.VisibleRange.Left + 100, Selection.Top + Selection.Height)
'    End If
'
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineSingle
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'        .EndArrowheadStyle = msoArrowheadOpen
'        .EndArrowheadLength = msoArrowheadLong
'        .EndArrowheadWidth = msoArrowheadWide
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a
'

End Sub
'
' エビデンス用四角（シェイプ）描画2
'
Sub drawEvidenceTextbox2()


    Dim obj As ShapeDrawEvidenceTextbox2
    
    Set obj = New ShapeDrawEvidenceTextbox2
    
    obj.Run
    
    Set obj = Nothing

'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.Left + (Selection.Width - C_RECT_X * C_NORMAL_WIDTH * 3) / 2, Selection.Top + (Selection.Height - C_RECT_Y * C_NORMAL_HEIGHT) / 2, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveWindow.ActivePane.VisibleRange.Left + 100, ActiveWindow.ActivePane.VisibleRange.Top + 100, C_RECT_X * C_NORMAL_WIDTH * 2, C_RECT_Y * C_NORMAL_HEIGHT)
'    End If
'
'    '基本のスタイルをセット
'    objDataSet.ShapeStyle = msoShapeStylePreset1
'
'    With objDataSet.Fill
'        .Solid
'        .visible = msoTrue
'        .Transparency = 1
'        .ForeColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineDash
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a

End Sub
'
' エビデンス用楕円（シェイプ）描画2
'
Sub drawEvidenceOval2()

    Dim obj As ShapeDrawEvidenceOval2
    
    Set obj = New ShapeDrawEvidenceOval2
    
    obj.Run
    
    Set obj = Nothing

'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeOval, Selection.Left + (Selection.Width - C_RECT_X * C_NORMAL_WIDTH * 1.5) / 2, Selection.Top + (Selection.Height - C_RECT_Y * C_NORMAL_HEIGHT) / 2, C_RECT_X * C_NORMAL_WIDTH * 1.5, C_RECT_Y * C_NORMAL_HEIGHT)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeOval, ActiveWindow.ActivePane.VisibleRange.Left + 100, ActiveWindow.ActivePane.VisibleRange.Top + 100, C_RECT_X * C_NORMAL_WIDTH * 1.5, C_RECT_Y * C_NORMAL_HEIGHT)
'    End If
'
'    '基本のスタイルをセット
'    objDataSet.ShapeStyle = msoShapeStylePreset1
'
'    With objDataSet.Fill
'        .Solid
'        .visible = msoTrue
'        .Transparency = 1
'        .ForeColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineDash
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a

End Sub
'
' エビデンス用ふきだし描画2
'
Sub drawEvidenceCallout2()

    Dim obj As ShapeDrawEvidenceCallout2
    
    Set obj = New ShapeDrawEvidenceCallout2
    
    obj.Run
    
    Set obj = Nothing
    
'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangularCallout, Selection.Left + (Selection.Width - C_RECT_X * C_NORMAL_WIDTH * 3) / 2, Selection.Top + (Selection.Height - C_RECT_Y * C_NORMAL_HEIGHT) / 2, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeRectangularCallout, ActiveWindow.ActivePane.VisibleRange.Left + 100, ActiveWindow.ActivePane.VisibleRange.Top + 100, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    End If
'
'    '基本のスタイルをセット
'    objDataSet.ShapeStyle = msoShapeStylePreset1
'
'    With objDataSet.Fill
'        .Solid
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineDash
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'    End With
'
'    With objDataSet.TextFrame2.TextRange.Characters.Font
'        .Fill.ForeColor.RGB = RGB(0, 0, 0)
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a

End Sub
'
' エビデンス用線ふきだし描画2
'
Sub drawEvidenceLineCallout2()

    Dim obj As ShapeDrawEvidenceLineCallout2
    
    Set obj = New ShapeDrawEvidenceLineCallout2
    
    obj.Run
    
    Set obj = Nothing
'
'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeLineCallout1, Selection.Left + (Selection.Width - C_RECT_X * C_NORMAL_WIDTH * 3) / 2, Selection.Top + (Selection.Height - C_RECT_Y * C_NORMAL_HEIGHT) / 2, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddShape(msoShapeLineCallout1, ActiveWindow.ActivePane.VisibleRange.Left + 100, ActiveWindow.ActivePane.VisibleRange.Top + 100, C_RECT_X * C_NORMAL_WIDTH * 3, C_RECT_Y * C_NORMAL_HEIGHT)
'    End If
'
'    '基本のスタイルをセット
'    objDataSet.ShapeStyle = msoShapeStylePreset1
'
'    With objDataSet.Fill
'        .Solid
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.RGB = RGB(255, 255, 255)
'    End With
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineDash
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'    End With
'
'    With objDataSet.TextFrame2.TextRange.Characters.Font
'        .Fill.ForeColor.RGB = RGB(0, 0, 0)
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a

End Sub
'
' エビデンス用矢印描画
'
Sub drawEvidenceArrow2()

    Dim obj As ShapeDrawEvidenceArrow2
    
    Set obj = New ShapeDrawEvidenceArrow2
    
    obj.Run
    
    Set obj = Nothing

'
'    Dim objDataSet As Shape
'    Dim z As Single
'    Dim r As Long
'    Dim c As Long
'    Dim a As Boolean
'
'    If ActiveWorkbook Is Nothing Then
'        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
'        Exit Sub
'    End If
'
'    a = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        z = ActiveWindow.Zoom
'        c = ActiveWindow.ScrollColumn
'        r = ActiveWindow.ScrollRow
'        ActiveWindow.Zoom = 100
'        Set objDataSet = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, Selection.Left + (Selection.Width / 2), Selection.Top + Selection.Height - (C_NORMAL_HEIGHT * 25), Selection.Left + (Selection.Width / 2), Selection.Top + Selection.Height)
'    Else
'        Set objDataSet = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, ActiveWindow.ActivePane.VisibleRange.Left + 100, Selection.Top + Selection.Height - (C_NORMAL_HEIGHT * 25), ActiveWindow.ActivePane.VisibleRange.Left + 100, Selection.Top + Selection.Height)
'    End If
'
'    With objDataSet.Line
'        .Weight = 2.25
'        .DashStyle = msoLineDash
'        .Style = msoLineSingle
'        .Transparency = 0#
'        .visible = msoTrue
'        .ForeColor.SchemeColor = 10
'        .BackColor.RGB = RGB(255, 255, 255)
'        .EndArrowheadStyle = msoArrowheadOpen
'        .EndArrowheadLength = msoArrowheadLong
'        .EndArrowheadWidth = msoArrowheadWide
'    End With
'
'    objDataSet.Select
'    objDataSet.Placement = xlMove
'
'    If CBool(GetSetting(C_TITLE, "Shape", "PickMode", False)) = False Then
'        ActiveWindow.Zoom = z
'        ActiveWindow.ScrollColumn = c
'        ActiveWindow.ScrollRow = r
'    Else
'        'シェイプをつまむ
'        PickShape objDataSet
'    End If
'
'    Set objDataSet = Nothing
'
'    Application.ScreenUpdating = a

End Sub

Sub shapeAllSelect()
    ActiveSheet.Shapes.SelectAll
End Sub
Sub shapeAllDelete()

    On Error Resume Next
    
    Dim WS As Worksheet
    
    Set WS = ActiveSheet
    If MsgBox("アクティブシートのシェイプ／画像をすべて削除します。よろしいですか？", vbYesNo + vbQuestion, C_TITLE) <> vbYes Then
        Exit Sub
    End If
    WS.Shapes.SelectAll
    Selection.Delete
    Set WS = Nothing
End Sub
