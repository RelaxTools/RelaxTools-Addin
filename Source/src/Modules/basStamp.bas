Attribute VB_Name = "basStamp"
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

Public Const C_STAMP_DATE_SYSTEM As String = "1"
Public Const C_STAMP_DATE_USER As String = "2"
Public Const C_STAMP_LINE_SINGLE As String = "1"
Public Const C_STAMP_LINE_DOUBLE As String = "2"
Public Const C_STAMP_LINE_BOLD As String = "3"
Public Const C_STAMP_FILE_NAME = "relaxStamp"

Public Const C_STAMP_WORDART_OFF As String = "0"
Public Const C_STAMP_WORDART_ON As String = "1"

Public Const C_STAMP_FILL_OFF As String = "0"
Public Const C_STAMP_FILL_ON As String = "1"

Public Const C_RASIO As Single = 2.83

'--------------------------------------------------------------
'　職印設定画面
'--------------------------------------------------------------
Sub showStamp()

    frmStamp.Show
    
End Sub

Public Function editStamp(ByRef s As StampDatDTO, ByVal lngFormat As Long) As StdPicture

    Dim strMiddle As String
    Dim WS As Worksheet
    
    Set editStamp = Nothing
    
    Set WS = ThisWorkbook.Worksheets("stampEx")

    strMiddle = getFormatDate(s.DateFormat, s.DateType, s.UserDate)
    
    If Len(s.Upper) > 1 Then
        '上段
        WS.Shapes("shpUp").TextFrame2.TextRange.Text = s.Upper
        With WS.Shapes("shpUp").TextFrame2.TextRange.Font
            .NameComplexScript = s.Font
            .NameFarEast = s.Font
            .Name = s.Font
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
    
             .Fill.ForeColor.RGB = CLng(s.Color)
        End With
        WS.Shapes("shpUp").visible = True
        WS.Shapes("shpUp2").visible = False
    Else
        '上段
        WS.Shapes("shpUp2").TextFrame2.TextRange.Text = s.Upper
        With WS.Shapes("shpUp2").TextFrame2.TextRange.Font
            .NameComplexScript = s.Font
            .NameFarEast = s.Font
            .Name = s.Font
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
    
             .Fill.ForeColor.RGB = CLng(s.Color)
        End With
        WS.Shapes("shpUp2").visible = True
        WS.Shapes("shpUp").visible = False
    End If
    
    '中段
    WS.Shapes("shpMid").TextFrame2.TextRange.Text = strMiddle
    With WS.Shapes("shpMid").TextFrame2.TextRange.Font
        .NameComplexScript = s.Font
        .NameFarEast = s.Font
        .Name = s.Font
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        
         .Fill.ForeColor.RGB = CLng(s.Color)
    End With
    
    If Len(s.Lower) > 1 Then
        '下段
        WS.Shapes("shpLow").TextFrame2.TextRange.Text = s.Lower
        With WS.Shapes("shpLow").TextFrame2.TextRange.Font
            .NameComplexScript = s.Font
            .NameFarEast = s.Font
            .Name = s.Font
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
             
             .Fill.ForeColor.RGB = CLng(s.Color)
        End With
        WS.Shapes("shpLow").visible = True
        WS.Shapes("shpLow2").visible = False
    Else
        '下段
        WS.Shapes("shpLow2").TextFrame2.TextRange.Text = s.Lower
        With WS.Shapes("shpLow2").TextFrame2.TextRange.Font
            .NameComplexScript = s.Font
            .NameFarEast = s.Font
            .Name = s.Font
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
             
             .Fill.ForeColor.RGB = CLng(s.Color)
        End With
        WS.Shapes("shpLow2").visible = True
        WS.Shapes("shpLow").visible = False
    End If
    
    On Error Resume Next
    
    Dim r As Shape
    
    Set r = WS.Shapes("grpStamp")
    
    r.Line.ForeColor.RGB = CLng(s.Color)
    
    Select Case s.Line
        Case C_STAMP_LINE_SINGLE
            WS.Shapes("shpCircle").Line.Weight = 10
            WS.Shapes("shpCircle").Line.style = msoLineSingle
        Case C_STAMP_LINE_DOUBLE
            WS.Shapes("shpCircle").Line.Weight = 20
            WS.Shapes("shpCircle").Line.style = msoLineThinThin
        Case Else
            WS.Shapes("shpCircle").Line.Weight = 20
            WS.Shapes("shpCircle").Line.style = msoLineSingle
    End Select
    
    If s.Fill = C_STAMP_FILL_ON Then
        WS.Shapes("shpCircle").Fill.visible = True
        WS.Shapes("shpCircle").Fill.ForeColor.RGB = vbWhite
    Else
        WS.Shapes("shpCircle").Fill.visible = False
    End If
    
    r.Rotation = getRect(s.rect)
    
    If lngFormat = xlBitmap Then
    
        Dim b As Shape
        Dim o As Object
        
        Set b = WS.Shapes("shpBack")
        
        b.Top = r.Top - ((r.width - r.Height) / 2)
        b.Left = r.Left
        b.Height = r.width
        b.width = r.width
        
        b.ZOrder msoSendToBack
        
        Set o = WS.Shapes.Range(Array(r.Name, b.Name)).Group
        
        Set editStamp = CreatePictureFromClipboard(o)
        
        o.Ungroup

    Else
        r.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        Call CopyClipboardSleep
    End If
    
    Set WS = Nothing
    
End Function
Function getRect(rect As String) As Single
    Dim a As Single
    Dim b As Single
    a = (Val(rect) / 100) * -1
    
    b = (a * 180)
    If b < 0 Then
        b = b + 360
    End If
    getRect = b '+ 90
    
End Function
'--------------------------------------------------------------
'　bzイメージファイル作成
'--------------------------------------------------------------
Function getImageStamp(ByVal Index As Long) As StdPicture

    '設定情報取得
    Dim col As Collection
    Dim bz As StampDatDTO
    
    Set getImageStamp = Nothing
    
    Set col = getProperty()
    
    Set bz = col(Index)
    
    Set getImageStamp = editStamp(bz, xlBitmap)

    Set bz = Nothing
    
End Function
Public Sub StampPaste()

    Dim lngNo As Long

    lngNo = GetSetting(C_TITLE, "Stamp", "stampNo", 1)
    Call pasteStamp2(lngNo)

End Sub

'--------------------------------------------------------------
'　データ印貼り付け
'--------------------------------------------------------------
Sub pasteStamp2(ByVal Index As Long)

    '設定情報取得
    Dim col As Collection
    Dim s As StampDatDTO
    Dim r As Shape
    Dim ss As Range

    On Error Resume Next

    If rlxCheckSelectRange() = False Then
        MsgBox "選択範囲が見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If

    If GetSetting(C_TITLE, "Stamp", "Confirm", False) Then
    Else
        If Selection.CountLarge > 1 And Selection.CountLarge <> Selection(1, 1).MergeArea.count Then
            If MsgBox("複数セル選択されています。すべてのセルに張り付けますがよろしいですか？", vbQuestion + vbYesNo, C_TITLE) <> vbYes Then
                Exit Sub
            End If
        End If
    End If
    
    Set col = getProperty()

    Select Case True
        Case col Is Nothing
            Exit Sub
        Case col.count = 0
            Exit Sub
        Case Else
    End Select

    Set s = col(Index)

    Call editStamp(s, xlPicture)

    Dim sngSize As Single

    sngSize = CSng(s.size) * C_RASIO

    Dim destLeft As Long
    Dim destWidth As Long
    Dim destTop As Long
    Dim destHeight As Long

    Application.ScreenUpdating = False

    For Each ss In Selection

        ''フィルタおよび非表示対策。
        If ss.Rows.Hidden Or ss.Columns.Hidden Then
            'フィルタまたは非表示の行・列の処理は行わない。
        Else

            If ss.Address = ss.MergeArea(1, 1).Address Then

                destLeft = ss.MergeArea.Left
                destWidth = ss.MergeArea.width
                destTop = ss.MergeArea.Top
                destHeight = ss.MergeArea.Height

                ActiveSheet.Paste

                Selection.ShapeRange.width = sngSize

                Selection.ShapeRange.Top = destTop + (destHeight / 2) - (Selection.ShapeRange.Height / 2)
                Selection.ShapeRange.Left = destLeft + (destWidth / 2) - (Selection.ShapeRange.width / 2)
            End If
        End If
    Next

    Selection.Copy
    
    Application.ScreenUpdating = True

End Sub

'--------------------------------------------------------------
'　レジストリ設定値取得
'--------------------------------------------------------------
Public Function getProperty() As Collection

    Dim strBuf As String
    Dim s As StampDatDTO
    Dim lngMax As Long
    Dim i As Long
    
    Dim col As Collection
    
    Set col = New Collection

    lngMax = GetSetting(C_TITLE, "Stamp", "Count", "-1")
    If lngMax = -1 Then
    
        Set s = New StampDatDTO
        
        s.Upper = "山"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = ""
        s.Lower = "田"
        s.Font = "ＭＳ ゴシック"
        s.Color = "&H0"
        s.Line = C_STAMP_LINE_SINGLE
        s.size = "15"
        s.WordArt = C_STAMP_WORDART_ON
        s.Fill = C_STAMP_FILL_OFF
        s.rect = "0"

        col.Add s
        
        Set s = Nothing
        
        Set s = New StampDatDTO
        
        s.Upper = "二課"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = ""
        s.Lower = "勅使河原"
        s.Font = "ＭＳ ゴシック"
        s.Color = "&HFF"
        s.Line = C_STAMP_LINE_SINGLE
        s.size = "15"
        s.WordArt = C_STAMP_WORDART_ON
        s.Fill = C_STAMP_FILL_OFF
        s.rect = "0"
    
        col.Add s
        
        Set s = Nothing
    
        Set s = New StampDatDTO
        
        s.Upper = "検"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "品質管理課"
        s.UserDate = ""
        s.Lower = "印"
        s.Font = "ＭＳ ゴシック"
        s.Color = "&H0"
        s.Line = C_STAMP_LINE_SINGLE
        s.size = "15"
        s.WordArt = C_STAMP_WORDART_ON
        s.Fill = C_STAMP_FILL_OFF
        s.rect = "0"
    
        col.Add s
        
        Set s = Nothing
    Else
        For i = 0 To lngMax - 1
            
            Set s = New StampDatDTO
        
            s.Upper = Replace(GetSetting(C_TITLE, "Stamp", "Upper" & Format$(i, "000"), "ＸＸ課"), vbVerticalTab, vbCrLf)
            s.DateType = GetSetting(C_TITLE, "Stamp", "DateType" & Format$(i, "000"), C_STAMP_DATE_SYSTEM)
            s.DateFormat = GetSetting(C_TITLE, "Stamp", "DateFormat" & Format$(i, "000"), "yyyy.m.d")
            s.UserDate = GetSetting(C_TITLE, "Stamp", "UserDate" & Format$(i, "000"), "")
            s.Lower = Replace(GetSetting(C_TITLE, "Stamp", "Lower" & Format$(i, "000"), "山田"), vbVerticalTab, vbCrLf)
            s.Font = GetSetting(C_TITLE, "Stamp", "Font" & Format$(i, "000"), "ＭＳ ゴシック")
            s.Color = GetSetting(C_TITLE, "Stamp", "Color" & Format$(i, "000"), "&H0")
            s.Line = GetSetting(C_TITLE, "Stamp", "Line" & Format$(i, "000"), C_STAMP_LINE_SINGLE)
            s.size = GetSetting(C_TITLE, "Stamp", "Size" & Format$(i, "000"), "15")
            s.WordArt = GetSetting(C_TITLE, "Stamp", "WordArt" & Format$(i, "000"), C_STAMP_WORDART_ON)
            s.Fill = GetSetting(C_TITLE, "Stamp", "Fill" & Format$(i, "000"), C_STAMP_FILL_OFF)
            s.rect = GetSetting(C_TITLE, "Stamp", "Rect" & Format$(i, "000"), "0")
    
            col.Add s
            
            Set s = Nothing
        Next
    End If
    
    Set getProperty = col
    
End Function
'--------------------------------------------------------------
'　レジストリ設定値取得
'--------------------------------------------------------------
Public Sub setProperty(ByRef col As Collection)

    Dim strBuf As String
    Dim s As StampDatDTO
    Dim lngMax As Long
    Dim i As Long
    
    On Error Resume Next
    DeleteSetting C_TITLE, "Stamp", "Upper"
    DeleteSetting C_TITLE, "Stamp", "Lower"
    DeleteSetting C_TITLE, "Stamp", "DateType"
    DeleteSetting C_TITLE, "Stamp", "DateFormat"
    DeleteSetting C_TITLE, "Stamp", "UserDate"
    DeleteSetting C_TITLE, "Stamp", "Color"
    DeleteSetting C_TITLE, "Stamp", "Font"
    DeleteSetting C_TITLE, "Stamp", "Line"
    DeleteSetting C_TITLE, "Stamp", "Size"
    DeleteSetting C_TITLE, "Stamp", "WordArt"
    DeleteSetting C_TITLE, "Stamp", "Fill"
    DeleteSetting C_TITLE, "Stamp", "Rect"
    
    For i = 0 To col.count - 1
        
        Set s = col(i + 1)
        
        Call SaveSetting(C_TITLE, "Stamp", "Upper" & Format$(i, "000"), Replace(s.Upper, vbCrLf, vbVerticalTab))
        Call SaveSetting(C_TITLE, "Stamp", "Lower" & Format$(i, "000"), Replace(s.Lower, vbCrLf, vbVerticalTab))
        Call SaveSetting(C_TITLE, "Stamp", "DateType" & Format$(i, "000"), s.DateType)
        Call SaveSetting(C_TITLE, "Stamp", "DateFormat" & Format$(i, "000"), s.DateFormat)
        Call SaveSetting(C_TITLE, "Stamp", "UserDate" & Format$(i, "000"), s.UserDate)
        Call SaveSetting(C_TITLE, "Stamp", "Color" & Format$(i, "000"), s.Color)
        Call SaveSetting(C_TITLE, "Stamp", "Font" & Format$(i, "000"), s.Font)
        Call SaveSetting(C_TITLE, "Stamp", "Line" & Format$(i, "000"), s.Line)
        Call SaveSetting(C_TITLE, "Stamp", "Size" & Format$(i, "000"), s.size)
        Call SaveSetting(C_TITLE, "Stamp", "WordArt" & Format$(i, "000"), s.WordArt)
        Call SaveSetting(C_TITLE, "Stamp", "Fill" & Format$(i, "000"), s.Fill)
        Call SaveSetting(C_TITLE, "Stamp", "Rect" & Format$(i, "000"), s.rect)
    
        Set s = Nothing
    Next
    
    Call SaveSetting(C_TITLE, "Stamp", "Count", col.count)
    
End Sub

'--------------------------------------------------------------
'　日付書式設定
'--------------------------------------------------------------
Private Function getFormatDate(ByVal strFormat As String, _
                        ByVal strType As String, _
                        ByVal strUserDate As String)
    
    On Error Resume Next

    If Len(Trim(strFormat)) = 0 Then
        getFormatDate = ""
        Exit Function
    End If
    
    Select Case strType
        Case C_STAMP_DATE_SYSTEM
            getFormatDate = Format(Now, strFormat)
            
        Case C_STAMP_DATE_USER
            If IsDate(strUserDate) Then
                getFormatDate = Format(CDate(strUserDate), strFormat)
            Else
                getFormatDate = ""
            End If
    End Select

End Function

