Attribute VB_Name = "basStampBz"
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

Public Const C_STAMP_BZ_RECTANGLE As String = "1"
Public Const C_STAMP_BZ_SQUARE As String = "2"
Public Const C_STAMP_BZ_CIRCLE As String = "3"

Public Const C_STAMP_ROTATE_HOLIZONTAL As String = "1"
Public Const C_STAMP_ROTATE_VERTICAL As String = "2"
'--------------------------------------------------------------
'　レジストリ設定値取得
'--------------------------------------------------------------
Public Function getPropertyBz() As Collection

    Dim strBuf As String
    Dim s As StampBzDTO
    Dim lngMax As Long
    Dim i As Long
    
    Dim col As Collection
    
    Set col = New Collection

    lngMax = GetSetting(C_TITLE, "StampBz", "Count", "-1")
    If lngMax = -1 Then
    
        Set s = New StampBzDTO
        
        s.StampType = C_STAMP_BZ_RECTANGLE
        s.Text = "至急"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = "2014/4/1"
        s.Color = "&H000000FF"
        s.Font = "ＭＳ ゴシック"
        s.Round = "0.15"
        s.Size = "42"
        s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
        s.LineSize = "5"
    
        col.Add s
        
        Set s = Nothing
    
        Set s = New StampBzDTO
        
        s.StampType = C_STAMP_BZ_RECTANGLE
        s.Text = "回覧"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = "2014/4/1"
        s.Color = "&H000000FF"
        s.Font = "ＭＳ ゴシック"
        s.Round = "0.15"
        s.Size = "42"
        s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
        s.LineSize = "5"
    
        col.Add s
        
        Set s = Nothing
        
        Set s = New StampBzDTO
        
        s.StampType = C_STAMP_BZ_RECTANGLE
        s.Text = "見本"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = "2014/4/1"
        s.Color = "&H000000FF"
        s.Font = "ＭＳ ゴシック"
        s.Round = "0.15"
        s.Size = "42"
        s.Rotate = C_STAMP_ROTATE_VERTICAL
        s.LineSize = "5"
    
        col.Add s
        
        Set s = Nothing
        
        Set s = New StampBzDTO
        
        s.StampType = C_STAMP_BZ_RECTANGLE
        s.Text = "社外秘"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = "2014/4/1"
        s.Color = "&H000000FF"
        s.Font = "ＭＳ ゴシック"
        s.Round = "0.15"
        s.Size = "42"
        s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
        s.LineSize = "5"
    
        col.Add s
        
        Set s = Nothing
        
        Set s = New StampBzDTO
        
        s.StampType = C_STAMP_BZ_RECTANGLE
        s.Text = "重要"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = "2014/4/1"
        s.Color = "&H000000FF"
        s.Font = "ＭＳ ゴシック"
        s.Round = "0.15"
        s.Size = "42"
        s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
        s.LineSize = "5"
    
        col.Add s
        
        Set s = Nothing
        
        Set s = New StampBzDTO
        
        s.StampType = C_STAMP_BZ_SQUARE
        s.Text = "取扱" & vbCrLf & "注意"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = "2014/4/1"
        s.Color = "&H000000FF"
        s.Font = "ＭＳ ゴシック"
        s.Round = "0.15"
        s.Size = "42"
        s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
        s.LineSize = "5"
    
        col.Add s
        
        Set s = Nothing
    
        Set s = New StampBzDTO
        
        s.StampType = C_STAMP_BZ_RECTANGLE
        s.Text = "秘密情報" & vbCrLf & "用途後は必ず破棄の事" & vbCrLf & "開示日：$d" & vbCrLf & "開示範囲：社内およびプロジェクト参加各社"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = "2014/4/1"
        s.Color = "&H000000FF"
        s.Font = "ＭＳ ゴシック"
        s.Round = "0.15"
        s.Size = "100"
        s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
        s.LineSize = "5"
    
        col.Add s
        
        Set s = Nothing
    
        Set s = New StampBzDTO
        
        s.StampType = C_STAMP_BZ_CIRCLE
        s.Text = "秘"
        s.DateType = C_STAMP_DATE_SYSTEM
        s.DateFormat = "yyyy.m.d"
        s.UserDate = "2014/4/1"
        s.Color = "&H000000FF"
        s.Font = "ＭＳ ゴシック"
        s.Round = "0.15"
        s.Size = "42"
        s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
        s.LineSize = "5"
    
        col.Add s
        
        Set s = Nothing
    
    Else
        For i = 0 To lngMax - 1
            
            Set s = New StampBzDTO
            
            s.StampType = GetSetting(C_TITLE, "StampBz", "StampType" & Format$(i, "000"), C_STAMP_BZ_RECTANGLE)
            s.Text = Replace(GetSetting(C_TITLE, "StampBz", "Text" & Format$(i, "000"), "至急"), vbVerticalTab, vbCrLf)
            s.DateType = GetSetting(C_TITLE, "StampBz", "DateType" & Format$(i, "000"), C_STAMP_DATE_SYSTEM)
            s.DateFormat = GetSetting(C_TITLE, "StampBz", "DateFormat" & Format$(i, "000"), "yyyy.m.d")
            s.UserDate = GetSetting(C_TITLE, "StampBz", "UserDate" & Format$(i, "000"), "2014/4/1")
            s.Color = GetSetting(C_TITLE, "StampBz", "Color" & Format$(i, "000"), "&H000000FF")
            s.Font = GetSetting(C_TITLE, "StampBz", "Font" & Format$(i, "000"), "ＭＳ ゴシック")
            s.Round = GetSetting(C_TITLE, "StampBz", "Round" & Format$(i, "000"), "0.15")
            s.Size = GetSetting(C_TITLE, "StampBz", "Size" & Format$(i, "000"), "42")
            s.Rotate = GetSetting(C_TITLE, "StampBz", "Rotate" & Format$(i, "000"), C_STAMP_ROTATE_HOLIZONTAL)
            s.LineSize = GetSetting(C_TITLE, "StampBz", "LineSize" & Format$(i, "000"), "5")
        
            col.Add s
            
            Set s = Nothing
        Next
    End If
    
    Set getPropertyBz = col
    
End Function
'--------------------------------------------------------------
'　レジストリ設定値取得
'--------------------------------------------------------------
Public Sub setPropertyBz(ByRef col As Collection)

    Dim strBuf As String
    Dim s As StampBzDTO
    Dim lngMax As Long
    Dim i As Long
    
    On Error Resume Next
    DeleteSetting C_TITLE, "StampBz", "StampType"
    DeleteSetting C_TITLE, "StampBz", "Text"
    DeleteSetting C_TITLE, "StampBz", "DateType"
    DeleteSetting C_TITLE, "StampBz", "DateFormat"
    DeleteSetting C_TITLE, "StampBz", "UserDate"
    DeleteSetting C_TITLE, "StampBz", "Color"
    DeleteSetting C_TITLE, "StampBz", "Font"
    DeleteSetting C_TITLE, "StampBz", "Round"
    DeleteSetting C_TITLE, "StampBz", "Size"
    DeleteSetting C_TITLE, "StampBz", "Rotate"
    DeleteSetting C_TITLE, "StampBz", "LineSize"
    
    For i = 0 To col.count - 1
        
        Set s = col(i + 1)
        
        Call SaveSetting(C_TITLE, "StampBz", "StampType" & Format$(i, "000"), s.StampType)
        Call SaveSetting(C_TITLE, "StampBz", "Text" & Format$(i, "000"), Replace(s.Text, vbCrLf, vbVerticalTab))
        Call SaveSetting(C_TITLE, "StampBz", "DateType" & Format$(i, "000"), s.DateType)
        Call SaveSetting(C_TITLE, "StampBz", "DateFormat" & Format$(i, "000"), s.DateFormat)
        Call SaveSetting(C_TITLE, "StampBz", "UserDate" & Format$(i, "000"), s.UserDate)
        Call SaveSetting(C_TITLE, "StampBz", "Color" & Format$(i, "000"), s.Color)
        Call SaveSetting(C_TITLE, "StampBz", "Font" & Format$(i, "000"), s.Font)
        Call SaveSetting(C_TITLE, "StampBz", "Round" & Format$(i, "000"), s.Round)
        Call SaveSetting(C_TITLE, "StampBz", "Size" & Format$(i, "000"), s.Size)
        Call SaveSetting(C_TITLE, "StampBz", "Rotate" & Format$(i, "000"), s.Rotate)
        Call SaveSetting(C_TITLE, "StampBz", "LineSize" & Format$(i, "000"), s.LineSize)
    
        Set s = Nothing
    Next
    
    Call SaveSetting(C_TITLE, "StampBz", "Count", col.count)
    
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


Public Sub pasteStampBz()

    Dim lngNo As Long
    
    lngNo = GetSetting(C_TITLE, "StampBz", "stampNo", 1)
    Call pasteStampBz2(lngNo)

End Sub
'--------------------------------------------------------------
'　bz貼り付け
'--------------------------------------------------------------
Sub pasteStampBz2(Optional ByVal Index As Variant)

    '設定情報取得
    Dim col As Collection
    Dim datStampBz As StampBzDTO
    Dim r As Worksheet
    
    On Error Resume Next
    
    Set col = getPropertyBz()

    Select Case True
        Case col Is Nothing
            Exit Sub
        Case col.count = 0
            Exit Sub
        Case Else
    End Select

    'Indexが指定されなかった場合
    If IsMissing(Index) Then
        Index = 1
    End If

    Set datStampBz = col(Index)

    Call editStampBz(datStampBz, xlPicture)
    
    Dim sngSize As Single
    
    sngSize = CSng(datStampBz.Size) * C_RASIO
    
    Dim destLeft As Long
    Dim destWidth As Long
    Dim destTop As Long
    Dim destHeight As Long
    
    Application.ScreenUpdating = False
    
    destLeft = ActiveCell.Left
    destWidth = ActiveCell.Width
    destTop = ActiveCell.Top
    destHeight = ActiveCell.Height
 
    ActiveSheet.Paste

    Select Case datStampBz.Rotate
        Case C_STAMP_ROTATE_HOLIZONTAL
            Selection.ShapeRange.Width = sngSize
        Case C_STAMP_ROTATE_VERTICAL
            Selection.ShapeRange.Height = sngSize
    End Select
    
    Selection.ShapeRange.Top = destTop + (destHeight / 2) - (Selection.ShapeRange.Height / 2)
    Selection.ShapeRange.Left = destLeft + (destWidth / 2) - (Selection.ShapeRange.Width / 2)
    
    Selection.Copy
    Application.ScreenUpdating = True

End Sub
'--------------------------------------------------------------
'　bzイメージファイル作成
'--------------------------------------------------------------
Function getImageStampBz(ByVal Index As Long) As StdPicture

    Dim col As Collection
    Dim bz As StampBzDTO
    
    Set getImageStampBz = Nothing
    
    Set col = getPropertyBz()

    Set bz = col(Index)
    
    Set getImageStampBz = editStampBz(bz, xlBitmap)

    Set bz = Nothing
    
End Function
Public Function editStampBz(ByRef datStampBz As StampBzDTO, ByVal lngFormat As Long) As StdPicture

    Dim lngColor As Long
    Dim lngStyle As Long
    Dim sngWeight As Single
    Dim strText As String
    
    Dim strSheet As String
    Dim WS As Worksheet
    
    Dim i As Long
    
    Set editStampBz = Nothing

    Select Case datStampBz.Rotate
        Case C_STAMP_ROTATE_HOLIZONTAL
            Select Case datStampBz.StampType
                Case C_STAMP_BZ_RECTANGLE
                    strSheet = "stampBz1"
                Case C_STAMP_BZ_SQUARE
                    strSheet = "stampBz2"
                Case C_STAMP_BZ_CIRCLE
                    strSheet = "stampBz3"
            End Select
        Case C_STAMP_ROTATE_VERTICAL
            Select Case datStampBz.StampType
                Case C_STAMP_BZ_RECTANGLE
                    strSheet = "stampBz1r"
                Case C_STAMP_BZ_SQUARE
                    strSheet = "stampBz2r"
                Case C_STAMP_BZ_CIRCLE
                    strSheet = "stampBz3r"
            End Select
    End Select
    
    Set WS = ThisWorkbook.Worksheets(strSheet)
    
    Dim strFormat As String
    Dim strType As String
    Dim strUserDate As String
    
    strFormat = datStampBz.DateFormat
    strType = datStampBz.DateType
    
    strUserDate = datStampBz.UserDate
    
    strText = datStampBz.Text
    
    Dim strDate As String
    strDate = getFormatDate(strFormat, strType, strUserDate)
    
    strText = Replace(strText, "$d", strDate)
    
    If InStr(strText, vbCrLf) = 0 Then
        
        With WS.Shapes("shpSquMid").TextFrame2.TextRange
            .Font.NameComplexScript = datStampBz.Font
            .Font.NameFarEast = datStampBz.Font
            .Font.Name = datStampBz.Font
            .Font.Fill.ForeColor.RGB = CLng(datStampBz.Color)
            .Text = strText
        End With
        
        With WS.Shapes("shpSquUp").TextFrame2.TextRange
            .Font.NameComplexScript = datStampBz.Font
            .Font.NameFarEast = datStampBz.Font
            .Font.Name = datStampBz.Font
            .Font.Fill.ForeColor.RGB = CLng(datStampBz.Color)
            .Text = ""
        End With
        
        With WS.Shapes("shpSquDown").TextFrame2.TextRange
            .Font.NameComplexScript = datStampBz.Font
            .Font.NameFarEast = datStampBz.Font
            .Font.Name = datStampBz.Font
            .Font.Fill.ForeColor.RGB = CLng(datStampBz.Color)
            .Text = ""
        End With
    
    Else
    
        Dim strHigh As String
        Dim strLow As String
        Dim lngPos As Long
        
        lngPos = InStr(strText, vbCrLf)
        
        strHigh = Mid$(strText, 1, lngPos - 1)
        strLow = Mid$(strText, lngPos + 2)
    
        With WS.Shapes("shpSquMid").TextFrame2.TextRange
            .Font.NameComplexScript = datStampBz.Font
            .Font.NameFarEast = datStampBz.Font
            .Font.Name = datStampBz.Font
            .Font.Fill.ForeColor.RGB = CLng(datStampBz.Color)
            .Text = ""
        End With
        
        With WS.Shapes("shpSquUp").TextFrame2.TextRange
            .Font.NameComplexScript = datStampBz.Font
            .Font.NameFarEast = datStampBz.Font
            .Font.Name = datStampBz.Font
            .Font.Fill.ForeColor.RGB = CLng(datStampBz.Color)
            .Text = strHigh
        End With
        
        With WS.Shapes("shpSquDown").TextFrame2.TextRange
            .Font.NameComplexScript = datStampBz.Font
            .Font.NameFarEast = datStampBz.Font
            .Font.Name = datStampBz.Font
            .Font.Fill.ForeColor.RGB = CLng(datStampBz.Color)
            .Text = strLow
        End With
    End If

    
    Select Case datStampBz.StampType
        Case C_STAMP_BZ_RECTANGLE, C_STAMP_BZ_SQUARE
            WS.Shapes("shpSquare").Adjustments.Item(1) = CDbl(datStampBz.Round)
    End Select
    
    If CLng(datStampBz.LineSize) = 0 Then
        WS.Shapes("shpSquare").Line.visible = msoFalse
    Else
        WS.Shapes("shpSquare").Line.visible = msoTrue
        WS.Shapes("shpSquare").Line.Weight = CLng(datStampBz.LineSize)
    End If
    
    WS.Shapes("grpSquare").Line.ForeColor.RGB = CLng(datStampBz.Color)
    
    Select Case datStampBz.Rotate
        Case C_STAMP_ROTATE_HOLIZONTAL
            WS.Shapes("grpSquare").Rotation = 0
            
            WS.Shapes("shpSquMid").TextFrame2.Orientation = msoTextOrientationHorizontal
            WS.Shapes("shpSquUp").TextFrame2.Orientation = msoTextOrientationHorizontal
            WS.Shapes("shpSquDown").TextFrame2.Orientation = msoTextOrientationHorizontal
            
            WS.Shapes("shpSquMid").Rotation = 0
            WS.Shapes("shpSquUp").Rotation = 0
            WS.Shapes("shpSquDown").Rotation = 0
            
        Case C_STAMP_ROTATE_VERTICAL
            WS.Shapes("grpSquare").Rotation = 90
            
            WS.Shapes("shpSquMid").TextFrame2.Orientation = msoTextOrientationVerticalFarEast
            WS.Shapes("shpSquUp").TextFrame2.Orientation = msoTextOrientationVerticalFarEast
            WS.Shapes("shpSquDown").TextFrame2.Orientation = msoTextOrientationVerticalFarEast
            
            WS.Shapes("shpSquMid").Rotation = 90
            WS.Shapes("shpSquUp").Rotation = 90
            WS.Shapes("shpSquDown").Rotation = 90
    End Select
    
    Select Case datStampBz.StampType
        Case C_STAMP_BZ_RECTANGLE
            WS.Shapes("grpSquare").Height = 50 * C_RASIO
            WS.Shapes("grpSquare").Width = 150 * C_RASIO
        
        Case Else
            WS.Shapes("grpSquare").Height = 150 * C_RASIO
            WS.Shapes("grpSquare").Width = 150 * C_RASIO
        
    End Select
    
    Dim r As Shape
    
    Set r = WS.Shapes("grpSquare")
    
    If lngFormat = xlBitmap Then
    
        Dim b As Shape
        Dim o As Object
        
        Set b = WS.Shapes("shpBack")
        
        b.Top = r.Top - ((r.Width - r.Height) / 2)
        b.Left = r.Left
        
        b.Height = r.Width
        b.Width = r.Width
        
        b.ZOrder msoSendToBack
        
        Set o = WS.Shapes.Range(Array(r.Name, b.Name)).Group
        
        Set editStampBz = CreatePictureFromClipboard(o)

        
        o.Ungroup

    Else
        r.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    End If
    
    Set WS = Nothing
    
End Function



