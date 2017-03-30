Attribute VB_Name = "basStampMitome"
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

Public Const C_STAMP_MITOME_NORMAL As Long = 1
Public Const C_STAMP_MITOME_FILE As Long = 2
'--------------------------------------------------------------
'　認印設定画面
'--------------------------------------------------------------
Sub showMitome()

    frmStampMitome.Show
End Sub
Public Function editStampMitome(ByRef s As StampMitomeDTO, ByVal lngFormat As Long) As StdPicture
   
    Dim r As Shape
    Dim WS As Worksheet
    
    Set editStampMitome = Nothing
    
    Set WS = ThisWorkbook.Worksheets("stampEx")
    
    Select Case s.Rotate
        Case C_STAMP_ROTATE_HOLIZONTAL
            Set r = WS.Shapes("shpMitomeR")
        Case C_STAMP_ROTATE_VERTICAL
            Set r = WS.Shapes("shpMitome")
    End Select
    
    r.TextFrame2.TextRange.Text = s.Text
    
    With r.TextFrame2.TextRange.Font
        .NameComplexScript = s.Font
        .NameFarEast = s.Font
        .Name = s.Font
        .Strikethrough = False
        .Superscript = False
        .Subscript = False

        .Fill.ForeColor.RGB = CLng(s.Color)
    
    End With
    
    If s.Fill = C_STAMP_FILL_ON Then
        r.Fill.visible = True
        r.Fill.ForeColor.RGB = vbWhite
    Else
        r.Fill.visible = False
    End If
    
    r.Line.ForeColor.RGB = CLng(s.Color)
    
    If CLng(s.LineSize) > 0 Then
        r.Line.Weight = CLng(s.LineSize)
        r.Line.visible = True
    Else
        r.Line.visible = False
    End If

    Select Case s.Line
        Case C_STAMP_LINE_SINGLE
        
            r.AutoShapeType = msoShapeOval
            r.Height = r.Width
        
        Case C_STAMP_LINE_DOUBLE
        
            
            r.AutoShapeType = msoShapeOval
            
            Select Case s.Rotate
                Case C_STAMP_ROTATE_HOLIZONTAL
                    r.Height = r.Width
                Case C_STAMP_ROTATE_VERTICAL
                    r.Height = r.Width * 0.8
            End Select
            
        Case C_STAMP_LINE_BOLD
        
            r.AutoShapeType = msoShapeRoundedRectangle
            r.Adjustments.Item(1) = CDbl(s.Round)
            r.Height = r.Width
    
    End Select
    
    r.Rotation = getRect(s.rect)
    
    If lngFormat = xlBitmap Then
    
        Dim lngWidth As Long
        Dim lngHeight As Long
        Dim b As Shape
        Dim o As Object
        
        Set b = WS.Shapes("shpBack")
        
        b.Top = r.Top - ((r.Width - r.Height) / 2)
        b.Left = r.Left
        b.Height = r.Width
        b.Width = r.Width
        
        b.ZOrder msoSendToBack
        
        Set o = WS.Shapes.Range(Array(r.Name, b.Name)).Group
        
        Set editStampMitome = CreatePictureFromClipboard(o)
        
        o.Ungroup

    Else
        r.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        Call CopyClipboardSleep
    End If
    
    Set WS = Nothing
    
End Function
'--------------------------------------------------------------
'　bzイメージファイル作成
'--------------------------------------------------------------
Function getImageStampMitome(ByVal Index As Long) As StdPicture

    '設定情報取得
    Dim Col As Collection
    Dim bz As StampMitomeDTO
    
    Set getImageStampMitome = Nothing
    
    Set Col = getPropertyMitome()
    Set bz = Col(Index)
    
    If bz.StampType = C_STAMP_MITOME_NORMAL Then
   
        Set getImageStampMitome = editStampMitome(bz, xlBitmap)
    Else
        If Not rlxIsFileExists(bz.FilePath) Then
            Exit Function
        End If

        Dim o As Object
        Set o = ThisWorkbook.Worksheets("stampEx").Pictures.Insert(bz.FilePath)
        Set getImageStampMitome = CreatePictureFromClipboard(o)
        o.Delete
        
    End If

    Set bz = Nothing
    
End Function
Public Sub MitomePaste()

    Dim lngNo As Long

    lngNo = GetSetting(C_TITLE, "StampMitome", "stampNo", 1)
    Call MitomePaste2(lngNo)

End Sub
'--------------------------------------------------------------
'　データ印貼り付け
'--------------------------------------------------------------
Sub MitomePaste2(Optional ByVal Index As Variant)

    '設定情報取得
    Dim Col As Collection
    Dim s As StampMitomeDTO
    Dim ss As Range
    
    On Error Resume Next
    
    Set Col = getPropertyMitome()
    
    Select Case True
        Case Col Is Nothing
            Exit Sub
        Case Col.count = 0
            Exit Sub
        Case Else
    End Select

    'Indexが指定されなかった場合
    If IsMissing(Index) Then
        Index = 1
    End If
    
    If rlxCheckSelectRange() = False Then
        MsgBox "選択範囲が見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If

    If GetSetting(C_TITLE, "StampMitome", "Confirm", False) Then
    Else
        If Selection.CountLarge > 1 And Selection.CountLarge <> Selection(1, 1).MergeArea.count Then
            If MsgBox("複数セル選択されています。すべてのセルに張り付けますがよろしいですか？", vbQuestion + vbYesNo, C_TITLE) <> vbYes Then
                Exit Sub
            End If
        End If
    End If

    Set s = Col(Index)

    If s.StampType = C_STAMP_MITOME_NORMAL Then
        Call editStampMitome(s, xlPicture)
    Else
        With ActiveSheet.Pictures.Insert(s.FilePath)
            .CopyPicture xlScreen, xlPicture
            Call CopyClipboardSleep
            .Delete
        End With
    End If
    
    Dim sngSize As Single
    
    sngSize = CSng(s.Size) * C_RASIO
    
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
                destWidth = ss.MergeArea.Width
                destTop = ss.MergeArea.Top
                destHeight = ss.MergeArea.Height

                ActiveSheet.Paste
                
                If Selection.ShapeRange.Height > Selection.ShapeRange.Width Then
                    Selection.ShapeRange.Height = sngSize
                Else
                    Selection.ShapeRange.Width = sngSize
                End If
                
                Selection.ShapeRange.Top = destTop + (destHeight / 2) - (Selection.ShapeRange.Height / 2)
                Selection.ShapeRange.Left = destLeft + (destWidth / 2) - (Selection.ShapeRange.Width / 2)
                Selection.ShapeRange.Rotation getRect(s.rect)
                
            End If
        End If
    Next
    
    Selection.Copy
    Application.ScreenUpdating = True

End Sub
Function getRect(rect As String) As Single
    Dim a As Single
    Dim b As Single
    a = (Val(rect) / 100) * -1
    
    b = (a * 180)
    If b < 0 Then
        b = b + 360
    End If
        getRect = b + 90
End Function

'--------------------------------------------------------------
'　レジストリ設定値取得
'--------------------------------------------------------------
Public Function getPropertyMitome() As Collection

    Dim strBuf As String
    Dim s As StampMitomeDTO
    Dim lngMax As Long
    Dim i As Long
    
    Dim Col As Collection
    
    Set Col = New Collection

    lngMax = GetSetting(C_TITLE, "StampMitome", "Count", "-1")
    If lngMax = -1 Then
    
        Set s = New StampMitomeDTO
        
        s.StampType = C_STAMP_MITOME_NORMAL
        s.Text = "山田"
        s.Font = "ＭＳ ゴシック"
        s.Color = "&H0"
        s.Line = C_STAMP_LINE_SINGLE
        s.Size = "10.5"
        s.FilePath = ""
        s.LineSize = "10"
        s.Round = "0.10"
        s.Rotate = C_STAMP_ROTATE_VERTICAL
        s.Fill = C_STAMP_FILL_OFF
        s.rect = "0"
    
        Col.Add s
        
        Set s = Nothing
        
        Set s = New StampMitomeDTO
        
        s.StampType = C_STAMP_MITOME_NORMAL
        s.Text = "田中"
        s.Font = "ＭＳ ゴシック"
        s.Color = "&HFF"
        s.Line = C_STAMP_LINE_DOUBLE
        s.Size = "10.5"
        s.FilePath = ""
        s.LineSize = "10"
        s.Round = "0.10"
        s.Rotate = C_STAMP_ROTATE_VERTICAL
        s.Fill = C_STAMP_FILL_OFF
        s.rect = "0"
    
        Col.Add s
        
        Set s = Nothing
    
        Set s = New StampMitomeDTO
        
        s.StampType = C_STAMP_MITOME_NORMAL
        s.Text = "株式会社" & vbCrLf & "日本工業" & vbCrLf & "企画之印"
        s.Font = "ＭＳ ゴシック"
        s.Color = "&HFF"
        s.Line = C_STAMP_LINE_BOLD
        s.Size = "30"
        s.FilePath = ""
        s.LineSize = "10"
        s.Round = "0.10"
        s.Rotate = C_STAMP_ROTATE_VERTICAL
        s.Fill = C_STAMP_FILL_OFF
        s.rect = "0"
    
        Col.Add s
        
        Set s = Nothing
    
    Else
        For i = 0 To lngMax - 1
            
            Set s = New StampMitomeDTO
        
            s.StampType = GetSetting(C_TITLE, "StampMitome", "StampType" & Format$(i, "000"), C_STAMP_MITOME_NORMAL)
            s.Text = Replace(GetSetting(C_TITLE, "StampMitome", "Text" & Format$(i, "000"), "田中"), vbVerticalTab, vbCrLf)
            s.Font = GetSetting(C_TITLE, "StampMitome", "Font" & Format$(i, "000"), "ＭＳ ゴシック")
            s.Color = GetSetting(C_TITLE, "StampMitome", "Color" & Format$(i, "000"), "&HFF")
            s.Line = GetSetting(C_TITLE, "StampMitome", "Line" & Format$(i, "000"), C_STAMP_LINE_SINGLE)
            s.Size = GetSetting(C_TITLE, "StampMitome", "Size" & Format$(i, "000"), "15")
            s.FilePath = GetSetting(C_TITLE, "StampMitome", "FilePath" & Format$(i, "000"), "")
            s.LineSize = GetSetting(C_TITLE, "StampMitome", "LIneSize" & Format$(i, "000"), "10")
            s.Round = GetSetting(C_TITLE, "StampMitome", "Round" & Format$(i, "000"), "0.10")
            s.Rotate = GetSetting(C_TITLE, "StampMitome", "Rotate" & Format$(i, "000"), C_STAMP_ROTATE_VERTICAL)
            s.Fill = GetSetting(C_TITLE, "StampMitome", "Fill" & Format$(i, "000"), C_STAMP_FILL_OFF)
            s.rect = GetSetting(C_TITLE, "StampMitome", "Rect" & Format$(i, "000"), "0")
    
            Col.Add s
            
            Set s = Nothing
        Next
    End If
    
    Set getPropertyMitome = Col
    
End Function
'--------------------------------------------------------------
'　レジストリ設定値取得
'--------------------------------------------------------------
Public Sub setPropertyMitome(ByRef Col As Collection)

    Dim strBuf As String
    Dim s As StampMitomeDTO
    Dim lngMax As Long
    Dim i As Long
    
    On Error Resume Next
    DeleteSetting C_TITLE, "StampMitome", "StampType"
    DeleteSetting C_TITLE, "StampMitome", "Text"
    DeleteSetting C_TITLE, "StampMitome", "Color"
    DeleteSetting C_TITLE, "StampMitome", "Font"
    DeleteSetting C_TITLE, "StampMitome", "Line"
    DeleteSetting C_TITLE, "StampMitome", "Size"
    DeleteSetting C_TITLE, "StampMitome", "FilePath"
    DeleteSetting C_TITLE, "StampMitome", "LineSize"
    DeleteSetting C_TITLE, "StampMitome", "Round"
    DeleteSetting C_TITLE, "StampMitome", "Rotate"
    DeleteSetting C_TITLE, "StampMitome", "Fill"
    DeleteSetting C_TITLE, "StampMitome", "Rect"
    
    For i = 0 To Col.count - 1
        
        Set s = Col(i + 1)
        
        Call SaveSetting(C_TITLE, "StampMitome", "StampType" & Format$(i, "000"), s.StampType)
        Call SaveSetting(C_TITLE, "StampMitome", "Text" & Format$(i, "000"), Replace(s.Text, vbCrLf, vbVerticalTab))
        Call SaveSetting(C_TITLE, "StampMitome", "Color" & Format$(i, "000"), s.Color)
        Call SaveSetting(C_TITLE, "StampMitome", "Font" & Format$(i, "000"), s.Font)
        Call SaveSetting(C_TITLE, "StampMitome", "Line" & Format$(i, "000"), s.Line)
        Call SaveSetting(C_TITLE, "StampMitome", "Size" & Format$(i, "000"), s.Size)
        Call SaveSetting(C_TITLE, "StampMitome", "FilePath" & Format$(i, "000"), s.FilePath)
        Call SaveSetting(C_TITLE, "StampMitome", "LineSize" & Format$(i, "000"), s.LineSize)
        Call SaveSetting(C_TITLE, "StampMitome", "Round" & Format$(i, "000"), s.Round)
        Call SaveSetting(C_TITLE, "StampMitome", "Rotate" & Format$(i, "000"), s.Rotate)
        Call SaveSetting(C_TITLE, "StampMitome", "Fill" & Format$(i, "000"), s.Fill)
        Call SaveSetting(C_TITLE, "StampMitome", "Rect" & Format$(i, "000"), s.rect)
        
        Set s = Nothing
    Next
    
    Call SaveSetting(C_TITLE, "StampMitome", "Count", Col.count)
    
End Sub

Sub FilePaste()
    MsgBox "この機能は「印鑑」機能に統合されました。", vbOKOnly + vbInformation, C_TITLE
End Sub
