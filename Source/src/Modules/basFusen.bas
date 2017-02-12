Attribute VB_Name = "basFusen"
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

Public Const C_FUSEN_DATE_SYSTEM As String = "1"
Public Const C_FUSEN_DATE_USER As String = "2"
'--------------------------------------------------------------
'　画像張付設定画面
'--------------------------------------------------------------
Sub showFusenSetting()

    frmFusen.Show
    
End Sub
'--------------------------------------------------------------
'　付箋管理画面
'--------------------------------------------------------------
Sub searchFusen()

    frmSearchFusen.Show
    
End Sub
'--------------------------------------------------------------
'　付箋設定値取得
'--------------------------------------------------------------
Sub getSettingFusen(ByRef strText As String, ByRef strTag As String, ByRef varPrint As Variant, ByRef strWidth As String, ByRef strHeight As String, ByRef strFormat As String, ByRef strUserDate As String, ByRef strFusenDate As String, ByRef strFont As String, ByRef strSize As String, ByRef varHorizontalAnchor As Variant, ByRef varVerticalAnchor As Variant, ByRef varAutoSize As Variant, ByRef varOverFlow As Variant, ByRef varWordWrap As Variant)

    strTag = GetSetting(C_TITLE, "Fusen", "Tag", "付箋検索用文字列")
    strText = GetSetting(C_TITLE, "Fusen", "Text", "$d" & " " & "$u" & vbCrLf & "【メモをここに入力してください】")
    varPrint = GetSetting(C_TITLE, "Fusen", "PrintObject", True)

    strWidth = GetSetting(C_TITLE, "Fusen", "Width", "7.5")
    strHeight = GetSetting(C_TITLE, "Fusen", "Height", "2.5")
    
    strUserDate = GetSetting(C_TITLE, "Fusen", "UserDate", "")
    strFormat = GetSetting(C_TITLE, "Fusen", "Format", "yyyy.mm.dd hh:mm:ss")
    strFusenDate = GetSetting(C_TITLE, "Fusen", "FusenDate", C_FUSEN_DATE_SYSTEM)
    
    strFont = GetSetting(C_TITLE, "Fusen", "Font", "Meiryo UI")
    strSize = GetSetting(C_TITLE, "Fusen", "Size", "9")
    
    varHorizontalAnchor = GetSetting(C_TITLE, "Fusen", "HorizontalAnchor", 0)
    varVerticalAnchor = GetSetting(C_TITLE, "Fusen", "VerticalAnchor", 0)
    
    varAutoSize = GetSetting(C_TITLE, "Fusen", "AutoSize", False)
    varOverFlow = GetSetting(C_TITLE, "Fusen", "OverFlow", False)
    varWordWrap = GetSetting(C_TITLE, "Fusen", "WordWrap", True)

End Sub
'--------------------------------------------------------------
'　付箋貼り付け
'--------------------------------------------------------------
Sub pasteFusen(ByVal strId As String, ByVal Index As Long)

    Dim r As Shape
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Set r = ThisWorkbook.Worksheets(strId).Shapes("shpSquare" & Format(Index, "00"))

    r.Copy
    Call CopyClipboardSleep
 
    ActiveSheet.Paste
    
    Dim strText As String
    Dim strTag As String
    Dim varPrint As Variant
    
    Dim strWidth  As String
    Dim strHeight  As String
    
    Dim strFormat As String
    Dim strUserDate  As String
    Dim strFusenDate As String
    
    Dim strFont  As String
    Dim strSize  As String
    
    Dim strHorizontalAnchor  As String
    Dim strVerticalAnchor  As String
    
    Dim varAutoSize  As Variant
    Dim varOverFlow As Variant
    Dim varWordWrap As Variant
    
    Call getSettingFusen(strText, strTag, varPrint, strWidth, strHeight, strFormat, strUserDate, strFusenDate, strFont, strSize, strHorizontalAnchor, strVerticalAnchor, varAutoSize, varOverFlow, varWordWrap)
    
    If strId <> "fsGallery05" Then
        Selection.ShapeRange.Width = CDbl(strWidth) * 10 * C_RASIO
        Selection.ShapeRange.Height = CDbl(strHeight) * 10 * C_RASIO
    End If
    
    Selection.ShapeRange.AlternativeText = strTag
    
    Dim strDate As String
    
    strDate = getFormatDate(strFormat, strFusenDate, strUserDate)
    strText = Replace(strText, "$d", strDate)
    strText = Replace(strText, "$u", Application.UserName)
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.name = strFont
    Selection.ShapeRange.TextFrame2.TextRange.Font.NameComplexScript = strFont
    Selection.ShapeRange.TextFrame2.TextRange.Font.NameFarEast = strFont
    Selection.ShapeRange.TextFrame2.TextRange.Font.NameAscii = strFont
    Selection.ShapeRange.TextFrame2.TextRange.Font.NameOther = strFont

    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = CDbl(strSize)
    Selection.ShapeRange.TextFrame2.TextRange.Text = strText
    
    If strId <> "fsGallery05" Then
        Select Case strVerticalAnchor
            Case "0"
                Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorTop
            Case "1"
                Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
            Case "2"
                Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorBottom
        End Select
            
        Select Case strHorizontalAnchor
            Case "0"
                Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
            Case "1"
                Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            Case "2"
                Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
        End Select
    End If
    
    Selection.PrintObject = CBool(varPrint)
    
    If strId <> "fsGallery05" Then
        If CBool(varAutoSize) Then
            Selection.ShapeRange.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        Else
            Selection.ShapeRange.TextFrame2.AutoSize = msoAutoSizeNone
        End If
    End If
    
#If VBA7 Then
    If strId <> "fsGallery05" Then
        If CBool(varOverFlow) Then
            Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
            Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
        Else
            Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowClip
            Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartVerticalOverflowClip
        End If
    End If
#End If

    If strId <> "fsGallery05" Then
        Selection.ShapeRange.TextFrame2.WordWrap = CBool(varWordWrap)
    End If
    

    Application.ScreenUpdating = True

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
        Case C_FUSEN_DATE_SYSTEM
            getFormatDate = Format(Now, strFormat)
            
        Case C_FUSEN_DATE_USER
            If IsDate(strUserDate) Then
                getFormatDate = Format(CDate(strUserDate), strFormat)
            Else
                getFormatDate = ""
            End If
    End Select

End Function
'--------------------------------------------------------------
'　イメージファイル作成
'--------------------------------------------------------------
Function getImageFusen(ByVal strId As String, ByVal Index As Long) As StdPicture
    
    On Error Resume Next
    
    Set getImageFusen = Nothing
    
    Dim r As Shape
    Set r = ThisWorkbook.Worksheets(strId).Shapes("shpSquare" & Format(Index, "00"))
    
    Set getImageFusen = CreatePictureFromClipboard(r)
    
End Function
'--------------------------------------------------------------
'　付箋貼り付けのショートカット用
'--------------------------------------------------------------
Sub pasteSquareW()
    pasteFusen "fsGallery01", 1
End Sub
Sub pasteSquareY()
    pasteFusen "fsGallery01", 2
End Sub
Sub pasteSquareP()
    pasteFusen "fsGallery01", 3
End Sub
Sub pasteSquareB()
    pasteFusen "fsGallery01", 4
End Sub
Sub pasteSquareG()
    pasteFusen "fsGallery01", 5
End Sub
Sub pasteMemoW()
    pasteFusen "fsGallery02", 1
End Sub
Sub pasteMemoY()
    pasteFusen "fsGallery02", 2
End Sub
Sub pasteMemoP()
    pasteFusen "fsGallery02", 3
End Sub
Sub pasteMemoB()
    pasteFusen "fsGallery02", 4
End Sub
Sub pasteMemoG()
    pasteFusen "fsGallery02", 5
End Sub
Sub pasteCallW()
    pasteFusen "fsGallery03", 1
End Sub
Sub pasteCallY()
    pasteFusen "fsGallery03", 2
End Sub
Sub pasteCallP()
    pasteFusen "fsGallery03", 3
End Sub
Sub pasteCallB()
    pasteFusen "fsGallery03", 4
End Sub
Sub pasteCallG()
    pasteFusen "fsGallery03", 5
End Sub
Sub pasteCircleW()
    pasteFusen "fsGallery04", 1
End Sub
Sub pasteCircleY()
    pasteFusen "fsGallery04", 2
End Sub
Sub pasteCircleP()
    pasteFusen "fsGallery04", 3
End Sub
Sub pasteCircleB()
    pasteFusen "fsGallery04", 4
End Sub
Sub pasteCircleG()
    pasteFusen "fsGallery04", 5
End Sub
Sub pastePinW()
    pasteFusen "fsGallery05", 1
End Sub
Sub pastePinY()
    pasteFusen "fsGallery05", 2
End Sub
Sub pastePinP()
    pasteFusen "fsGallery05", 3
End Sub
Sub pastePinB()
    pasteFusen "fsGallery05", 4
End Sub
Sub pastePinG()
    pasteFusen "fsGallery05", 5
End Sub
Sub pasteLineW()
    pasteFusen "fsGallery06", 1
End Sub
Sub pasteLineY()
    pasteFusen "fsGallery06", 2
End Sub
Sub pasteLineP()
    pasteFusen "fsGallery06", 3
End Sub
Sub pasteLineB()
    pasteFusen "fsGallery06", 4
End Sub
Sub pasteLineG()
    pasteFusen "fsGallery06", 5
End Sub

