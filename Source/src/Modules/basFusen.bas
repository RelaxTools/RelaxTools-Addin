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

Private Const IID_IPictureDisp As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"
Private Const PICTYPE_BITMAP As Long = 1
    
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As LongPtr, bitmap As LongPtr) As LongPtr
    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As Long) As LongPtr
    Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal image As LongPtr) As LongPtr
    Private Declare PtrSafe Function GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr) As LongPtr
    Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As LongPtr
    Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, lpiid As Any) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, RefIID As Long, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As LongPtr
    
    Private Type PICTDESC
        Size As Long
        Type As Long
        hPic As LongPtr
        hPal As LongPtr
    End Type
    
    Private Type GdiplusStartupInput
        GdiplusVersion As Long
        DebugEventCallback As LongPtr
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type
    
#Else
    Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, bitmap As Long) As Long
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
    Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal image As Long) As Long
    Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
    Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, lpiid As Any) As Long
    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As Long, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    
    Private Type PICTDESC
      Size As Long
      Type As Long
      hPic As Long
      hPal As Long
    End Type

    Private Type GdiplusStartupInput
      GdiplusVersion As Long
      DebugEventCallback As Long
      SuppressBackgroundThread As Long
      SuppressExternalCodecs As Long
    End Type


#End If

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

    Dim obj As Object
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    If ActiveWorkbook.MultiUserEditing Then
        MsgBox "共有中はシェイプを追加できません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
    Select Case strId
        Case "fsGallery01"
            Set obj = New ShapePasteFusenSquare
        Case "fsGallery02"
            Set obj = New ShapePasteFusenMemo
        Case "fsGallery03"
            Set obj = New ShapePasteFusenCall
        Case "fsGallery04"
            Set obj = New ShapePasteFusenCircle
        Case "fsGallery05"
            Set obj = New ShapePasteFusenPin
        Case "fsGallery06"
            Set obj = New ShapePasteFusenCall2
    End Select

    obj.Id = strId
    obj.No = Index

    obj.Run

    Set obj = Nothing

End Sub
'--------------------------------------------------------------
'　付箋貼り付け
'--------------------------------------------------------------
Sub pasteFusenOrg(ByVal strId As String, ByVal Index As Long)

    Dim r As Shape
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "アクティブなブックが見つかりません。", vbCritical, C_TITLE
        Exit Sub
    End If
    
'    Application.ScreenUpdating = False
    
    On Error GoTo e
    
    Set r = ThisWorkbook.Worksheets(strId).Shapes("shpSquare" & Format(Index, "00"))

    r.Copy
    Call CopyClipboardSleep
 
10    ActiveSheet.Paste
    
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
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Name = strFont
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
    
    Exit Sub
e:
    If Erl = 10 Then
        Resume
    End If
    
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

    Dim obj As ShapePasteFusenSquare
    
    Set obj = New ShapePasteFusenSquare
    
    obj.Id = "fsGallery01"
    obj.No = 1
    
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteSquareY()


    Dim obj As ShapePasteFusenSquare
    
    Set obj = New ShapePasteFusenSquare
    
    obj.Id = "fsGallery01"
    obj.No = 2
    
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteSquareP()
    Dim obj As ShapePasteFusenSquare
    
    Set obj = New ShapePasteFusenSquare
    
    obj.Id = "fsGallery01"
    obj.No = 3
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteSquareB()
    Dim obj As ShapePasteFusenSquare
    
    Set obj = New ShapePasteFusenSquare
    
    obj.Id = "fsGallery01"
    obj.No = 4
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteSquareG()
    Dim obj As ShapePasteFusenSquare
    
    Set obj = New ShapePasteFusenSquare
    
    obj.Id = "fsGallery01"
    obj.No = 5
    obj.Run
    
    Set obj = Nothing

End Sub
Sub beforePasteSquare()
    Dim obj As ShapePasteFusenSquare
    
    Set obj = New ShapePasteFusenSquare
    
    obj.Id = "fsGallery01"
    obj.No = Val(GetSetting(C_TITLE, "Fusen", obj.Id, "2"))
    obj.Run
    
    Set obj = Nothing
End Sub
Sub pasteMemoW()
    Dim obj As ShapePasteFusenMemo
    
    Set obj = New ShapePasteFusenMemo
    
    obj.Id = "fsGallery02"
    obj.No = 1
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteMemoY()
    Dim obj As ShapePasteFusenMemo
    
    Set obj = New ShapePasteFusenMemo
    
    obj.Id = "fsGallery02"
    obj.No = 2
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteMemoP()
    Dim obj As ShapePasteFusenMemo
    
    Set obj = New ShapePasteFusenMemo
    
    obj.Id = "fsGallery02"
    obj.No = 3
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteMemoB()
    Dim obj As ShapePasteFusenMemo
    
    Set obj = New ShapePasteFusenMemo
    
    obj.Id = "fsGallery02"
    obj.No = 4
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteMemoG()
    Dim obj As ShapePasteFusenMemo
    
    Set obj = New ShapePasteFusenMemo
    
    obj.Id = "fsGallery02"
    obj.No = 5
    obj.Run
    
    Set obj = Nothing

End Sub
Sub beforePasteMemo()
    Dim obj As ShapePasteFusenMemo
    
    Set obj = New ShapePasteFusenMemo
    
    obj.Id = "fsGallery02"
    obj.No = Val(GetSetting(C_TITLE, "Fusen", obj.Id, "2"))
    obj.Run
    
    Set obj = Nothing
End Sub
Sub pasteCallW()
    Dim obj As ShapePasteFusenCall
    
    Set obj = New ShapePasteFusenCall
    
    obj.Id = "fsGallery03"
    obj.No = 1
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteCallY()
    Dim obj As ShapePasteFusenCall
    
    Set obj = New ShapePasteFusenCall
    
    obj.Id = "fsGallery03"
    obj.No = 2
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteCallP()
    Dim obj As ShapePasteFusenCall
    
    Set obj = New ShapePasteFusenCall
    
    obj.Id = "fsGallery03"
    obj.No = 3
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteCallB()
    Dim obj As ShapePasteFusenCall
    
    Set obj = New ShapePasteFusenCall
    
    obj.Id = "fsGallery03"
    obj.No = 4
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteCallG()
    Dim obj As ShapePasteFusenCall
    
    Set obj = New ShapePasteFusenCall
    
    obj.Id = "fsGallery03"
    obj.No = 5
    obj.Run
    
    Set obj = Nothing

End Sub
Sub beforePasteCall()
    Dim obj As ShapePasteFusenCall
    
    Set obj = New ShapePasteFusenCall
    
    obj.Id = "fsGallery03"
    obj.No = Val(GetSetting(C_TITLE, "Fusen", obj.Id, "2"))
    obj.Run
    
    Set obj = Nothing
End Sub
Sub pasteCircleW()
    Dim obj As ShapePasteFusenCircle
    
    Set obj = New ShapePasteFusenCircle
    
    obj.Id = "fsGallery04"
    obj.No = 1
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteCircleY()
    Dim obj As ShapePasteFusenCircle
    
    Set obj = New ShapePasteFusenCircle
    
    obj.Id = "fsGallery04"
    obj.No = 2
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteCircleP()
    Dim obj As ShapePasteFusenCircle
    
    Set obj = New ShapePasteFusenCircle
    
    obj.Id = "fsGallery04"
    obj.No = 3
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteCircleB()
    Dim obj As ShapePasteFusenCircle
    
    Set obj = New ShapePasteFusenCircle
    
    obj.Id = "fsGallery04"
    obj.No = 4
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteCircleG()
    Dim obj As ShapePasteFusenCircle
    
    Set obj = New ShapePasteFusenCircle
    
    obj.Id = "fsGallery04"
    obj.No = 5
    obj.Run
    
    Set obj = Nothing

End Sub
Sub beforePasteCircle()
    Dim obj As ShapePasteFusenCircle
    
    Set obj = New ShapePasteFusenCircle
    
    obj.Id = "fsGallery04"
    obj.No = Val(GetSetting(C_TITLE, "Fusen", obj.Id, "2"))
    obj.Run
    
    Set obj = Nothing
End Sub
Sub pastePinW()
    Dim obj As ShapePasteFusenPin
    
    Set obj = New ShapePasteFusenPin
    
    obj.Id = "fsGallery05"
    obj.No = 1
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pastePinY()
    Dim obj As ShapePasteFusenPin
    
    Set obj = New ShapePasteFusenPin
    
    obj.Id = "fsGallery05"
    obj.No = 2
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pastePinP()
    Dim obj As ShapePasteFusenPin
    
    Set obj = New ShapePasteFusenPin
    
    obj.Id = "fsGallery05"
    obj.No = 3
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pastePinB()
    Dim obj As ShapePasteFusenPin
    
    Set obj = New ShapePasteFusenPin
    
    obj.Id = "fsGallery05"
    obj.No = 4
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pastePinG()
    Dim obj As ShapePasteFusenPin
    
    Set obj = New ShapePasteFusenPin
    
    obj.Id = "fsGallery05"
    obj.No = 5
    obj.Run
    
    Set obj = Nothing

End Sub
Sub beforePastePin()
    Dim obj As ShapePasteFusenPin
    
    Set obj = New ShapePasteFusenPin
    
    obj.Id = "fsGallery05"
    obj.No = Val(GetSetting(C_TITLE, "Fusen", obj.Id, "2"))
    obj.Run
    
    Set obj = Nothing
End Sub
Sub pasteLineW()
    Dim obj As ShapePasteFusenCall2
    
    Set obj = New ShapePasteFusenCall2
    
    obj.Id = "fsGallery06"
    obj.No = 1
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteLineY()
    Dim obj As ShapePasteFusenCall2
    
    Set obj = New ShapePasteFusenCall2
    
    obj.Id = "fsGallery06"
    obj.No = 2
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteLineP()
    Dim obj As ShapePasteFusenCall2
    
    Set obj = New ShapePasteFusenCall2
    
    obj.Id = "fsGallery06"
    obj.No = 3
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteLineB()
    Dim obj As ShapePasteFusenCall2
    
    Set obj = New ShapePasteFusenCall2
    
    obj.Id = "fsGallery06"
    obj.No = 4
    obj.Run
    
    Set obj = Nothing

End Sub
Sub pasteLineG()
    Dim obj As ShapePasteFusenCall2
    
    Set obj = New ShapePasteFusenCall2
    
    obj.Id = "fsGallery06"
    obj.No = 5
    obj.Run
    
    Set obj = Nothing

End Sub
Sub beforePasteLine()
    Dim obj As ShapePasteFusenCall2
    
    Set obj = New ShapePasteFusenCall2
    
    obj.Id = "fsGallery06"
    obj.No = Val(GetSetting(C_TITLE, "Fusen", obj.Id, "2"))
    obj.Run
    
    Set obj = Nothing
End Sub
Sub getFusenImage(control As IRibbonControl, ByRef image) ' 画像の設定

    Dim pictureId As String
    
    Select Case control.Id
        Case "beforePasteSquare"
            Select Case Val(GetSetting(C_TITLE, "Fusen", "fsGallery01", "2"))
                 Case 1
                     pictureId = "fusen01w"
                 Case 2
                     pictureId = "fusen01"
                 Case 3
                     pictureId = "fusen01p"
                 Case 4
                     pictureId = "fusen01b"
                 Case 5
                     pictureId = "fusen01g"
             End Select
        Case "beforePasteMemo"
            Select Case Val(GetSetting(C_TITLE, "Fusen", "fsGallery02", "2"))
                 Case 1
                     pictureId = "fusen02w"
                 Case 2
                     pictureId = "fusen02"
                 Case 3
                     pictureId = "fusen02p"
                 Case 4
                     pictureId = "fusen02b"
                 Case 5
                     pictureId = "fusen02g"
             End Select
        Case "beforePasteCall"
            Select Case Val(GetSetting(C_TITLE, "Fusen", "fsGallery03", "2"))
                 Case 1
                     pictureId = "fusen03w"
                 Case 2
                     pictureId = "fusen03"
                 Case 3
                     pictureId = "fusen03p"
                 Case 4
                     pictureId = "fusen03b"
                 Case 5
                     pictureId = "fusen03g"
             End Select
        Case "beforePasteLine"
            Select Case Val(GetSetting(C_TITLE, "Fusen", "fsGallery06", "2"))
                 Case 1
                     pictureId = "fusen06w"
                 Case 2
                     pictureId = "fusen06"
                 Case 3
                     pictureId = "fusen06p"
                 Case 4
                     pictureId = "fusen06b"
                 Case 5
                     pictureId = "fusen06g"
             End Select
        Case "beforePasteCircle"
            Select Case Val(GetSetting(C_TITLE, "Fusen", "fsGallery04", "2"))
                 Case 1
                     pictureId = "fusen04w"
                 Case 2
                     pictureId = "fusen04"
                 Case 3
                     pictureId = "fusen04p"
                 Case 4
                     pictureId = "fusen04b"
                 Case 5
                     pictureId = "fusen04g"
             End Select
        Case "beforePastePin"
            Select Case Val(GetSetting(C_TITLE, "Fusen", "fsGallery05", "2"))
                 Case 1
                     pictureId = "fusen05w"
                 Case 2
                     pictureId = "fusen05"
                 Case 3
                     pictureId = "fusen05p"
                 Case 4
                     pictureId = "fusen05b"
                 Case 5
                     pictureId = "fusen05g"
             End Select
    End Select
    
    Dim file As String
    
    file = rlxGetAppDataFolder & "images\" & pictureId & ".png"
    
    'イメージが見つからなかったら「×」表示する
    If rlxIsFileExists(file) Then
        Set image = LoadImage(file)
    Else
        image = "CancelRequest"
    End If
    
    Call RefreshRibbon
    DoEvents
    
End Sub

' 参考
' 初心者備忘録
' http://www.ka-net.org/ribbon/ri27.html
' ボタンのイメージを外部から読み込む(PNG対応版)
Private Function LoadImage(ByVal strFName As String) As IPicture

    Dim uGdiInput As GdiplusStartupInput
    
#If VBA7 And Win64 Then
    Dim hGdiPlus As LongPtr
    Dim hGdiImage As LongPtr
    Dim hBitmap As LongPtr
#Else
    Dim hGdiPlus As Long
    Dim hGdiImage As Long
    Dim hBitmap As Long
#End If

    uGdiInput.GdiplusVersion = 1&

    If GdiplusStartup(hGdiPlus, uGdiInput) = 0& Then
  
        If GdipCreateBitmapFromFile(StrPtr(strFName), hGdiImage) = 0& Then
        
            Call GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0&)
          
            Dim IID(0 To 3) As Long
            Dim IPic As IPicture
            Dim uPicInfo As PICTDESC
            
            With uPicInfo
              .Size = LenB(uPicInfo)
              .Type = PICTYPE_BITMAP
              .hPic = hBitmap
              .hPal = 0&
            End With
                
            Call IIDFromString(StrPtr(IID_IPictureDisp), IID(0))
            Call OleCreatePictureIndirect(uPicInfo, IID(0), True, LoadImage)
          
            Call GdipDisposeImage(hGdiImage)
          
        End If
        
        Call GdiplusShutdown(hGdiPlus)
    
    End If
  
End Function

