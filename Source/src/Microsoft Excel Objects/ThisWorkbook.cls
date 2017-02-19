VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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

Private WithEvents XL As Excel.Application
Attribute XL.VB_VarHelpID = -1
Private WithEvents XL_LINE As Excel.Application
Attribute XL_LINE.VB_VarHelpID = -1


Private mScrollRow() As Long
Private mScrollCol() As Long
Private mZoom() As Long
Private mActiveSheet As String
Private mblnEventCancel As Boolean
Private mSelection() As Object



Private Const C_HOLIZONTAL_BAR_NAME As String = "rlxHolizontalBar"
Private Const C_VERTICAL_BAR_NAME As String = "rlxVerticalBar"
Private Const C_XY_BAR_NAME As String = "rlxXYBar"
Private mWS As Worksheet
Public Sub setIRibbon(ByRef r As IRibbonUI)
    Set XL = Excel.Application
End Sub

'--------------------------------------------------------------
'　ブックオープンイベント
'--------------------------------------------------------------
Private Sub Workbook_Open()
    
    Logger.LogBegin "Workbook_Open"
    
    '情報ログの出力
    Call DispInformation
    
    'ショートカットキー設定
    Call setShortCutKey
    
    Call createAllSectionObject
    Call createAllItemObject
    
    Set mColSection = rlxInitSectionSetting()
    
    Logger.LogFinish "Workbook_Open"

End Sub
Private Sub DispInformation()

    Dim strVer As String
    Dim strTitle As String

    strTitle = ThisWorkbook.BuiltinDocumentProperties("Title").Value
    strVer = ThisWorkbook.BuiltinDocumentProperties("Comments").Value
    
    
    Dim strBuf As String
    Dim i As Long
    Dim obj As Object

    strBuf = strTitle
    Dim s() As String
    s = Split(strVer, vbLf)
    strBuf = strBuf & " " & s(0)
    
    Logger.LogInfo strBuf

    strBuf = "Microsoft "
    Select Case True
        Case InStr(Application.OperatingSystem, "5.00") > 0
            strBuf = strBuf & "Windows 2000"
        Case InStr(Application.OperatingSystem, "5.01") > 0
            strBuf = strBuf & "Windows XP"
        Case InStr(Application.OperatingSystem, "6.00") > 0
            strBuf = strBuf & "Windows Vista"
        Case InStr(Application.OperatingSystem, "6.01") > 0
            strBuf = strBuf & "Windows 7"
        Case InStr(Application.OperatingSystem, "6.02") > 0
            strBuf = strBuf & "Windows 8 or 8.1"
        Case Else
            strBuf = strBuf & "Windows 10 or Later"
    End Select
    If Isx64 Then
        strBuf = strBuf & " (64bit)"
    Else
        strBuf = strBuf & " (32bit)"
    End If
    Logger.LogInfo strBuf

    strBuf = "Microsoft Excel "

    Select Case Val(Application.Version)
        Case Is = 0
            strBuf = strBuf & "不明"
        Case Is <= 11
            strBuf = strBuf & "2003以前"
        Case 12
            strBuf = strBuf & "2007"
        Case 14
            strBuf = strBuf & "2010"
        Case 15
            strBuf = strBuf & "2013"
        Case 16
            strBuf = strBuf & "2016"
        Case Else
            strBuf = strBuf & "2016より未来のバージョン"
    End Select
    strBuf = strBuf & " Build " & Application.Build
#If Win64 Then
    strBuf = strBuf & " (64bit)"
#Else
    strBuf = strBuf & " (32bit)"
#End If
    Logger.LogInfo strBuf

End Sub
Private Function Isx64() As Boolean

    On Error GoTo xp

    Dim colItems As Object
    Dim itm As Object
    Dim ret As Boolean

    ret = False '初期化

    Set colItems = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery("Select * From Win32_OperatingSystem")

    For Each itm In colItems
        If InStr(itm.OSArchitecture, "64") Then
            ret = True
            Exit For
        End If
    Next

    Isx64 = ret

    Exit Function
xp:
    Isx64 = False

End Function
'--------------------------------------------------------------
'　ブッククローズ前イベント
'--------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error Resume Next
    
    Logger.LogBegin "Workbook_BeforeClose"
    
    Call removeShortCutKey
    If XL_LINE Is Nothing Then
    Else
        Call deleteCrossLine
    End If
    If mScreenEnable Then
        Unload frmScreenShot
    End If
    Call DeleteTemporaryFile

    Logger.LogFinish "Workbook_BeforeClose"

End Sub
Private Sub DeleteTemporaryFile()

    Dim FS As Object
    Dim strTmpBook As String
    
    On Error Resume Next

    Set FS = CreateObject("Scripting.FileSystemObject")
    
    strTmpBook = rlxGetTempFolder() & "*.*"

    
    FS.DeleteFile strTmpBook, True


    Set FS = Nothing

End Sub


Private Sub XL_WindowResize(ByVal WB As Workbook, ByVal wn As Window)
    Call RefreshRibbon
End Sub

'Private Sub XL_WorkbookActivate(ByVal WB As Workbook)
'    Call RefreshRibbon
'End Sub

'Private Sub XL_WorkbookBeforeClose(ByVal WB As Workbook, Cancel As Boolean)
'    Call holdBookClose(WB)
'End Sub

'--------------------------------------------------------------
'　ブック保存前イベント
'--------------------------------------------------------------
Private Sub XL_WorkbookBeforeSave(ByVal WB As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Dim strName As String
    Dim strFolder As String
    Dim strFile As String
    Dim strList() As String
    Dim blnFind As Boolean
    Dim i As Long
    Dim strSourceFile As String
    Dim strDestFile As String
    Dim exec As Boolean
    Dim lngGen As Long
    
    Dim blnScreenUpdating As Boolean
    Dim blnDisplayAlerts As Boolean
    
    'イベントのキャンセル
    If mblnEventCancel Then
        Exit Sub
    End If
    
    If Val(Application.Version) > C_EXCEL_VERSION_2007 Then
    Else
        Exit Sub
    End If
    
    '-------------------------------------------------
    ' 簡易世代管理
    '-------------------------------------------------
    exec = CBool(GetSetting(C_TITLE, "Backup", "Check", False))
    If rlxIsExcelFile(WB.FullName) And exec Then
    Else
        Exit Sub
    End If
    
    blnScreenUpdating = Application.ScreenUpdating
    blnDisplayAlerts = Application.DisplayAlerts
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    strFile = GetSetting(C_TITLE, "Backup", "FileList", "")
    strList = Split(strFile, vbTab)
    
    blnFind = False
    For i = 0 To UBound(strList)
'        If InStr(LCase(Application.ActiveWorkbook.FullName), LCase(strList(i))) > 0 Then
        If InStr(LCase(WB.FullName), LCase(strList(i))) > 0 Then
             blnFind = True
            Exit For
        End If
    Next
    If Not blnFind Then
        GoTo pass
    End If
    
    
    strName = WB.FullName
    
    'パス存在チェック
    strFolder = rlxGetFullpathFromPathName(strName)
    If Len(Trim(strFolder)) = 0 Then
        GoTo pass
    End If

    'ファイル存在チェック
    If Not rlxIsFileExists(strName) Then
        GoTo pass
    End If
    
    strFolder = GetSetting(C_TITLE, "Backup", "Folder", "")
    Dim strFullName As String
    Dim strPath As String
    
    lngGen = Val(GetSetting(C_TITLE, "Backup", "Gen", "99"))
    
    For i = lngGen To 0 Step -1
    
        If Len(Trim(strFolder)) = 0 Then
            strFullName = WB.FullName
        Else
            strFullName = rlxAddFileSeparator(strFolder) & rlxGetFullpathFromFileName(WB.FullName)
        End If
    
        strSourceFile = strFullName & "." & Format$(i, "000")
        If rlxIsFileExists(strSourceFile) Then
            If i = lngGen Then
                Kill strSourceFile
            Else
                strDestFile = strFullName & "." & Format$(i + 1, "000")
                Name strSourceFile As strDestFile
            End If
        End If
    
    Next

pass:
    
    Application.ScreenUpdating = blnScreenUpdating
    Application.DisplayAlerts = blnDisplayAlerts
    
End Sub

'--------------------------------------------------------------
'　ブック保存後イベント
'--------------------------------------------------------------
Private Sub XL_WorkbookAfterSave(ByVal WB As Workbook, ByVal Success As Boolean)

    Dim a1exec As Boolean
    Dim backexec As Boolean
    Dim Target As Boolean
    Dim strFile As String
    Dim strList() As String
    Dim blnFind As Boolean
    Dim i As Long
    
    Dim blnScreenUpdating As Boolean
    Dim blnDisplayAlerts As Boolean
    
    'イベントのキャンセル
    If mblnEventCancel Or Success = False Then
        Exit Sub
    End If
    
    backexec = GetSetting(C_TITLE, "Backup", "Check", False)
    
    '-------------------------------------------------
    ' 簡易世代管理
    '-------------------------------------------------
    If rlxIsExcelFile(WB.FullName) And backexec Then
    
        blnDisplayAlerts = Application.DisplayAlerts
        blnScreenUpdating = Application.ScreenUpdating
        
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
    
        strFile = GetSetting(C_TITLE, "Backup", "FileList", "")
        strList = Split(strFile, vbTab)
        
        blnFind = False
        For i = 0 To UBound(strList)
            If InStr(LCase(Application.ActiveWorkbook.FullName), LCase(strList(i))) > 0 Then
                 blnFind = True
                Exit For
            End If
        Next
        If blnFind Then
            Dim strFolder As String
            Dim strFullName As String
            strFolder = GetSetting(C_TITLE, "Backup", "Folder", "")
            If Len(Trim(strFolder)) = 0 Then
                strFullName = WB.FullName
            Else
                strFullName = rlxAddFileSeparator(strFolder) & rlxGetFullpathFromFileName(WB.FullName)
            End If
            mblnEventCancel = True
            WB.SaveCopyAs strFullName & ".000"
            mblnEventCancel = False
        
        End If
        
        Application.DisplayAlerts = blnDisplayAlerts
        Application.ScreenUpdating = blnScreenUpdating
    
    End If

End Sub
Public Sub enableCrossLine()
    
    On Error Resume Next
    
    Call addCrossLine
    
    Set XL_LINE = Excel.Application
    
End Sub
Private Sub addCrossLine()

    Dim s As Shape
    Dim b As Boolean
    Dim z As Single
    Dim r As Long
    Dim c As Long
    Dim a As Boolean
    
    If ActiveWorkbook.ActiveSheet Is Nothing Then
    Else
        
        a = Application.ScreenUpdating
        Application.ScreenUpdating = False
        
        b = ActiveWorkbook.Saved
        Set mWS = ActiveWorkbook.ActiveSheet
        
        z = ActiveWindow.Zoom
        c = ActiveWindow.ScrollColumn
        r = ActiveWindow.ScrollRow
        
        ActiveWindow.Zoom = 100
        
        Call deleteCrossLine
        
        Dim lngFillVisible As Long
        Dim lngFillColor As Long
        Dim dblFillTransparency As Double
        Dim lngLineVisible As Long
        Dim lngLineColor As Long
        Dim sngLineWeight As Single
        Dim strOnAction As String
        Dim lngType As Long
        Dim blnGuid As Boolean
        Dim lngFontColor As Long
        Dim blnEdit As Boolean
    
        Call getCrossLineSetting(lngType, lngFillVisible, lngFillColor, dblFillTransparency, lngLineVisible, lngLineColor, sngLineWeight, strOnAction, blnGuid, lngFontColor, blnEdit)
        
        If lngType And C_HOLIZON Then
            Set s = mWS.Shapes.AddShape(msoShapeRectangle, 10, 10, 10, 10)
            
            s.name = C_HOLIZONTAL_BAR_NAME
            s.Line.visible = msoFalse
            
            s.Select
            Selection.PrintObject = msoFalse
            Selection.Placement = xlFreeFloating
        
            With s.Fill
                .visible = msoTrue
                .ForeColor.RGB = RGB(0, 176, 80)
                .Transparency = 0.5
                .Solid
            End With
            
            Set s = Nothing
        End If
        
        If lngType And C_VERTICAL Then
            Set s = mWS.Shapes.AddShape(msoShapeRectangle, 20, 20, 10, 10)
        
            s.name = C_VERTICAL_BAR_NAME
            s.Line.visible = msoFalse
            
            s.Select
            Selection.PrintObject = msoFalse
            Selection.Placement = xlFreeFloating
        
            With s.Fill
                .visible = msoTrue
                .ForeColor.RGB = RGB(0, 176, 80)
                .Transparency = 0.5
                .Solid
            End With
            
            Set s = Nothing
        End If
        
        If blnGuid Then
            Set s = mWS.Shapes.AddShape(msoShapeRectangle, 20, 20, 10, 10)
        
            s.name = C_XY_BAR_NAME
            s.Line.visible = msoFalse
            
            s.Select
            Selection.PrintObject = msoFalse
            Selection.Placement = xlFreeFloating
        
            With s.Fill
                .visible = msoTrue
                .ForeColor.RGB = RGB(0, 176, 80)
                .Transparency = 0.5
                .Solid
            End With
            
            Set s = Nothing
        End If
        
        Call XL_LINE_SheetSelectionChange(mWS, ActiveCell)
        
        ActiveCell.Select
        ActiveWorkbook.Saved = b
        
        ActiveWindow.Zoom = z
        ActiveWindow.ScrollColumn = c
        ActiveWindow.ScrollRow = r
        
        Application.ScreenUpdating = a
    
    End If

End Sub

Public Sub disableCrossLine()

    Set XL_LINE = Nothing
    
    If mWS Is Nothing Then
    Else
        Call deleteCrossLine
    End If
    Set mWS = Nothing
End Sub

Private Sub deleteCrossLine()
    On Error Resume Next
    
    Dim b As Boolean
    
    b = mWS.Parent.Saved
    
    mWS.Shapes(C_HOLIZONTAL_BAR_NAME).Delete
    mWS.Shapes(C_VERTICAL_BAR_NAME).Delete
    mWS.Shapes(C_XY_BAR_NAME).Delete
    
    mWS.Parent.Saved = b

End Sub

Private Sub XL_LINE_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)

    On Error Resume Next

    Const C_MAX As Long = 50000
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim r As Range
    Dim b As Boolean
    Dim s As Boolean
    
    Static blnCall As Boolean

    If blnCall Then
        Exit Sub
    End If
    
    If mWS.name <> sh.name Then
        Exit Sub
    End If
    
    If existCrossLine(sh) Then
    Else
        Exit Sub
    End If
    
    s = sh.Parent.Saved

    b = Application.ScreenUpdating

    blnCall = True

    lngRow = Target(1).row
    lngCol = Target(1).Column

    Dim sngTop As Single
    Dim sngLeft As Single
    Dim sngHeight As Single
    Dim sngWidth As Single

    sngTop = Target(1).Top - C_MAX / 2
    If sngTop < 0 Then
        sngTop = 0
    End If
    sngHeight = Target(1).Top + C_MAX / 2
    If sngHeight > C_MAX Then
        sngHeight = C_MAX
    End If
    
    sngLeft = Target(1).Left - C_MAX / 2
    If sngLeft < 0 Then
        sngLeft = 0
    End If
    sngWidth = Target(1).Left + C_MAX / 2
    If sngWidth > C_MAX Then
        sngWidth = C_MAX
    End If

    Dim lngFillVisible As Long
    Dim lngFillColor As Long
    Dim dblFillTransparency As Double
    Dim lngLineVisible As Long
    Dim lngLineColor As Long
    Dim sngLineWeight As Single
    Dim strOnAction As String
    Dim lngType As Long
    Dim blnGuid As Boolean
    Dim lngFontColor As Long
    Dim blnEdit As Boolean

    Call getCrossLineSetting(lngType, lngFillVisible, lngFillColor, dblFillTransparency, lngLineVisible, lngLineColor, sngLineWeight, strOnAction, blnGuid, lngFontColor, blnEdit)

    '横軸
    If lngType And C_HOLIZON Then
        Set r = Range(sh.Cells(lngRow, 1), sh.Cells(lngRow, sh.Columns.count - 1))
        
        sh.Shapes(C_HOLIZONTAL_BAR_NAME).Top = r.Top
        sh.Shapes(C_HOLIZONTAL_BAR_NAME).Left = sngLeft
        sh.Shapes(C_HOLIZONTAL_BAR_NAME).Width = sngWidth 'r.Width
        sh.Shapes(C_HOLIZONTAL_BAR_NAME).Height = r.Height
        
        If lngFillVisible = msoTrue Then
            sh.Shapes(C_HOLIZONTAL_BAR_NAME).OnAction = strOnAction
        End If
        
        With sh.Shapes(C_HOLIZONTAL_BAR_NAME).Fill
            .visible = lngFillVisible
            .ForeColor.RGB = lngFillColor
            .Transparency = dblFillTransparency / 100
            .Solid
        End With
        
        With sh.Shapes(C_HOLIZONTAL_BAR_NAME).Line
            .visible = lngLineVisible
            .ForeColor.RGB = lngLineColor
            .Weight = sngLineWeight
            .Transparency = 0
        End With
        
        Set r = Nothing
    End If
    
    '縦軸
    If lngType And C_VERTICAL Then
        Set r = Range(sh.Cells(1, lngCol), sh.Cells(sh.Rows.count - 1, lngCol))
    
        sh.Shapes(C_VERTICAL_BAR_NAME).Top = sngTop
        sh.Shapes(C_VERTICAL_BAR_NAME).Left = r.Left
        sh.Shapes(C_VERTICAL_BAR_NAME).Width = r.Width
        sh.Shapes(C_VERTICAL_BAR_NAME).Height = sngHeight 'r.Height
    
        If lngFillVisible = msoTrue Then
            sh.Shapes(C_VERTICAL_BAR_NAME).OnAction = strOnAction
        End If
        
        With sh.Shapes(C_VERTICAL_BAR_NAME).Fill
            .visible = lngFillVisible
            .ForeColor.RGB = lngFillColor
            .Transparency = dblFillTransparency / 100
            .Solid
        End With
        
        With sh.Shapes(C_VERTICAL_BAR_NAME).Line
            .visible = lngLineVisible
            .ForeColor.RGB = lngLineColor
            .Weight = sngLineWeight
            .Transparency = 0
        End With
        
        Set r = Nothing
    End If

    'ＸＹラベル
    If blnGuid Then
        Set r = Range(sh.Cells(1, lngCol), sh.Cells(sh.Rows.count - 1, lngCol))
    
        sh.Shapes(C_XY_BAR_NAME).Width = 80
        sh.Shapes(C_XY_BAR_NAME).Height = 50
    
        sh.Shapes(C_XY_BAR_NAME).Top = Target(1).Top + Target(1).Height + 10
        sh.Shapes(C_XY_BAR_NAME).Left = Target(1).Left + Target(1).Width + 10
        
        If lngFillVisible = msoTrue Then
            sh.Shapes(C_XY_BAR_NAME).OnAction = strOnAction
        End If
        
        With sh.Shapes(C_XY_BAR_NAME).Fill
            .visible = lngFillVisible
            .ForeColor.RGB = lngFillColor
            .Transparency = dblFillTransparency / 100
            .Solid
        End With
        
        With sh.Shapes(C_XY_BAR_NAME).Line
            .visible = lngLineVisible
            .ForeColor.RGB = lngLineColor
            .Weight = sngLineWeight
            .Transparency = 0
        End With
        
        Dim strBuf As String
        
        strBuf = "  列:" & Target(1).Column & vbCrLf
        strBuf = strBuf & "  行:" & Target(1).row
        With sh.Shapes(C_XY_BAR_NAME).TextFrame2
            .TextRange.Text = strBuf
            .TextRange.Font.NameComplexScript = "Meiryo UI"
            .TextRange.Font.NameFarEast = "Meiryo UI"
            .TextRange.Font.name = "Meiryo UI"
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = lngFontColor
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .VerticalAnchor = msoAnchorMiddle
        End With
        
        Set r = Nothing
    End If


    blnCall = False

    sh.Parent.Saved = s

End Sub
Private Sub XL_LINE_WorkbookBeforeSave(ByVal WB As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
#If VBA7 Then
    On Error Resume Next
    Call deleteCrossLine
#Else
    On Error Resume Next
    Dim obj As Object
    Call lineOnAction(obj, False)
#End If
End Sub
Private Sub XL_LINE_WorkbookAfterSave(ByVal WB As Workbook, ByVal Success As Boolean)
    On Error Resume Next
    If mLineEnable Then
        Call addCrossLine
    End If
End Sub
Private Sub XL_LINE_SheetDeactivate(ByVal sh As Object)
    On Error Resume Next
    Call deleteCrossLine
End Sub
Private Sub XL_LINE_WorkbookDeactivate(ByVal WB As Workbook)
    On Error Resume Next
    Call deleteCrossLine
End Sub
Private Sub XL_LINE_SheetActivate(ByVal sh As Object)
    On Error Resume Next
    If mLineEnable Then
        Call addCrossLine
    End If

End Sub
Private Sub XL_LINE_WorkbookActivate(ByVal WB As Workbook)
    On Error Resume Next
    If mLineEnable Then
        Call addCrossLine
    End If

End Sub



Private Sub XL_LINE_WorkbookBeforeClose(ByVal WB As Workbook, Cancel As Boolean)
    On Error Resume Next
    Dim obj As Object
    
    Call lineOnAction(obj, False)

End Sub
Function existCrossLine(ByVal sh As Worksheet) As Boolean
    
    Dim s As Shape
    
    On Error Resume Next
    
    err.Clear
    Set s = sh.Shapes(C_HOLIZONTAL_BAR_NAME)
    Set s = Nothing
    
    If err.Number = 0 Then
        existCrossLine = True
        Exit Function
    End If
    
    err.Clear
    Set s = sh.Shapes(C_VERTICAL_BAR_NAME)
    Set s = Nothing
    
    If err.Number = 0 Then
        existCrossLine = True
        Exit Function
    End If
    
    existCrossLine = False

End Function



