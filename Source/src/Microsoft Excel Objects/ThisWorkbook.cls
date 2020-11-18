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

Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_FORCEMINIMIZE = 11

Private TV As TaskTrayView
Private MultiProcess As Boolean
'--------------------------------------------------------------
'　ブックオープンイベント
'--------------------------------------------------------------
Private Sub Workbook_Open()
    
    Logger.LogBegin "Workbook_Open"
    
    'CreateObjectでRelaxToolsが開かれた場合
    If Application.visible Then
        MultiProcess = False
    Else
        MultiProcess = True
    End If
    
    '情報ログの出力
    Logger.LogInfo vbCrLf & getVersionInfo
    
    'ショートカットキー設定
    Call setShortCutKey
    
    Call createAllSectionObject
    Call createAllItemObject
    
    Call SetTimeLeap
    
    Set mColSection = rlxInitSectionSetting()
    
    'Excel常駐
    Me.Regident
    
    Logger.LogFinish "Workbook_Open"

End Sub
'--------------------------------------------------------------
'　ブッククローズ前イベント
'--------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error Resume Next
    
    Logger.LogBegin "Workbook_BeforeClose"
    
    If XL_LINE Is Nothing Then
    Else
        Call deleteCrossLine
    End If
    If mScreenEnable Then
        Unload frmScreenShot
    End If
    
    Call DeleteTemporaryFile
    
    '常駐モード
    If CBool(GetSetting(C_TITLE, "Option", "RegidentMode", False)) And MultiProcess = False Then
        If Workbooks.Count > 0 Then
            Dim WB As Workbook
            For Each WB In Workbooks
                WB.Close
            Next
'            Workbooks.Add
            ShowWindow Application.hWnd, SW_FORCEMINIMIZE
            DoEvents
            Cancel = True
        Else
            If MsgBox("Excelを終了します。よろしいですか？", vbQuestion + vbOKCancel, C_TITLE) <> vbOK Then
'                ShowWindow Application.hwnd, SW_FORCEMINIMIZE
'                DoEvents
                Cancel = True
                Exit Sub
            End If
            
            'ショートカットの削除
            Call removeShortCutKey
        
            Me.Unregident
        
        End If
    Else
        'ショートカットの削除
        Call removeShortCutKey
    End If
    
    Logger.LogFinish "Workbook_BeforeClose"

End Sub
Private Sub DeleteTemporaryFile()

    Dim FS As Object
    Dim strTmpBook As String
    
    On Error Resume Next

    Set FS = CreateObject("Scripting.FileSystemObject")
    
    strTmpBook = rlxGetTempFolder() & "*.*"

    
    FS.Deletefile strTmpBook, True


    Set FS = Nothing

End Sub

Sub SetTimeLeap()

    Dim returnValue As Boolean
    
    returnValue = CBool(GetSetting(C_TITLE, "TimeLeap", "Check", False))
    
    If returnValue Then
        ThisWorkbook.StartTimeLeap
    Else
        ThisWorkbook.StopTimeLeap
    End If

End Sub
'Private Sub XL_WindowResize(ByVal WB As Workbook, ByVal wn As Window)
'    Call RefreshRibbon
'End Sub

'Private Sub XL_WorkbookActivate(ByVal WB As Workbook)
'    Call RefreshRibbon
'End Sub

'Private Sub XL_WorkbookBeforeClose(ByVal WB As Workbook, Cancel As Boolean)
'    Call holdBookClose(WB)
'End Sub

'--------------------------------------------------------------
''　ブック保存前イベント
''--------------------------------------------------------------
'Private Sub XL_WorkbookBeforeSave(ByVal WB As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
'
'    Dim strName As String
'    Dim strFolder As String
'    Dim strFile As String
'    Dim strList() As String
'    Dim blnFind As Boolean
'    Dim i As Long
'    Dim strSourceFile As String
'    Dim strDestFile As String
'    Dim exec As Boolean
'    Dim lngGen As Long
'
'    Dim blnScreenUpdating As Boolean
'    Dim blnDisplayAlerts As Boolean
'
'    'イベントのキャンセル
'    If mblnEventCancel Then
'        Exit Sub
'    End If
'
'    If Val(Application.Version) > C_EXCEL_VERSION_2007 Then
'    Else
'        Exit Sub
'    End If
'
'    '-------------------------------------------------
'    ' 簡易世代管理
'    '-------------------------------------------------
'    exec = CBool(GetSetting(C_TITLE, "Backup", "Check", False))
'    If rlxIsExcelFile(WB.FullName) And exec Then
'    Else
'        Exit Sub
'    End If
'
'    blnScreenUpdating = Application.ScreenUpdating
'    blnDisplayAlerts = Application.DisplayAlerts
'
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'
'
'    strFile = GetSetting(C_TITLE, "Backup", "FileList", "")
'    strList = Split(strFile, vbTab)
'
'    blnFind = False
'    For i = 0 To UBound(strList)
''        If InStr(LCase(Application.ActiveWorkbook.FullName), LCase(strList(i))) > 0 Then
'        If InStr(LCase(WB.FullName), LCase(strList(i))) > 0 Then
'             blnFind = True
'            Exit For
'        End If
'    Next
'    If Not blnFind Then
'        GoTo pass
'    End If
'
'
'    strName = WB.FullName
'
'    'パス存在チェック
'    strFolder = rlxGetFullpathFromPathName(strName)
'    If Len(Trim(strFolder)) = 0 Then
'        GoTo pass
'    End If
'
'    'ファイル存在チェック
'    If Not rlxIsFileExists(strName) Then
'        GoTo pass
'    End If
'
'    strFolder = GetSetting(C_TITLE, "Backup", "Folder", "")
'    Dim strFullName As String
'    Dim strPath As String
'
'    lngGen = Val(GetSetting(C_TITLE, "Backup", "Gen", "99"))
'
'    For i = lngGen To 0 Step -1
'
'        If Len(Trim(strFolder)) = 0 Then
'            strFullName = WB.FullName
'        Else
'            strFullName = rlxAddFileSeparator(strFolder) & rlxGetFullpathFromFileName(WB.FullName)
'        End If
'
'        strSourceFile = strFullName & "." & Format$(i, "000")
'        If rlxIsFileExists(strSourceFile) Then
'            If i = lngGen Then
'                Kill strSourceFile
'            Else
'                strDestFile = strFullName & "." & Format$(i + 1, "000")
'                Name strSourceFile As strDestFile
'            End If
'        End If
'
'    Next
'
'pass:
'
'    Application.ScreenUpdating = blnScreenUpdating
'    Application.DisplayAlerts = blnDisplayAlerts
'
'End Sub
'
''--------------------------------------------------------------
''　ブック保存後イベント
''--------------------------------------------------------------
'Private Sub XL_WorkbookAfterSave(ByVal WB As Workbook, ByVal Success As Boolean)
'
'    Dim a1exec As Boolean
'    Dim backexec As Boolean
'    Dim Target As Boolean
'    Dim strFile As String
'    Dim strList() As String
'    Dim blnFind As Boolean
'    Dim i As Long
'
'    Dim blnScreenUpdating As Boolean
'    Dim blnDisplayAlerts As Boolean
'
'    'イベントのキャンセル
'    If mblnEventCancel Or Success = False Then
'        Exit Sub
'    End If
'
'    backexec = GetSetting(C_TITLE, "Backup", "Check", False)
'
'    '-------------------------------------------------
'    ' 簡易世代管理
'    '-------------------------------------------------
'    If rlxIsExcelFile(WB.FullName) And backexec Then
'
'        blnDisplayAlerts = Application.DisplayAlerts
'        blnScreenUpdating = Application.ScreenUpdating
'
'        Application.DisplayAlerts = False
'        Application.ScreenUpdating = False
'
'        strFile = GetSetting(C_TITLE, "Backup", "FileList", "")
'        strList = Split(strFile, vbTab)
'
'        blnFind = False
'        For i = 0 To UBound(strList)
'            If InStr(LCase(Application.ActiveWorkbook.FullName), LCase(strList(i))) > 0 Then
'                 blnFind = True
'                Exit For
'            End If
'        Next
'        If blnFind Then
'            Dim strFolder As String
'            Dim strFullName As String
'            strFolder = GetSetting(C_TITLE, "Backup", "Folder", "")
'            If Len(Trim(strFolder)) = 0 Then
'                strFullName = WB.FullName & ".000"
'            Else
'                strFullName = rlxAddFileSeparator(strFolder) & rlxGetFullpathFromFileName(WB.FullName) & ".000"
'            End If
'
'            Dim DateCreated As Date
'            With CreateObject("Scripting.FileSystemObject")
'                DateCreated = .GetFile(WB.FullName).DateCreated
'            End With
'
'            mblnEventCancel = True
'            WB.SaveCopyAs strFullName
'            mblnEventCancel = False
'
'            Dim a As FileTime
'            Set a = New FileTime
'            a.SetCreationTime strFullName, DateCreated
'
'        End If
'
'        Application.DisplayAlerts = blnDisplayAlerts
'        Application.ScreenUpdating = blnScreenUpdating
'
'    End If
'
'End Sub
Public Sub enableCrossLine()
    
    On Error Resume Next
    
    Call addCrossLine
    
    Set XL_LINE = Excel.Application
    
End Sub
Private Sub addCrossLine()

    Dim blnFillVisible As Boolean
    Dim lngFillColor As Long
    Dim dblFillTransparency As Double
    Dim lngLineVisible As Long
    Dim lngLineColor As Long
    Dim lngFontColor As Long
    Dim sngLineWeight As Single
    Dim strOnAction As String
    Dim lngType As Long
    Dim blnGuid As Boolean
    Dim blnEdit As Boolean
    Dim blnLineWidth As Boolean
    Dim lngBackColor As Long
    Dim dblGuidTransparency As Double

    Call getCrossLineSetting(lngType, blnFillVisible, lngFillColor, dblFillTransparency, lngLineVisible, lngLineColor, sngLineWeight, strOnAction, blnGuid, lngFontColor, blnEdit, blnLineWidth, lngBackColor, dblGuidTransparency)
    
    
    
    
'    Select Case lngType
'        Case C_HOLIZON
'            optHolizon.Value = True
'        Case C_VERTICAL
'            optVertical.Value = True
'        Case Else
'            optAll.Value = True
'    End Select
'
'    If blnFillVisible Then
'        chkFillVisible.Value = False
'    Else
'        chkFillVisible.Value = True
'    End If
'
'    lblFillColor.BackColor = lngFillColor
'    txtFillTransparency.Value = dblFillTransparency
'
'    chkGuid.Value = blnGuid
'
'    lblEven.BackColor = lngLineColor
'
'    txtCol.Value = sngLineWeight
'
'    lblFont.BackColor = lngFontColor
'    lblBack.BackColor = lngBackColor

'    frmCrossLineGauge.BackColor = lngBackColor
'    frmCrossLineGauge.Transparency = dblGuidTransparency

    If blnGuid Then
        frmCrossLineGauge.Run
    End If
    
    If lngType And 2 Then
        frmCrossLineH.BackColor = lngFillColor
        frmCrossLineH.Transparency = dblFillTransparency
        frmCrossLineH.Run
    End If
    
    If lngType And 1 Then
        frmCrossLineV.BackColor = lngFillColor
        frmCrossLineV.Transparency = dblFillTransparency
        frmCrossLineV.Run
    End If
    
    mstrCrossBook = ActiveWorkbook.Name
    
    Call ShowCursor

End Sub

Public Sub disableCrossLine()

    Set XL_LINE = Nothing
    
'    If mWS Is Nothing Then
'    Else
        Call deleteCrossLine
'    End If
'    Set mWS = Nothing
    
End Sub

Private Sub deleteCrossLine()

    On Error Resume Next
    
    Unload frmCrossLineH
    Unload frmCrossLineV
    Unload frmCrossLineGauge

End Sub

Private Sub XL_LINE_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)

    On Error Resume Next
    
    Static blnCall As Boolean

    If blnCall Then
        Exit Sub
    End If
    
    If Val(Application.Version) > C_EXCEL_VERSION_2010 Then
        If mstrCrossBook <> sh.Parent.Name Then
            Exit Sub
        End If
    End If
    
    blnCall = True

    Call ShowCursor

    blnCall = False
    
End Sub
Private Sub ShowCursor()

    Dim spx As Long
    Dim spy As Long

    Dim RE As rect
    Dim re2 As rect
    Dim r As Range
    Dim s As Range

    GetWindowRect Application.hWnd, RE

    With ActiveWindow.ActivePane

        Set r = ActiveWindow.Panes(1).VisibleRange(1)
        
        spx = ActiveWindow.Panes(1).PointsToScreenPixelsX(r.Left)
        spy = ActiveWindow.Panes(1).PointsToScreenPixelsY(r.Top)
    
'        If TypeName(Selection) = "Range" Then
'            Set s = Selection
'        Else
            Set s = ActiveCell.MergeArea
'        End If
        
        Dim X As Long
        Dim Y As Long
        
        X = .PointsToScreenPixelsX(s.Left + s.width) - .PointsToScreenPixelsX(s.Left)
'        If x > re.Bottom - spy Then
'            x = re.Bottom - spy
'        End If
        
        Y = .PointsToScreenPixelsY(ActiveCell.Top + s.Height) - .PointsToScreenPixelsY(s.Top)
'        If y > re.Right - spx Then
'            y = re.Right - spx
'        End If
        
        MoveWindow frmCrossLineH.hWnd, .PointsToScreenPixelsX(s.Left), spy, X, RE.Bottom - spy, True
        MoveWindow frmCrossLineV.hWnd, spx, .PointsToScreenPixelsY(s.Top), RE.Right - spx, Y, True
        
        GetWindowRect frmCrossLineGauge.hWnd, re2
        
        If .PointsToScreenPixelsX(s.Left + s.width + 10) + (re2.Right - re2.Left) > RE.Right Then
            MoveWindow frmCrossLineGauge.hWnd, .PointsToScreenPixelsX(s.Left - 10) - (re2.Right - re2.Left), .PointsToScreenPixelsY(s.Top - 10) - (re2.Bottom - re2.Top), re2.Right - re2.Left, re2.Bottom - re2.Top, True
        Else
            MoveWindow frmCrossLineGauge.hWnd, .PointsToScreenPixelsX(s.Left + s.width + 10), .PointsToScreenPixelsY(s.Top - 10) - (re2.Bottom - re2.Top), re2.Right - re2.Left, re2.Bottom - re2.Top, True
        End If

'        frmCrossLineGauge.txtAddress = s(1).Address(False, False)
        If s.Rows.Count = 1 Then
            frmCrossLineGauge.txtRow = s(1).Row
        Else
            frmCrossLineGauge.txtRow = s(1).Row & ":" & s(s.Count).Row
        End If
        
        If s.Columns.Count = 1 Then
            frmCrossLineGauge.txtColumn = s(1).Column
        Else
            frmCrossLineGauge.txtColumn = s(1).Column & ":" & s(s.Count).Column
        End If
        
    End With
    
End Sub
'Private Sub XL_LINE_WorkbookBeforeSave(ByVal WB As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
'#If VBA7 Then
'    On Error Resume Next
'    Call deleteCrossLine
'#Else
'    On Error Resume Next
'    Dim obj As Object
'    Call lineOnAction(obj, False)
'#End If
'End Sub
'Private Sub XL_LINE_WorkbookAfterSave(ByVal WB As Workbook, ByVal Success As Boolean)
'    On Error Resume Next
'    If mLineEnable Then
'        Call addCrossLine
'    End If
'End Sub
'Private Sub XL_LINE_SheetDeactivate(ByVal sh As Object)
'    On Error Resume Next
'    Call deleteCrossLine
'End Sub
'Private Sub XL_LINE_WorkbookDeactivate(ByVal WB As Workbook)
'    On Error Resume Next
'    Call deleteCrossLine
'End Sub
'Private Sub XL_LINE_SheetActivate(ByVal sh As Object)
'    On Error Resume Next
'    If mLineEnable Then
'        Call addCrossLine
'    End If
'
'End Sub
'Private Sub XL_LINE_WorkbookActivate(ByVal WB As Workbook)
'    On Error Resume Next
'    If mLineEnable Then
'        Call addCrossLine
'    End If
'
'End Sub
'Private Sub XL_LINE_WorkbookBeforeClose(ByVal WB As Workbook, Cancel As Boolean)
'    On Error Resume Next
'    Dim obj As Object
'
'    Call lineOnAction(obj, False)
'
'End Sub
Function existCrossLine(ByVal sh As Worksheet) As Boolean
    
    Dim s As Shape
    
    On Error Resume Next
    
    Err.Clear
    Set s = sh.Shapes(C_HOLIZONTAL_BAR_NAME)
    Set s = Nothing
    
    If Err.Number = 0 Then
        existCrossLine = True
        Exit Function
    End If
    
    Err.Clear
    Set s = sh.Shapes(C_VERTICAL_BAR_NAME)
    Set s = Nothing
    
    If Err.Number = 0 Then
        existCrossLine = True
        Exit Function
    End If
    
    existCrossLine = False

End Function

Public Sub StartTimeLeap()
    Set XL = Excel.Application
End Sub
Public Sub StopTimeLeap()
    Set XL = Nothing
End Sub



'Private Sub XL_LINE_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
'
'    On Error Resume Next
'
'    Static blnCall As Boolean
'
'    If blnCall Then
'        Exit Sub
'    End If
'
'    If Val(Application.Version) > C_EXCEL_VERSION_2010 Then
'        If mstrCrossBook <> Wb.Name Then
'            Exit Sub
'        End If
'    End If
'
'    blnCall = True
'
'    Call ShowCursor
'
'    blnCall = False
'
'End Sub

'--------------------------------------------------------------
'　ブック保存前イベント
'--------------------------------------------------------------
Private Sub XL_WorkbookBeforeSave(ByVal WB As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Dim blnDisplayAlerts As Boolean
    Dim strName As String
    
    If WB.IsAddin Then
        Exit Sub
    End If
    
    blnDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    If Val(Application.Version) > C_EXCEL_VERSION_2007 Then
    Else
        GoTo pass
    End If
    
    '-------------------------------------------------
    ' 簡易世代管理
    '-------------------------------------------------
    Dim exec As Boolean
    exec = CBool(GetSetting(C_TITLE, "TimeLeap", "Check", False))
    
    strName = WB.FullName
    
    If rlxIsExcelFile(strName) And exec Then
    Else
        GoTo pass
    End If
    
    'パス存在チェック
    If Len(Trim(rlxGetFullpathFromPathName(strName))) = 0 Then
        GoTo pass
    End If
    
    On Error GoTo e
    
    Call CreateHistory(strName)

pass:
    Application.DisplayAlerts = blnDisplayAlerts
    Exit Sub

e:
    Application.DisplayAlerts = blnDisplayAlerts
    MsgBox Err.Description, vbOKOnly + vbCritical, C_TITLE
End Sub

'--------------------------------------------------------------
'　ブック保存後イベント
'--------------------------------------------------------------
Private Sub XL_WorkbookAfterSave(ByVal WB As Workbook, ByVal Success As Boolean)

'    Dim DateCreated As Date
'    Dim FT As FileTime
    Dim blnDisplayAlerts As Boolean
    
    If WB.IsAddin Then
        Exit Sub
    End If
    
    blnDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    'イベントのキャンセル
    If Not Success Then
        GoTo pass
    End If
    
    On Error GoTo e
    
    Dim DateCreated As Date
    Dim DateLastModified As Date

    Dim a As FILETIME
    Set a = New FILETIME
    
        
    '-------------------------------------------------
    ' 簡易世代管理
    '-------------------------------------------------
    Dim backexec As Boolean
    backexec = GetSetting(C_TITLE, "TimeLeap", "Check", False)
    If rlxIsExcelFile(WB.FullName) And backexec Then
    
        With CreateObject("Scripting.FileSystemObject")
            
            Dim strFolder As String
            Dim strFullName As String
            
            strFolder = GetSetting(C_TITLE, "TimeLeap", "Folder", GetTimeLeapFolder())
            If Not .FolderExists(strFolder) Then
                .createFolder strFolder
            End If

            '000作成
            DateCreated = .GetFile(WB.FullName).DateCreated
            DateLastModified = .GetFile(WB.FullName).DateLastModified
            
            strFullName = .BuildPath(strFolder, WB.Name) & ".000"
            .CopyFile WB.FullName, strFullName, True
        
            a.SetCreationTime strFullName, DateCreated
            a.SetLastWriteTime strFullName, DateLastModified
        
        End With
    End If

pass:
    Application.DisplayAlerts = blnDisplayAlerts
    Exit Sub

e:
    Application.DisplayAlerts = blnDisplayAlerts
    MsgBox Err.Description, vbOKOnly + vbCritical, C_TITLE
End Sub
'--------------------------------------------------------------
' Excel常駐コマンド
'--------------------------------------------------------------
Public Sub Regident()
    '常駐モード
    If CBool(GetSetting(C_TITLE, "Option", "RegidentMode", False)) And MultiProcess = False And TV Is Nothing Then
        Set TV = New TaskTrayView
        TV.AddIcon Application.hWnd, "Excel常駐モード"
        TV.ShowBalloon "Excel常駐モード" & vbCrLf & "完全に終了させたい場合は、最小化されている空のウィンドウを閉じてください。"
    End If
End Sub
'--------------------------------------------------------------
' Excel常駐コマンドの解除
'--------------------------------------------------------------
Public Sub Unregident()
    '常駐モードの解除
    If TV Is Nothing Then
    Else
        TV.DeleteIcon
        Set TV = Nothing
    End If
End Sub

