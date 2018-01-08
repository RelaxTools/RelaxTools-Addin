Attribute VB_Name = "basRibbon"
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

Private Const C_START_ROW As Long = 25 '13
Private Const C_COL_NO As Long = 1
Private Const C_COL_CATEGORY As Long = 2
Private Const C_COL_MACRO As Long = 3
Private Const C_COL_LABEL As Long = 4
Private Const C_COL_DIVISION As Long = 5
Private Const C_COL_HELP As Long = 6
Private Const C_COL_DESCRIPTION As Long = 7

Private Const C_COLOR_OTHER As String = "99"

Private mIR As IRibbonUI

Private mSecTog01 As Boolean
Private mSecTog02 As Boolean
Private mSecTog03 As Boolean
Private mSecTog04 As Boolean
Private mSecTog05 As Boolean
Private mSecTog06 As Boolean

'Ａ１保存のパブリック変数
Public pblnA1SaveCheck As Boolean


Public mLineEnable As Boolean
Public mScrollEnable As Boolean
Public mScreenEnable As Boolean

'メニュー
Public mObjMenu As Object

Public mblnSushi As Boolean


Private Const IID_IPictureDisp As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"
Private Const PICTYPE_BITMAP As Long = 1
    
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As LongPtr, bitmap As LongPtr) As LongPtr
    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As Long) As LongPtr
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As LongPtr
    Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As LongPtr
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As LongPtr
    Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, lpiid As Any) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, RefIID As Long, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As LongPtr
    
    Private Type PICTDESC
        size As Long
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
    
    Private Declare PtrSafe Function GdipCreateSolidFill Lib "gdiplus" (ByVal pColor As Long, ByRef brush As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, graphics As LongPtr) As Long
    Private Declare PtrSafe Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
    Private Declare PtrSafe Function GdipSetSmoothingMode Lib "gdiplus" (ByVal mGraphics As LongPtr, ByVal mSmoothingMode As Long) As Long
    Private Declare PtrSafe Function GdipDeleteBrush Lib "gdiplus" (ByVal mBrush As LongPtr) As Long
    Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    
    Private Declare PtrSafe Function GdipCreatePen1 Lib "gdiplus" (ByVal pColor As Long, ByVal width As Long, ByVal unit As Long, ByRef hPen As LongPtr) As Long
    Private Declare PtrSafe Function GdipDrawRectangle Lib "gdiplus" (ByVal hGraphics As LongPtr, ByVal hPen As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
    Private Declare PtrSafe Function GdipDeletePen Lib "gdiplus" (ByVal hPen As LongPtr) As Long
    
#Else
    Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, bitmap As Long) As Long
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
    Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
    Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, lpiid As Any) As Long
    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As Long, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    
    Private Type PICTDESC
      size As Long
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

    Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal pColor As Long, ByRef brush As Long) As Long
    Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, graphics As Long) As Long
    Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
    Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
    Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal mBrush As Long) As Long
    Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal pColor As Long, ByVal width As Long, ByVal unit As Long, ByRef hPen As Long) As Long
    Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal X As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
    Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As Long
#End If


Private Const SmoothingModeAntiAlias    As Long = &H4

'--------------------------------------------------------------------
' マクロ名取得
'--------------------------------------------------------------------
Private Function getMacroName(control As IRibbonControl) As String
    
    Dim lngPos As Long
    
    '同じマクロを複数登録可能とするためにドット以降の文字を削除
    lngPos = InStr(control.id, ".")

    If lngPos = 0 Then
        getMacroName = control.id
    Else
        getMacroName = Mid$(control.id, 1, lngPos - 1)
    End If

End Function
'--------------------------------------------------------------------
' シートから指定項目を取得する
'--------------------------------------------------------------------
Private Function getSheetItem(control As IRibbonControl, lngItem As Long) As String

    Dim lngPos As Long
    Dim strBuf As String
    Dim i As Long
    Dim m As MenuDTO
    Dim Key As String
    
    getSheetItem = ""
    
    strBuf = getMacroName(control)
    
    If mObjMenu Is Nothing Then
    
        Set mObjMenu = CreateObject("Scripting.Dictionary")
    
        i = C_START_ROW
        
        Do Until ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_NO).Value = ""
            
            Set m = New MenuDTO
            m.Category = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_CATEGORY).Value
            m.Macro = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_MACRO).Value
            m.Label = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_LABEL).Value
            m.Devision = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_DIVISION).Value
            m.Help = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_HELP).Value
            m.Description = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_DESCRIPTION).Value
            
            If Not mObjMenu.Exists(m.Macro) Then
                mObjMenu.Add m.Macro, m
            Else
                MsgBox "メニューのマクロ名が重複しています。" & strBuf
            End If
            i = i + 1
        Loop
        
    End If
    
    If mObjMenu.Exists(strBuf) Then
        Select Case lngItem
            Case C_COL_CATEGORY
                getSheetItem = mObjMenu.Item(strBuf).Category
            Case C_COL_MACRO
                getSheetItem = mObjMenu.Item(strBuf).Macro
            Case C_COL_LABEL
                getSheetItem = mObjMenu.Item(strBuf).Label
            Case C_COL_DIVISION
                getSheetItem = mObjMenu.Item(strBuf).Devision
            Case C_COL_HELP
                getSheetItem = mObjMenu.Item(strBuf).Help
            Case C_COL_DESCRIPTION
                getSheetItem = mObjMenu.Item(strBuf).Description
        End Select
    Else
        getSheetItem = ""
    End If
End Function
'--------------------------------------------------------------------
' リボン表示設定取得
'--------------------------------------------------------------------
Sub tabGetVisible(control As IRibbonControl, ByRef visible)

    visible = GetSetting(C_TITLE, "Ribbon", Replace(control.id, "Tab", ""), True)

End Sub
'--------------------------------------------------------------------
' スシ表示取得
'--------------------------------------------------------------------
Sub sushiGetVisible(control As IRibbonControl, ByRef visible)

    visible = mblnSushi
1
End Sub
'--------------------------------------------------------------------
' リボン押下状態取得
'--------------------------------------------------------------------
Sub tabGetPressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = GetSetting(C_TITLE, "Ribbon", control.id, True)

End Sub
'--------------------------------------------------------------------
' リボン表示設定
'--------------------------------------------------------------------
Sub tabOnAction(control As IRibbonControl, pressed As Boolean)
    
    SaveSetting C_TITLE, "Ribbon", control.id, pressed
    
    Call RefreshRibbon
    
End Sub
'--------------------------------------------------------------------
'リボンより受け取ったIDをそのままマクロ名として実行するラッパー関数
'--------------------------------------------------------------------
Public Sub RibbonOnAction(control As IRibbonControl)

    Dim lngPos As Long
    Dim strBuf As String
    
    On Error GoTo e
    
    strBuf = getMacroName(control)
    
    '開始ログ
    Logger.LogBegin strBuf
    
    '文字列のマクロ名を実行する。
    Application.Run strBuf
    
    
    Call RefreshRibbon(control)

    '繰り返しが有効の場合
    If CBool(GetSetting(C_TITLE, "Option", "OnRepeat", True)) Then
        Dim strLabel As String
        strLabel = getSheetItem(control, C_COL_LABEL)
        Application.OnRepeat strLabel, strBuf
    End If
    
    '終了ログ
    Logger.LogFinish strBuf
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'チェックボックス設定取得
'--------------------------------------------------------------------
Public Sub CheckGetPressed(control As IRibbonControl, ByRef returnValue)
    
    On Error GoTo e
    
    returnValue = GetSetting(C_TITLE, "Backup", "Check", False)
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'チェックボックス設定
'--------------------------------------------------------------------
Public Sub CheckOnAction(control As IRibbonControl, pressed As Boolean)
    
    On Error GoTo e
    
    SaveSetting C_TITLE, "Backup", "Check", pressed
    
    Call RefreshRibbon(control)
        
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'チェックボックスEnable/Disable
'--------------------------------------------------------------------
Sub CheckSetEnabled(control As IRibbonControl, ByRef enabled)

    On Error GoTo e
    
    If Val(Application.Version) > C_EXCEL_VERSION_2007 Then
        
'        enabled = CBool(GetSetting(C_TITLE, "Backup", "Check", False))
        enabled = True
        
    Else
        enabled = False
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' ヘルプ内容を表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub GetSupertip(control As IRibbonControl, ByRef Screentip)

    On Error GoTo e
    
    Screentip = getSheetItem(control, C_COL_HELP)

    Call RefreshRibbon

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' メニュー表示内容を表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub GetDescription(control As IRibbonControl, ByRef Screentip)

    On Error GoTo e
    
    Screentip = getSheetItem(control, C_COL_DESCRIPTION)

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' ラベルを表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub GetLabel(control As IRibbonControl, ByRef Screentip)

    On Error GoTo e
    
    Screentip = getSheetItem(control, C_COL_LABEL)
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' 2003互換色　背景色自動
'--------------------------------------------------------------------
Sub legacyBackDefault()

    On Error Resume Next
    
    SaveSetting C_TITLE, "Color2003", "back", C_COLOR_OTHER
    execSelectionFormatBackColor

    Call RefreshRibbon

End Sub
'--------------------------------------------------------------------
' 2003互換色　文字色自動
'--------------------------------------------------------------------
Sub legacyFontDefault()

    On Error Resume Next
    
    SaveSetting C_TITLE, "Color2003", "font", C_COLOR_OTHER
    execSelectionFormatFontColor
    Call RefreshRibbon

End Sub
'--------------------------------------------------------------------
' 2003互換色　線色自動
'--------------------------------------------------------------------
Sub legacyLineDefault()

    On Error Resume Next
    
    SaveSetting C_TITLE, "Color2003", "line", C_COLOR_OTHER
    execSelectionFormatLineColor

    Call RefreshRibbon
    
End Sub
'--------------------------------------------------------------------
' 2003互換色選択時イベント
'--------------------------------------------------------------------
Public Sub colorOnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    On Error Resume Next
    Dim strBuf As String

    Select Case control.id
        Case "fontColorGallery"
            SaveSetting C_TITLE, "Color2003", "font", Right(selectedId, 2)
            strBuf = "execSelectionFormatFontColor"
        Case "backColorGallery"
            SaveSetting C_TITLE, "Color2003", "back", Right(selectedId, 2)
            strBuf = "execSelectionFormatBackColor"
        Case "lineColorGallery"
            SaveSetting C_TITLE, "Color2003", "line", Right(selectedId, 2)
            strBuf = "execSelectionFormatLineColor"
    End Select
    
    '開始ログ
    Logger.LogBegin strBuf
    
    '文字列のマクロ名を実行する。
    Application.Run strBuf
    
    Call RefreshRibbon(control)

    '繰り返しが有効の場合
    If CBool(GetSetting(C_TITLE, "Option", "OnRepeat", True)) Then
        Dim strLabel As String
        strLabel = getSheetItem(control, C_COL_LABEL)
        Application.OnRepeat strLabel, strBuf
    End If
    
    '終了ログ
    Logger.LogFinish strBuf
    
    Call RefreshRibbon

End Sub
'--------------------------------------------------------------------
' Dynamicメニュー
'--------------------------------------------------------------------
Private Function ribbonDinamicMenu(control As IRibbonControl, ByRef content)

'ByRef objApp As Object, ByRef WS As Worksheet
'<menu xmlns="http://schemas.microsoft.com/office/2006/01/customui">
'  <button id="dynaButton" label="Button"
'    onAction="OnAction" imageMso="FoxPro"/>
'  <toggleButton id="dynaToggleButton" label="Toggle Button"
'    onAction="OnToggleAction" image="logo.bmp"/>
'  <menuSeparator id="div2"/>
'  <dynamicMenu id="subMenu" label="Sub Menu" getContent="GetSubContent" />
'</menu>


    'On Error Resume Next

    Dim WS As Worksheet

    Dim strNo As String
    Dim strMenu As String
    Dim strSubMenu As String
    Dim strMacro As String
    Dim strBikou As String
    Dim lngRow As Long
    
    Dim blnBeginGroup As Boolean
    Dim blnBeginGroupSub As Boolean
    Dim blnBeginSubMenu As Boolean
    
    Dim blnFirst As Boolean
    
    Dim strXML As String
    Dim lngNo As Long
    
    'コントロールIDからメニュー名を取得
    Set WS = ThisWorkbook.Worksheets(control.id)
    
    
    Const C_START_ROW As Long = 3
    Const C_COL_NO As Long = 1
    Const C_COL_MENU As Long = 2
    Const C_COL_SUB_MENU As Long = 3
    Const C_COL_MACRO As Long = 4
    Const C_COL_BIKOU As Long = 5

    blnBeginGroup = False
    blnBeginSubMenu = False
    
    strXML = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" & vbCrLf
    lngNo = 1
    lngRow = C_START_ROW
    strNo = WS.Cells(lngRow, C_COL_NO)
    Do Until strNo = ""
    
        'メニュー名
        strMenu = WS.Cells(lngRow, C_COL_MENU)
        
        'サブメニュー名
        strSubMenu = WS.Cells(lngRow, C_COL_SUB_MENU)
            
        'マクロ名
        strMacro = WS.Cells(lngRow, C_COL_MACRO)
        
        '備考
        strBikou = WS.Cells(lngRow, C_COL_BIKOU)
        
        Select Case strMenu
            Case ""
                'メニューが空の場合以前作成したメニューの下
            Case "-"
                '次回作成するメニューの前にセパレータを作成
                blnBeginGroup = True
            Case Else
                If blnBeginSubMenu Then
                    strXML = strXML & "  </menu>" & vbCrLf
                    blnBeginSubMenu = False
                End If
                If strSubMenu <> "" Then
                    strXML = strXML & "  <menu id=""menu" & lngNo & """ label=""" & rlxHtmlSanitizing(strMenu) & """ >" & vbCrLf
                    lngNo = lngNo + 1
                    blnBeginSubMenu = True
                Else

                    If blnBeginGroup Then
                        strXML = strXML & "  <menuSeparator id=""div" & lngNo & """/>" & vbCrLf
                        lngNo = lngNo + 1
                    End If
                    
                    If strBikou = "" Then
                        strXML = strXML & "  <button id=""" & strMacro & """ label=""" & rlxHtmlSanitizing(strMenu) & """ onAction=""ribbonOnAction""/>" & vbCrLf
                    Else
                        strXML = strXML & "  <button id=""" & strMacro & """ label=""" & rlxHtmlSanitizing(strMenu) & """ onAction=""ribbonOnAction"" supertip=""" & strBikou & """/>" & vbCrLf
                    End If
                End If
                
                blnBeginGroup = False
        End Select
    
        Select Case strSubMenu
            Case ""
            Case "-"
                blnBeginGroupSub = True
            Case Else
                
                If blnBeginGroupSub Then
                    strXML = strXML & "    <menuSeparator id=""div" & lngNo & """/>" & vbCrLf
                    lngNo = lngNo + 1
                End If
            
                If strBikou = "" Then
                    strXML = strXML & "    <button id=""" & strMacro & """ label=""" & rlxHtmlSanitizing(strSubMenu) & """ onAction=""ribbonOnAction""/>" & vbCrLf
                Else
                    strXML = strXML & "    <button id=""" & strMacro & """ label=""" & rlxHtmlSanitizing(strSubMenu) & """ onAction=""ribbonOnAction"" supertip=""" & strBikou & """/>" & vbCrLf
                End If

                blnBeginGroupSub = False
        End Select
        
        lngRow = lngRow + 1
        strNo = WS.Cells(lngRow, C_COL_NO)
    Loop
    
    strXML = strXML & "</menu>" & vbCrLf
    
    Set WS = Nothing

    '作成したXMLを戻す
    content = strXML

End Function
'--------------------------------------------------------------------
' リボン状態取得
'--------------------------------------------------------------------
Sub getRibbonEnabled(control As IRibbonControl, ByRef enabled)

    enabled = True
    
End Sub
'--------------------------------------------------------------------
' リボンロード時イベント
'--------------------------------------------------------------------
Sub ribbonLoaded(ByRef IR As IRibbonUI)
    
    On Error GoTo e
    
    Set mIR = IR
    Call ThisWorkbook.setIRibbon(IR)
    
    'リボンハンドルのアドレスをレジストリに保存、実行時エラーの場合に復元する。
    SaveSetting C_TITLE, "Ribbon", "Address", CStr(ObjPtr(IR))
        
    Dim strPos As String
    
    '段落番号の規定のボタンを押下済みにする
    strPos = GetSetting(C_TITLE, "Section", "pos", "1")
    Select Case strPos
        Case "1"
            mSecTog01 = True
        Case "2"
            mSecTog02 = True
        Case "3"
            mSecTog03 = True
        Case "4"
            mSecTog04 = True
        Case "5"
            mSecTog05 = True
        Case "6"
            mSecTog06 = True
    End Select
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' リボンのリフレッシュ
'--------------------------------------------------------------------
Public Sub RefreshRibbon(Optional control As IRibbonControl)

    Dim strBuf As String
    
    On Error GoTo e
    
    'グローバル変数がクリアされたしまった場合、レジストリから復帰
    If mIR Is Nothing Then
        
        strBuf = GetSetting(C_TITLE, "Ribbon", "Address", 0)
        Set mIR = getObjectFromAddres(strBuf)
        
    End If
    
    If mIR Is Nothing Then
    Else
        If control Is Nothing Then
            mIR.Invalidate
        Else
            mIR.InvalidateControl control.id
        End If
    End If

e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  段落番号トグルボタン
'--------------------------------------------------------------------
Sub sectionPressed(control As IRibbonControl, ByRef returnValue)
    
    On Error GoTo e
    
    Select Case control.id
        Case "sectionSetting01"
            returnValue = mSecTog01
    
        Case "sectionSetting02"
            returnValue = mSecTog02
    
        Case "sectionSetting03"
            returnValue = mSecTog03
    
        Case "sectionSetting04"
            returnValue = mSecTog04
    
        Case "sectionSetting05"
            returnValue = mSecTog05
    
        Case "sectionSetting06"
            returnValue = mSecTog06
    End Select
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  現在の段落番号の設定
'--------------------------------------------------------------------
Sub sectionOnAction(control As IRibbonControl, pressed As Boolean)
  
    On Error GoTo e
    
    mSecTog01 = False
    mSecTog02 = False
    mSecTog03 = False
    mSecTog04 = False
    mSecTog05 = False
    mSecTog06 = False
  
    Select Case control.id
        Case "sectionSetting01"
            mSecTog01 = pressed
            Call SaveSetting(C_TITLE, "Section", "pos", "1")
            
        Case "sectionSetting02"
            mSecTog02 = pressed
            Call SaveSetting(C_TITLE, "Section", "pos", "2")
            
        Case "sectionSetting03"
            mSecTog03 = pressed
            Call SaveSetting(C_TITLE, "Section", "pos", "3")
            
        Case "sectionSetting04"
            mSecTog04 = pressed
            Call SaveSetting(C_TITLE, "Section", "pos", "4")
            
        Case "sectionSetting05"
            mSecTog05 = pressed
            Call SaveSetting(C_TITLE, "Section", "pos", "5")
            
        Case "sectionSetting06"
            mSecTog06 = pressed
            Call SaveSetting(C_TITLE, "Section", "pos", "6")
            
    End Select
  
    Call RefreshRibbon
    Set mColSection = rlxInitSectionSetting()
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  十字カーソルの押下状態の取得
'--------------------------------------------------------------------
Sub linePressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = mLineEnable
    
End Sub

'--------------------------------------------------------------------
'  十字カーソルの押下時イベント
'--------------------------------------------------------------------
Sub lineOnAction(control As IRibbonControl, pressed As Boolean)

    If ActiveSheet.ProtectContents Then
        MsgBox "保護されているシートでは十字カーソルは実行できません。", vbOKOnly + vbExclamation, C_TITLE
        pressed = False
        Exit Sub
    End If
  
    On Error GoTo e
    
    mLineEnable = pressed
  
    Call RefreshRibbon

    If pressed Then
        ThisWorkbook.enableCrossLine
    Else
        ThisWorkbook.disableCrossLine
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

''--------------------------------------------------------------------
''  フック固定の押下状態の取得
''--------------------------------------------------------------------
'Sub holdPressed(control As IRibbonControl, ByRef returnValue)
'
'    Dim obj As Object
'
'    Set obj = GetHoldList()
'
'    If ActiveWorkbook Is Nothing Then
'    Else
'        returnValue = obj.Exists(ActiveWorkbook.FullName)
'    End If
'    Set obj = Nothing
'
'End Sub
''--------------------------------------------------------------------
''  フック固定の押下時イベント
''--------------------------------------------------------------------
'Sub holdOnAction(control As IRibbonControl, pressed As Boolean)
'
'    On Error GoTo e
'
'    Dim obj As Object
'    Dim hold As HoldDto
'
'    If ActiveWorkbook Is Nothing Then
'        Exit Sub
'    End If
'
'    Set hold = New HoldDto
'
'    hold.FullName = ActiveWorkbook.FullName
'    hold.ReadOnly = ActiveWorkbook.ReadOnly
'
'    hold.Top = Application.Top
'    hold.Left = Application.Left
'    hold.Height = Application.Height
'    hold.Width = Application.Width
'
'    hold.WindowState = Application.WindowState
'
'    Set obj = GetHoldList()
'
'    If pressed Then
'
'        If Not rlxIsFileExists(hold.FullName) Then
'            MsgBox "ブックが存在しません。保存してから処理を行ってください。", vbOKOnly + vbExclamation, C_TITLE
'            pressed = False
'            Call RefreshRibbon(control)
'            Exit Sub
'        End If
'
'        If Not obj.Exists(hold.FullName) Then
'            obj.Add hold.FullName, hold
'        End If
'    Else
'        If obj.Exists(hold.FullName) Then
'            obj.Remove hold.FullName
'        End If
'    End If
'
'    SaveHoldList obj
'
'    Set obj = Nothing
'
'    Call RefreshRibbon(control)
'
'    Exit Sub
'e:
'    Call rlxErrMsg(err)
'End Sub
''--------------------------------------------------------------------
''  フック終了時イベント
''--------------------------------------------------------------------
'Sub holdBookClose(ByRef WB As Workbook)
'
'    On Error GoTo e
'
'    Dim obj As Object
'    Dim hold As HoldDto
'
'    If WB Is Nothing Then
'        Exit Sub
'    End If
'
'    If Not rlxIsFileExists(WB.FullName) Then
'        Exit Sub
'    End If
'
'    Set obj = GetHoldList()
'
'    If obj.Exists(WB.FullName) Then
'
'        Set hold = New HoldDto
'        hold.FullName = WB.FullName
'        hold.ReadOnly = WB.ReadOnly
'
'        hold.WindowState = WB.Application.WindowState
'        hold.Top = WB.Application.Top
'        hold.Left = WB.Application.Left
'        hold.Height = WB.Application.Height
'        hold.Width = WB.Application.Width
'
'        Logger.LogTrace "---------------------------------------------"
'        Logger.LogTrace "hold.FullName = " & WB.FullName
'        Logger.LogTrace "hold.WindowState = " & hold.WindowState
'        Logger.LogTrace "hold.Top = " & hold.Top
'        Logger.LogTrace "hold.Left = " & hold.Left
'        Logger.LogTrace "hold.Height = " & hold.Height
'        Logger.LogTrace "hold.Width = " & hold.Width
'        Logger.LogTrace "---------------------------------------------"
'
'        obj.Remove hold.FullName
'        obj.Add hold.FullName, hold
'
'        SaveHoldList obj
'    End If
'
'    Set obj = Nothing
'
'    Exit Sub
'e:
'    Call rlxErrMsg(err)
'End Sub
''--------------------------------------------------------------------
''  フック固定のオープン
''--------------------------------------------------------------------
'Sub Auto_Open()
'
'    UnSyncRun "holdOpen"
'
'End Sub
'Sub holdOpen()
'
'    On Error Resume Next
'
'    Dim obj As Object
'    Dim o As Variant
'    Dim hold As HoldDto
'    Dim WB As Workbook
'
'    Set obj = GetHoldList()
'
'    Application.WindowState = xlNormal
'
'    If obj.count > 0 Then
'        If MsgBox("ピン留めされたブックがあります。復元しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
'            Exit Sub
'        End If
'    End If
'
'    Application.DisplayAlerts = False
'
'    For Each o In obj.keys
'
'        Set hold = obj.Item(o)
'        DoEvents
'        Set WB = Workbooks.Open(FileName:=hold.FullName, ReadOnly:=hold.ReadOnly, IgnoreReadOnlyRecommended:=True, Notify:=False, Local:=True)
'        DoEvents
'
'        Logger.LogTrace "hold.WindowState = " & hold.WindowState
'        Logger.LogTrace "xlMaximized = " & xlMaximized
'        Select Case hold.WindowState
'            Case xlMaximized
'
'                WB.Application.Top = hold.Top + 5
'                WB.Application.Left = hold.Left + 5
'
'                WB.Application.Height = hold.Height - 10
'                WB.Application.Width = hold.Width - 10
'
''                WB.Application.WindowState = xlMaximized
'
'            Case Else
'
'                WB.Application.WindowState = xlNormal
'                WB.Application.Top = hold.Top
'                WB.Application.Left = hold.Left
'                If hold.Height < 150 Or hold.Width < 150 Then
'                Else
'                    WB.Application.Height = hold.Height
'                    WB.Application.Width = hold.Width
'                End If
'
'        End Select
'    Next
'    Application.DisplayAlerts = True
'
'    Set obj = Nothing
'
'End Sub
'
''--------------------------------------------------------------------
''  フック固定の設定内容取得
''--------------------------------------------------------------------
'Function GetHoldList() As Object
'
'    Dim strFile As String
'    Dim varFile As Variant
'    Dim varATTB As Variant
'    Dim i As Long
'    Dim obj As Object
'    Dim hold As HoldDto
'
'    Set obj = CreateObject("Scripting.Dictionary")
'
'    strFile = GetSetting(C_TITLE, "HoldFile", "HoldFile", "")
'
'    Const C_FULLNAME As Long = 0
'    Const C_READONLY As Long = 1
'    Const C_TOP As Long = 2
'    Const C_LEFT As Long = 3
'    Const C_HEIGHT As Long = 4
'    Const C_WIDTH As Long = 5
'    Const C_WINDOWSTATE As Long = 6
'
'    If Len(strFile) <> 0 Then
'
'        varFile = Split(strFile, vbVerticalTab)
'        For i = LBound(varFile) To UBound(varFile)
'
'            varATTB = Split(varFile(i), vbTab)
'
'            Dim j As Long
'
'            Set hold = New HoldDto
'
'            For j = LBound(varATTB) To UBound(varATTB)
'
'                Select Case j
'                    Case C_FULLNAME
'                        hold.FullName = varATTB(j)
'                    Case C_READONLY
'                        hold.ReadOnly = varATTB(j)
'                    Case C_TOP
'                        hold.Top = Val(varATTB(j))
'                    Case C_LEFT
'                        hold.Left = Val(varATTB(j))
'                    Case C_HEIGHT
'                        hold.Height = Val(varATTB(j))
'                    Case C_WIDTH
'                        hold.Width = Val(varATTB(j))
'                    Case C_WINDOWSTATE
'                        hold.WindowState = Val(varATTB(j))
'                End Select
'            Next
'            obj.Add hold.FullName, hold
'        Next
'    End If
'
'    Set GetHoldList = obj
'
'End Function
''--------------------------------------------------------------------
''  フック固定の設定内容保存
''--------------------------------------------------------------------
'Sub SaveHoldList(ByRef obj As Object)
'
'    Dim o As Variant
'    Dim strFile As String
'    Dim hold As HoldDto
'
'    For Each o In obj.keys
'        Set hold = obj.Item(o)
'        If strFile = "" Then
'            strFile = hold.FullName & vbTab & hold.ReadOnly & vbTab & hold.Top & vbTab & hold.Left & vbTab & hold.Height & vbTab & hold.Width & vbTab & hold.WindowState
'        Else
'            strFile = strFile & vbVerticalTab & hold.FullName & vbTab & hold.ReadOnly & vbTab & hold.Top & vbTab & hold.Left & vbTab & hold.Height & vbTab & hold.Width & vbTab & hold.WindowState
'        End If
'    Next
'
'    SaveSetting C_TITLE, "HoldFile", "HoldFile", strFile
'
'End Sub
'--------------------------------------------------------------------
'  ホイール量(小)の押下状態取得
'--------------------------------------------------------------------
Sub scrollPressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = scrollPush
    
End Sub
'--------------------------------------------------------------------
'  ホイール量(小)の押下時イベント
'--------------------------------------------------------------------
Sub scrollOnAction(control As IRibbonControl, pressed As Boolean)

    On Error GoTo e

    mScrollEnable = pressed

    Call RefreshRibbon

    If pressed Then
        scrollLine1
    Else
        scrollLine3
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  スクショモードの押下状態取得
'--------------------------------------------------------------------
Sub screenPressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = mScreenEnable
    
End Sub
'--------------------------------------------------------------------
'  スクショモードの押下時イベント
'--------------------------------------------------------------------
Sub screenOnAction(control As IRibbonControl, pressed As Boolean)
  
    On Error GoTo e
    
    mScreenEnable = pressed
  
    Call RefreshRibbon

    If pressed Then
        frmScreenShot.Show
    Else
        Unload frmScreenShot
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

'--------------------------------------------------------------------
'  リボンサイズ取得(未動作)
'--------------------------------------------------------------------
 Sub GetSize(control As IRibbonControl, ByRef size)
 
    If Application.UsableWidth / 0.75 < 1420 Then
    
        size = RibbonControlSize.RibbonControlSizeRegular
    Else
    
        size = RibbonControlSize.RibbonControlSizeLarge
    End If
 
 End Sub
'--------------------------------------------------------------------
'  リボンサイズ取得(未動作)
'--------------------------------------------------------------------
Public Sub GetSizeLabel(control As IRibbonControl, ByRef lbl)

    If Application.UsableWidth / 0.75 < 1420 Then
        Select Case control.id
            Case "MitomePaste.1"
                lbl = "認め印"
            Case "FilePaste.1"
                lbl = "画像指定"
            Case "bzGallery"
                lbl = "ビジネス印"
        End Select
    Else
        Select Case control.id
            Case "MitomePaste.1"
                lbl = "認め印" & vbCrLf
            Case "FilePaste.1"
                lbl = "画像指定" & vbCrLf
            Case "bzGallery"
                lbl = "ビジネス印" & vbCrLf
        End Select
    End If
 
 End Sub
'--------------------------------------------------------------------
' データ印の数を取得
'--------------------------------------------------------------------
 Sub stampGetItemCount(control As IRibbonControl, ByRef count)

    '設定情報取得
    Dim Col As Collection
    
    Set Col = getProperty()

    count = Col.count

End Sub
'--------------------------------------------------------------------
' データ印のIDを取得
'--------------------------------------------------------------------
Sub stampGetItemId(control As IRibbonControl, Index As Integer, ByRef id)

    id = C_STAMP_FILE_NAME & Format$(Index + 1, "000")

End Sub
'--------------------------------------------------------------------
' データ印のイメージを取得
'--------------------------------------------------------------------
Sub stampGetItemImage(control As IRibbonControl, Index As Integer, ByRef image)

    Set image = getImageStamp(Index + 1)
    
End Sub
'--------------------------------------------------------------------
' データ印押下時イベント
'--------------------------------------------------------------------
Public Sub stampOnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    Call pasteStamp2(selectedIndex + 1)
    Call SaveSetting(C_TITLE, "Stamp", "stampNo", selectedIndex + 1)

End Sub
'--------------------------------------------------------------------
' 認印の数を取得
'--------------------------------------------------------------------
Sub stampMitomeGetItemCount(control As IRibbonControl, ByRef count)

    '設定情報取得
    Dim Col As Collection
    
    Set Col = getPropertyMitome()

    count = Col.count

End Sub
'--------------------------------------------------------------------
' 認印のIDを取得
'--------------------------------------------------------------------
Sub stampMitomeGetItemId(control As IRibbonControl, Index As Integer, ByRef id)

    id = C_STAMP_FILE_NAME & Format$(Index + 1, "000")

End Sub
'--------------------------------------------------------------------
' 認印イメージ取得
'--------------------------------------------------------------------
Sub stampMitomeGetItemImage(control As IRibbonControl, Index As Integer, ByRef image)
    
    Set image = getImageStampMitome(Index + 1)
    
End Sub
'--------------------------------------------------------------------
' 認印押下時イベント
'--------------------------------------------------------------------
Public Sub stampMitomeOnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    Call MitomePaste2(selectedIndex + 1)
    Call SaveSetting(C_TITLE, "StampMitome", "stampNo", selectedIndex + 1)

End Sub
'--------------------------------------------------------------------
'ビジネス印の数を取得
'--------------------------------------------------------------------
Sub stampBzGetItemCount(control As IRibbonControl, ByRef count)

    '設定情報取得
    Dim Col As Collection
    
    Set Col = getPropertyBz()

    count = Col.count

End Sub
'--------------------------------------------------------------------
' ビジネス印のIDを取得
'--------------------------------------------------------------------
Sub stampBzGetItemId(control As IRibbonControl, Index As Integer, ByRef id)

    id = C_STAMP_FILE_NAME & Format$(Index + 1, "000")

End Sub
'--------------------------------------------------------------------
' ビジネス印イメージ取得
'--------------------------------------------------------------------
Sub stampBzGetItemImage(control As IRibbonControl, Index As Integer, ByRef image)

     Set image = getImageStampBz(Index + 1)
    
End Sub
'--------------------------------------------------------------------
' ビジネス印押下時イベント
'--------------------------------------------------------------------
Public Sub stampBzOnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    Call pasteStampBz2(selectedIndex + 1)
    Call SaveSetting(C_TITLE, "StampBz", "stampNo", selectedIndex + 1)

End Sub

Sub GetItemSuperTip(control As IRibbonControl, Index As Integer, ByRef screen)

End Sub
'--------------------------------------------------------------------
'  さくら印の数を取得
'--------------------------------------------------------------------
Sub sakuraGetItemCount(control As IRibbonControl, ByRef count)

    count = 3

End Sub
'--------------------------------------------------------------------
'  さくら印のIDを取得
'--------------------------------------------------------------------
Sub sakuraGetItemId(control As IRibbonControl, Index As Integer, ByRef id)

    id = C_STAMP_FILE_NAME & Format$(Index + 1, "000")

End Sub
'--------------------------------------------------------------------
'  さくら印イメージ取得
'--------------------------------------------------------------------
Sub sakuraGetItemImage(control As IRibbonControl, Index As Integer, ByRef image)

    Set image = getImageSakura(control.id, Index + 1)
    
End Sub
'--------------------------------------------------------------------
'  さくら印押下時イベント
'--------------------------------------------------------------------
Public Sub sakuraOnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    Call pasteSakura(control.id, selectedIndex + 1)

End Sub

'--------------------------------------------------------------------
'  付箋の数を取得
'--------------------------------------------------------------------
Sub fusenGetItemCount(control As IRibbonControl, ByRef count)

    count = 5

End Sub
'--------------------------------------------------------------------
'  付箋のIDを取得
'--------------------------------------------------------------------
Sub fusenGetItemId(control As IRibbonControl, Index As Integer, ByRef id)

    id = C_STAMP_FILE_NAME & Format$(Index + 1, "000")

End Sub
'--------------------------------------------------------------------
'  付箋イメージ取得
'--------------------------------------------------------------------
Sub fusenGetItemImage(control As IRibbonControl, Index As Integer, ByRef image)

    Set image = getImageFusen(control.id, Index + 1)
    
End Sub
'--------------------------------------------------------------------
'  付箋押下時イベント
'--------------------------------------------------------------------
Public Sub fusenOnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    Call pasteFusen(control.id, selectedIndex + 1)

End Sub
'--------------------------------------------------------------------
'  スクショモード設定のEnabled/Disabled
'--------------------------------------------------------------------
Sub getScreenShotEnabled(control As IRibbonControl, ByRef enabled)

    On Error GoTo e
    
    enabled = Not (mScreenEnable)

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  十字カーソル設定のEnabled/Disabled
'--------------------------------------------------------------------
Sub getCrossEnabled(control As IRibbonControl, ByRef enabled)

    On Error GoTo e
    
    enabled = Not (mLineEnable)

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  ホイール量設定のEnabled/Disabled
'--------------------------------------------------------------------
Sub getScrollEnabled(control As IRibbonControl, ByRef enabled)

    On Error GoTo e
    
    enabled = Not (mScrollEnable)

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

'--------------------------------------------------------------------
' かんたん表の数を取得
'--------------------------------------------------------------------
 Sub kantanGetItemCount(control As IRibbonControl, ByRef count)

    '設定情報取得
    Dim Col As Collection
    
    Set Col = getPropertyKantan()

    count = Col.count

End Sub
'--------------------------------------------------------------------
' かんたん表のIDを取得
'--------------------------------------------------------------------
Sub kantanGetItemId(control As IRibbonControl, Index As Integer, ByRef id)

    id = "kantan" & Format$(Index + 1, "000")

End Sub
'--------------------------------------------------------------------
' かんたん表のイメージを取得
'--------------------------------------------------------------------
Sub kantanGetItemImage(control As IRibbonControl, Index As Integer, ByRef image)

    Set image = getImageKantan(Index + 1)
    
End Sub
'--------------------------------------------------------------------
' かんたん表押下時イベント
'--------------------------------------------------------------------
Public Sub KantanOnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    Call kantanPaste2(selectedIndex + 1)
    Call SaveSetting(C_TITLE, "KantanDx", "kantanNo", selectedIndex + 1)

End Sub

'--------------------------------------------------------------------
'  エラーチェックの押下状態の取得
'--------------------------------------------------------------------
Sub errCheckPressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = Application.ErrorCheckingOptions.BackgroundChecking
    
End Sub
'--------------------------------------------------------------------
'  エラーチェックの押下時イベント
'--------------------------------------------------------------------
Sub errCheckOnAction(control As IRibbonControl, pressed As Boolean)
  
    On Error GoTo e
  
    Call RefreshRibbon

    Application.ErrorCheckingOptions.BackgroundChecking = pressed

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  エラーチェックのToggle
'--------------------------------------------------------------------
Sub errCheckToggle()
  
    On Error GoTo e
  
    Application.ErrorCheckingOptions.BackgroundChecking = Not (Application.ErrorCheckingOptions.BackgroundChecking)
    Call RefreshRibbon

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  A1⇔R1C1の押下状態の取得
'--------------------------------------------------------------------
Sub a1Pressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = (Application.ReferenceStyle = xlR1C1)
    
End Sub
'--------------------------------------------------------------------
'  A1⇔R1C1の押下時イベント
'--------------------------------------------------------------------
Sub a1OnAction(control As IRibbonControl, pressed As Boolean)
  
    On Error GoTo e
  
    Call RefreshRibbon

    If pressed Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  A1⇔R1C1トグル動作
'--------------------------------------------------------------------
Sub a1Toggle()
  
    On Error GoTo e
  
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
    Call RefreshRibbon

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  Return時のカーソルの取得
'--------------------------------------------------------------------
Sub MoveAfterReturnPressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = (Application.MoveAfterReturnDirection = xlToRight)
    
End Sub
'--------------------------------------------------------------------
'  Return時のカーソルの押下時イベント
'--------------------------------------------------------------------
Sub MoveAfterReturnOnAction(control As IRibbonControl, pressed As Boolean)
  
    On Error GoTo e
  
    Call RefreshRibbon

    If pressed Then
        Application.MoveAfterReturnDirection = xlToRight
    Else
        Application.MoveAfterReturnDirection = xlDown
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  Return時のカーソルの押下時イベント
'--------------------------------------------------------------------
Sub MoveAfterReturnToggle()
  
    On Error GoTo e
  
    If Application.MoveAfterReturnDirection = xlDown Then
        Application.MoveAfterReturnDirection = xlToRight
    Else
        Application.MoveAfterReturnDirection = xlDown
    End If
    
    Call RefreshRibbon

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  オートコンプリートの取得
'--------------------------------------------------------------------
Sub AutoCompletePressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = Application.EnableAutoComplete
    
End Sub
'--------------------------------------------------------------------
'  オートコンプリート押下時イベント
'--------------------------------------------------------------------
Sub AutoCompleteOnAction(control As IRibbonControl, pressed As Boolean)
  
    On Error GoTo e
  
    Call RefreshRibbon

    Application.EnableAutoComplete = pressed

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  オートコンプリート切り替え
'--------------------------------------------------------------------
Sub AutoCompleteTogggle()
  
    On Error GoTo e
  
    Application.EnableAutoComplete = Not (Application.EnableAutoComplete)
    Call RefreshRibbon

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
'  Pickの押下状態の取得
'--------------------------------------------------------------------
Sub pickPressed(control As IRibbonControl, ByRef returnValue)
    
    returnValue = CBool(GetSetting(C_TITLE, "Shape", "PickMode", False))
    
End Sub
'--------------------------------------------------------------------
'  Pickの押下時イベント
'--------------------------------------------------------------------
Sub pickOnAction(control As IRibbonControl, pressed As Boolean)
  
    On Error GoTo e
  
    Call RefreshRibbon

    Call SaveSetting(C_TITLE, "Shape", "PickMode", pressed)

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

Sub getColorImage(control As IRibbonControl, ByRef image) ' 画像の設定

    Dim pictureId As String
    Dim strColor As String
    Dim lngColor As Long
    Dim blnTranspairent As Boolean
    
    blnTranspairent = False
    
    Select Case control.id
        Case "execSelectionFormatFontColor"
            pictureId = "fontColor"
            strColor = GetSetting(C_TITLE, "Color2003", "font", C_COLOR_OTHER)
        Case "execSelectionFormatLineColor"
            pictureId = "lineColor"
            strColor = GetSetting(C_TITLE, "Color2003", "line", C_COLOR_OTHER)
        Case "execSelectionFormatBackColor"
            pictureId = "backColor"
            strColor = GetSetting(C_TITLE, "Color2003", "back", C_COLOR_OTHER)
    End Select
        
    If strColor = C_COLOR_OTHER Then
        If control.id = "execSelectionFormatBackColor" Then
            strColor = "02"
            blnTranspairent = True
        Else
            strColor = "01"
        End If
    End If
    lngColor = ThisWorkbook.Colors(Val(strColor))
    
    Dim file As String
    
    file = rlxGetAppDataFolder & "images\" & pictureId & ".png"
    
    'イメージが見つからなかったら「×」表示する
    If rlxIsFileExists(file) Then
        Set image = LoadImageColor(file, lngColor, blnTranspairent)
    Else
        image = "CancelRequest"
    End If
    
'    Call RefreshRibbon
'    DoEvents
    
End Sub
Sub getFusenImage(control As IRibbonControl, ByRef image) ' 画像の設定

    Dim pictureId As String
    
    Select Case control.id
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
    
'    Call RefreshRibbon
'    DoEvents
    
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
              .size = LenB(uPicInfo)
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
' イメージを読み込む（色バーつき）
Private Function LoadImageColor(ByVal strFName As String, ByVal lngColor As Long, ByVal blnTranspairent As Boolean) As IPicture

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
        
            If blnTranspairent Then
                Call DrawRectangle(hGdiImage, RGB(150, 150, 150), 100, 0, 12, 15, 3)
                
            Else
                Call FillRectangle(hGdiImage, lngColor, 100, 0, 12, 15, 3)
            End If
            
            Call GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0&)
          
            Dim IID(0 To 3) As Long
            Dim IPic As IPicture
            Dim uPicInfo As PICTDESC
            
            With uPicInfo
              .size = LenB(uPicInfo)
              .Type = PICTYPE_BITMAP
              .hPic = hBitmap
              .hPal = 0&
            End With
                
            Call IIDFromString(StrPtr(IID_IPictureDisp), IID(0))
            Call OleCreatePictureIndirect(uPicInfo, IID(0), True, LoadImageColor)
          
            Call GdipDisposeImage(hGdiImage)
          
        End If
        
        Call GdiplusShutdown(hGdiPlus)
    
    End If
  
End Function

#If VBA7 And Win64 Then
Private Function FillRectangle(ByVal hBitmap As LongPtr, ByVal lColor As Long, ByVal Alpha As Long, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal Height As Long) As Boolean

    Dim hGraphics As LongPtr, hBrush As LongPtr

    If GdipGetImageGraphicsContext(hBitmap, hGraphics) = 0 Then
   
        If GdipCreateSolidFill(ConvertColor(lColor, Alpha), hBrush) = 0 Then
    
            Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
           
            FillRectangle = GdipFillRectangle(hGraphics, hBrush, X, Y, width, Height) = 0
        
            Call GdipDeleteBrush(hBrush)
        End If
        
        Call GdipDeleteGraphics(hGraphics)
    End If
    
End Function
Private Function DrawRectangle(ByVal hBitmap As LongPtr, ByVal lColor As Long, ByVal Alpha As Long, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal Height As Long) As Boolean

    Dim hGraphics As LongPtr, hPen As LongPtr

    If GdipGetImageGraphicsContext(hBitmap, hGraphics) = 0 Then
   
        If GdipCreatePen1(ConvertColor(lColor, Alpha), 1, 2&, hPen) = 0 Then
    
            Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
           
            DrawRectangle = GdipDrawRectangle(hGraphics, hPen, X, Y, width, Height) = 0
        
            Call GdipDeletePen(hPen)
        End If
        
        Call GdipDeleteGraphics(hGraphics)
    End If
    
End Function
#Else
Private Function FillRectangle(ByVal hBitmap As Long, ByVal lColor As Long, ByVal Alpha As Long, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal Height As Long) As Boolean

    Dim hGraphics As Long, hBrush As Long

    If GdipGetImageGraphicsContext(hBitmap, hGraphics) = 0 Then
   
        If GdipCreateSolidFill(ConvertColor(lColor, Alpha), hBrush) = 0 Then
    
            Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
           
            FillRectangle = GdipFillRectangle(hGraphics, hBrush, X, Y, width, Height) = 0
        
            Call GdipDeleteBrush(hBrush)
        End If
        
        Call GdipDeleteGraphics(hGraphics)
    End If
    
End Function
Private Function DrawRectangle(ByVal hBitmap As Long, ByVal lColor As Long, ByVal Alpha As Long, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal Height As Long) As Boolean

    Dim hGraphics As Long, hPen As Long

    If GdipGetImageGraphicsContext(hBitmap, hGraphics) = 0 Then
   
        If GdipCreatePen1(ConvertColor(lColor, Alpha), 1, 2&, hPen) = 0 Then
    
            Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
           
            DrawRectangle = GdipDrawRectangle(hGraphics, hPen, X, Y, width, Height) = 0
        
            Call GdipDeletePen(hPen)
        End If
        
        Call GdipDeleteGraphics(hGraphics)
    End If
    
End Function
#End If

Private Function ConvertColor(Color As Long, Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
 
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function
Public Sub ContextMenus_GetVisible(control As IRibbonControl, ByRef visible)

    Dim strBuf As String
    
    '改ページプレビューも一緒に判定
    strBuf = GetSetting(C_TITLE, "ContextMenu", Replace(control.id, "Layout", ""), "")
    
    If Len(strBuf) = 0 Then
        visible = False
    Else
        visible = True
    End If
    
End Sub
Public Sub ContextMenus_GetContent(control As IRibbonControl, ByRef returnedVal)

    Dim D As Object
    Dim elmMenu As Object
    Dim elmButton As Object
    Dim i As Long
    Dim j As Long
    Dim lngCount As Long
    Dim varRow As Variant
    Dim varCol As Variant
    Dim strBuf As String
    
    '改ページプレビューも一緒に判定
    strBuf = GetSetting(C_TITLE, "ContextMenu", Replace(control.id, "Layout", ""), "")
    
    If Len(strBuf) <> 0 Then
    
        Set D = CreateObject("Msxml2.DOMDocument")
        
        Set elmMenu = D.createElement("menu")
'        elmMenu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2006/01/customui"
        elmMenu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  
        varRow = Split(strBuf, vbCrLf)
        
        For j = LBound(varRow) To UBound(varRow) - 1
        
            varCol = Split(varRow(j), vbTab)
            
            If varCol(2) = "-" Then
                Set elmButton = D.createElement("menuSeparator")
                With elmButton
                  .setAttribute "id", control.id & "Sep" & j
                End With
                elmMenu.appendChild elmButton
                Set elmButton = Nothing
            Else
                Set elmButton = D.createElement("button")
                With elmButton
                  .setAttribute "id", varCol(2) & ".dyn" & j
                  .setAttribute "label", varCol(1)
                  .setAttribute "onAction", "ribbonOnAction"
                End With
                elmMenu.appendChild elmButton
                Set elmButton = Nothing
            End If
        Next
    
      D.appendChild elmMenu
    
      returnedVal = D.XML
      
      Set elmMenu = Nothing
      Set D = Nothing
      
    Else
        returnedVal = ""
    End If

End Sub



'--------------------------------------------------------------------
'リボンより受け取ったIDをそのままマクロ名として実行するラッパー関数
'--------------------------------------------------------------------
Public Sub ribbonOnFastPin(control As IRibbonControl)

    
    On Error GoTo e
    
    Dim strBook As String
    
    strBook = GetSetting(C_TITLE, "FastPin", control.id, "")
    
    If strBook = "" Then
        Exit Sub
    End If
    
    
    Select Case True
        Case rlxIsExcelFile(strBook)
            If Not rlxIsFileExists(strBook) Then
                MsgBox "ブックが存在しません。", vbOKOnly + vbExclamation, C_TITLE
            Else
                On Error Resume Next
                Err.Clear
                Workbooks.Open filename:=strBook
                If Err.Number <> 0 Then
                    MsgBox "ブックを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
                End If
                AppActivate Application.Caption
            End If
        
        Case rlxIsPowerPointFile(strBook)
            On Error Resume Next
            Err.Clear
            With CreateObject("PowerPoint.Application")
                .visible = True
                Call .Presentations.Open(filename:=strBook)
                If Err.Number <> 0 Then
                    MsgBox "ファイルを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
                End If
                AppActivate .Caption
            End With
            
        Case rlxIsWordFile(strBook)
            On Error Resume Next
            Err.Clear
            With CreateObject("Word.Application")
                .visible = True
                .Documents.Open filename:=strBook
                If Err.Number <> 0 Then
                    MsgBox "ファイルを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
                End If
                AppActivate .Caption
            End With
            
        Case Else
            On Error Resume Next
            Dim WSH As Object
            Set WSH = CreateObject("WScript.Shell")
            
            WSH.Run ("""" & strBook & """")
             If Err.Number <> 0 Then
                MsgBox "ファイルを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
            End If
            Set WSH = Nothing
    End Select
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' ヘルプ内容を表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub getFastPinSupertip(control As IRibbonControl, ByRef Screentip)

    On Error GoTo e

    Dim strBuf As String

    strBuf = GetSetting(C_TITLE, "FastPin", control.id, "")
    Screentip = rlxGetFullpathFromPathName(strBuf)
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' メニュー表示内容を表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub getFastPinDescription(control As IRibbonControl, ByRef Screentip)

    GetDescription control, Screentip

End Sub
'--------------------------------------------------------------------
' ラベルを表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub getFastPinLabel(control As IRibbonControl, ByRef Screentip)

    On Error GoTo e
    
    Dim strBuf As String
    
    strBuf = GetSetting(C_TITLE, "FastPin", control.id, "")
    
    If strBuf = "" Then
        Screentip = "(未登録)"
    Else
        Screentip = rlxGetFullpathFromFileName(strBuf)
    End If
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

'--------------------------------------------------------------------
' リボン表示設定取得
'--------------------------------------------------------------------
Sub getFastPinVisible(control As IRibbonControl, ByRef visible)

    On Error GoTo e
    
    Dim strBuf As String
    
    strBuf = GetSetting(C_TITLE, "FastPin", control.id, "")
    
    If strBuf = "" Then
        visible = False
    Else
        visible = True
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
