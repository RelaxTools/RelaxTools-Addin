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
    
    getSheetItem = ""
    
    strBuf = getMacroName(control)
    
    i = C_START_ROW
    
    Do Until ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_NO).value = ""
        If strBuf = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_MACRO).value Then
            getSheetItem = ThisWorkbook.Worksheets("HELP").Cells(i, lngItem).value
            Exit Do
        End If
        i = i + 1
    Loop

End Function
'--------------------------------------------------------------------
' リボン表示設定取得
'--------------------------------------------------------------------
Sub tabGetVisible(control As IRibbonControl, ByRef visible)

    visible = GetSetting(C_TITLE, "Ribbon", Replace(control.id, "Tab", ""), True)

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
    Call rlxErrMsg(err)
End Sub
'--------------------------------------------------------------------
'チェックボックス設定取得
'--------------------------------------------------------------------
Public Sub CheckGetPressed(control As IRibbonControl, ByRef returnValue)
    
    On Error GoTo e
    
    returnValue = GetSetting(C_TITLE, "Backup", "Check", False)
    
    Exit Sub
e:
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
End Sub
'--------------------------------------------------------------------
' メニュー表示内容を表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub GetDescription(control As IRibbonControl, ByRef Screentip)

    On Error GoTo e
    
    Screentip = getSheetItem(control, C_COL_DESCRIPTION)

    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub
'--------------------------------------------------------------------
' ラベルを表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub GetLabel(control As IRibbonControl, ByRef Screentip)

    On Error GoTo e
    
    Screentip = getSheetItem(control, C_COL_LABEL)
    
    Exit Sub
e:
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
End Sub
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
    Call rlxErrMsg(err)
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
    Call rlxErrMsg(err)
End Sub

'--------------------------------------------------------------------
'  リボンサイズ取得(未動作)
'--------------------------------------------------------------------
 Sub GetSize(control As IRibbonControl, ByRef Size)
 
    If Application.UsableWidth / 0.75 < 1420 Then
    
        Size = RibbonControlSize.RibbonControlSizeRegular
    Else
    
        Size = RibbonControlSize.RibbonControlSizeLarge
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
    Dim col As Collection
    
    Set col = getProperty()

    count = col.count

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
    Dim col As Collection
    
    Set col = getPropertyMitome()

    count = col.count

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
    Dim col As Collection
    
    Set col = getPropertyBz()

    count = col.count

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
    Call rlxErrMsg(err)
End Sub
'--------------------------------------------------------------------
'  十字カーソル設定のEnabled/Disabled
'--------------------------------------------------------------------
Sub getCrossEnabled(control As IRibbonControl, ByRef enabled)

    On Error GoTo e
    
    enabled = Not (mLineEnable)

    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub
'--------------------------------------------------------------------
'  ホイール量設定のEnabled/Disabled
'--------------------------------------------------------------------
Sub getScrollEnabled(control As IRibbonControl, ByRef enabled)

    On Error GoTo e
    
    enabled = Not (mScrollEnable)

    Exit Sub
e:
    Call rlxErrMsg(err)
End Sub
