Attribute VB_Name = "basScreenShot"
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

#If VBA7 And Win64 Then

    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long

    Private Declare PtrSafe Function AddClipboardFormatListener Lib "user32.dll" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function RemoveClipboardFormatListener Lib "user32.dll" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    
#Else
    
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
    
    Private Declare Function AddClipboardFormatListener Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function RemoveClipboardFormatListener Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    
#End If

Private Const GWL_WNDPROC = (-4)

Private Const CF_BITMAP = 2                          'ビットマップ
Private Const WM_CLIPBOARDUPDATE = &H31D

#If VBA7 And Win64 Then
    Private mGetHWnd As LongPtr
    Private mSetHWnd As LongPtr
#Else
    Private mGetHWnd As Long
    Private mSetHWnd As Long
#End If

Private mblnZoomEnable As Boolean
Private mlngZoomNum As Long
Private mblnSave As Boolean
Private mlngBlankNum As Long
Private mblnPageBreakEnable As Boolean
Private mlngPageBreakNun As Long

Public tray As TaskTrayView

Private WS As Worksheet
Sub GetScreenSetting(ByRef blnZoomEnable As Boolean, ByRef lngZoomNum As Long, ByRef blnSave As Boolean, ByRef lngBlankNum As Long, ByRef blnPageBreakEnable As Boolean, ByRef lngPageBreakNun As Long)

    blnZoomEnable = GetSetting(C_TITLE, "ScreenShot", "ZoomEnable", False)
    lngZoomNum = GetSetting(C_TITLE, "ScreenShot", "ZoomNum", 100)
    blnSave = GetSetting(C_TITLE, "ScreenShot", "Save", False)
    lngBlankNum = GetSetting(C_TITLE, "ScreenShot", "BlankNum", 2)
    blnPageBreakEnable = GetSetting(C_TITLE, "ScreenShot", "PageBreakEnable", False)
    lngPageBreakNun = GetSetting(C_TITLE, "ScreenShot", "PageBreakNum", 1)

End Sub
Sub SetScreenSetting(ByRef blnZoomEnable As Boolean, ByRef lngZoomNum As Long, ByRef blnSave As Boolean, ByRef lngBlankNum As Long, ByRef blnPageBreakEnable As Boolean, ByRef lngPageBreakNun As Long)

    Call SaveSetting(C_TITLE, "ScreenShot", "ZoomEnable", blnZoomEnable)
    Call SaveSetting(C_TITLE, "ScreenShot", "ZoomNum", lngZoomNum)
    Call SaveSetting(C_TITLE, "ScreenShot", "Save", blnSave)
    Call SaveSetting(C_TITLE, "ScreenShot", "BlankNum", lngBlankNum)
    Call SaveSetting(C_TITLE, "ScreenShot", "PageBreakEnable", blnPageBreakEnable)
    Call SaveSetting(C_TITLE, "ScreenShot", "PageBreakNum", lngPageBreakNun)

End Sub
Public Sub StartScreenShot()

    Dim blnZoomEnable As Boolean
    Dim lngZoomNum As Long
    Dim blnSave As Boolean
    Dim lngBlankNum As Long
    Dim blnPageBreakEnable As Boolean
    Dim lngPageBreakNun As Long

    GetScreenSetting blnZoomEnable, lngZoomNum, blnSave, lngBlankNum, blnPageBreakEnable, lngPageBreakNun
    
    mblnZoomEnable = blnZoomEnable
    mlngZoomNum = lngZoomNum
    mblnSave = blnSave
    mlngBlankNum = lngBlankNum
    mblnPageBreakEnable = blnPageBreakEnable
    mlngPageBreakNun = lngPageBreakNun
    Set WS = ActiveSheet

    mSetHWnd = frmScreenShot.hWnd
    
    'クリップボードを監視
    AddClipboardFormatListener mSetHWnd
    
    mGetHWnd = SetWindowLong(mSetHWnd, GWL_WNDPROC, AddressOf WndProc)

    Set tray = New TaskTrayView
    
    Call tray.AddIcon(Application.hWnd, "Excelスクショ")  'システムトレイにアイコンを登録
    Call tray.ShowBalloon("Excelスクショモード開始")

End Sub

Public Sub StopScreenShot()

    If mGetHWnd <> 0 Then
       Call SetWindowLong(mSetHWnd, GWL_WNDPROC, mGetHWnd)
    End If
    
    RemoveClipboardFormatListener mSetHWnd
    
    Set WS = Nothing
    
    tray.DeleteIcon
    Set tray = Nothing
    
End Sub

'-------------------------------------------------------
' ホットキー検出
'-------------------------------------------------------
#If VBA7 And Win64 Then
Public Function WndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

    Static bolWndProcCheck As Boolean

    If Not bolWndProcCheck Then
        
        bolWndProcCheck = True
        
        Select Case uMsg
            Case WM_CLIPBOARDUPDATE
                If IsClipboardFormatAvailable(CF_BITMAP) <> 0 And GetForegroundWindow() <> Application.hWnd Then
                    Application.OnTime Now, "pasteScreenShot"
                End If
        End Select
        
        bolWndProcCheck = False
    
    End If
    
    WndProc = CallWindowProc(mGetHWnd, hWnd, uMsg, wParam, lParam)

End Function
#Else
Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Static bolWndProcCheck As Boolean

    If Not bolWndProcCheck Then
        
        bolWndProcCheck = True
        
        Select Case uMsg
            Case WM_CLIPBOARDUPDATE
                If IsClipboardFormatAvailable(CF_BITMAP) <> 0 And GetForegroundWindow() <> Application.hWnd Then
                    Application.OnTime Now, "pasteScreenShot"
                End If
        End Select
        
        bolWndProcCheck = False
    
    End If
    
    WndProc = CallWindowProc(mGetHWnd, hWnd, uMsg, wParam, lParam)

End Function
#End If
Public Sub pasteScreenShot()
                        
    Dim r As Range
    On Error Resume Next
    
    If WS Is Nothing Then
        tray.ShowBalloon "貼りつけるシートが見つかりません。コピー失敗しました"
    Else
        
        '画像を張付
        WS.Paste
    
        If err.Number <> 0 Then
            tray.ShowBalloon "コピー失敗しました"
        Else
    
            ' 画像の縮小
            If mblnZoomEnable Then
                Selection.ShapeRange.ScaleHeight CDbl(mlngZoomNum) / 100, msoFalse, msoScaleFromTopLeft
            End If
        
            Dim lngRow As Long
            
            ' ブランク行を計算しておく
            lngRow = Selection.BottomRightCell.row + mlngBlankNum
        
            ' 改行の設定
            If mblnPageBreakEnable Then
                ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=WS.Cells(Selection.BottomRightCell.row + mlngPageBreakNun, Selection.TopLeftCell.Column)
            End If
        
            WS.Cells(lngRow, Selection.TopLeftCell.Column).Select
            
            '保存
            If mblnSave Then
                WS.Parent.Save
            End If
            
           
        End If

    End If
    
End Sub


