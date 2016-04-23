VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaskTrayView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
 
#If VBA7 And Win64 Then

    Private Declare PtrSafe Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Private Declare PtrSafe Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As LongPtr, phiconSmall As LongPtr, ByVal nIcons As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    
    Private Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
    End Type

    Private Type NOTIFYICONDATA
        cbsize As Long
        hWnd As LongPtr
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As LongPtr
        szTip As String * 128
        dwState As Long
        dwStateMask As Long
        szInfo As String * 256
        uTimeoutOrVersion As Long
        szInfoTitle As String * 64
        dwInfoFlags As Long
        guidItem As GUID
        hBalloonIcon As LongPtr
    End Type
    
#Else

    Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    
    Private Type NOTIFYICONDATA
        cbsize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 128
        dwState As Long
        dwStateMask As Long
        szInfo As String * 256
        uTimeoutOrVersion As Long
        szInfoTitle As String * 64
        dwInfoFlags As Long
    End Type
    
#End If

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
 
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIF_GUID = &H20

Private typ_Icon As NOTIFYICONDATA '自クラスで管理するアイコンのデータ(構造体)
#If VBA7 And Win64 Then
    Public Function AddIcon(ByVal hWnd As LongPtr, ByVal ToolTip As String) As Long
     
        Dim IconLarge As LongPtr
        Dim IconSmall As LongPtr
        Dim IconIndex As Long
     
        Dim strFile As String
        
        strFile = Application.Path & "\" & "EXCEL.EXE"
        IconIndex = 1
        
        Call ExtractIconEx(strFile, IconIndex, IconLarge, IconSmall, 1)
    
        With typ_Icon
            .cbsize = Len(typ_Icon)
            .hWnd = hWnd
            .uID = 0
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
            .hIcon = IconSmall
            .szTip = ToolTip & vbNullChar
        End With
         
        DrawMenuBar hWnd
         
        AddIcon = Shell_NotifyIcon(NIM_ADD, typ_Icon)
     
    End Function
#Else
    Public Function AddIcon(ByVal hWnd As Long, ByVal ToolTip As String) As Long
     
        Dim IconLarge As Long
        Dim IconSmall As Long
        Dim IconIndex As Long
     
        Dim strFile As String
        
        strFile = Application.Path & "\" & "EXCEL.EXE"
        IconIndex = 1
        
        Call ExtractIconEx(strFile, IconIndex, IconLarge, IconSmall, 1)
    
        With typ_Icon
            .cbsize = Len(typ_Icon)
            .hWnd = hWnd
            .uID = 0
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
            .hIcon = IconSmall
            .szTip = ToolTip & vbNullChar
        End With
         
        DrawMenuBar hWnd
     
        AddIcon = Shell_NotifyIcon(NIM_ADD, typ_Icon)
        
    End Function
#End If
Public Function ShowBalloon(ByVal Message As String, Optional ByVal BalloonTimeOutSecond As Long = 5) As Long

    With typ_Icon
        .cbsize = Len(typ_Icon)
        .uFlags = NIF_INFO
        .szInfo = Message & vbNullChar
        .uTimeoutOrVersion = BalloonTimeOutSecond * 1000 'タイムアウト時間(ms)
    End With
 
    ShowBalloon = Shell_NotifyIcon(NIM_MODIFY, typ_Icon)
     
End Function
Public Sub DeleteIcon()

    With typ_Icon
        .cbsize = Len(typ_Icon)
        .uFlags = NIF_ICON
        .uID = 0
    End With
     
    Call Shell_NotifyIcon(NIM_DELETE, typ_Icon) 'API実行
     
End Sub
Private Sub Class_Terminate()
    Dim typ_Dummy As NOTIFYICONDATA
    Call DeleteIcon
    typ_Icon = typ_Dummy
End Sub

