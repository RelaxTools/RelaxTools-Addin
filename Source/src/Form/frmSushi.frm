VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSushi 
   Caption         =   "スシ"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4380
   OleObjectBlob   =   "frmSushi.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmSushi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hWnd As LongPtr, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Private Declare PtrSafe Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Sub ReleaseCapture Lib "user32.dll" ()
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Sub ReleaseCapture Lib "user32.dll" ()

#End If
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private mblnUnload As Boolean
#If VBA7 And Win64 Then
    Private hWnd As LongPtr
#Else
    Private hWnd As Long
#End If
Private Sub imgSushi_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ReleaseCapture
    Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub UserForm_Initialize()

    Dim OvalSet As Long
    Dim rc As Long

    Dim X As Single
    Dim Y As Single
    
    hWnd = FindWindowA("ThunderDFrame", Me.Caption)

    SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
    SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_CAPTION

    X = 5
    Y = 3
    OvalSet = CreateEllipticRgn(X, Y, X + 35, Y + 35)
    rc = SetWindowRgn(hWnd, OvalSet, True)
    
End Sub

Public Property Let Neta(v As String)
    
    Dim c As Object
    For Each c In Controls
        If c.tag = v Then
            imgSushi.Picture = c.Picture
        End If
    Next
End Property



'
'Private Sub UserForm_Activate()
'    Dim i As Long
'    Dim lngTopMergin As Long
'    Dim lngLeftMergin As Long
'    Dim lngMove As Long
'    Dim lngWait As Long
'
'    '移動量
'    lngMove = 1
'
'    '待ち
'    lngWait = 10
'
'    lngTopMergin = 10
'    lngLeftMergin = 40
'
'    Me.Left = Application.Left
'
'    '→
'    Do
'        DoEvents
'        Sleep lngWait
'        Me.Left = Me.Left + lngMove
'        Me.Top = Application.Top + Application.Height - lngLeftMergin
'        If Me.Left > (Application.Left + Application.Width - 36) Then
'            Exit Do
'        End If
'        If mblnUnload Then
'            Exit Sub
'        End If
'    Loop
'
'    '↑
'    Do
'        DoEvents
'        Sleep lngWait
'        Me.Top = Me.Top - lngMove
'        Me.Left = Application.Left + Application.Width - lngLeftMergin
'        If Me.Top < (Application.Top + lngTopMergin) Then
'            Exit Do
'        End If
'        If mblnUnload Then
'            Exit Sub
'        End If
'    Loop
'
'    '←
'    Do
'        DoEvents
'        Sleep lngWait
'        Me.Left = Me.Left - lngMove
'        Me.Top = Application.Top + lngTopMergin
'        If Me.Left < (Application.Left + lngTopMergin) Then
'            Exit Do
'        End If
'        If mblnUnload Then
'            Exit Sub
'        End If
'    Loop
'
'    '↓
'    Do
'        DoEvents
'        Sleep lngWait
'        Me.Top = Me.Top + lngMove
'        Me.Left = Application.Left + lngTopMergin
'        If Me.Top > (Application.Top + Application.Height - lngLeftMergin) Then
'            Exit Do
'        End If
'        If mblnUnload Then
'            Exit Sub
'        End If
'    Loop
'
'
'    Unload Me
'End Sub
'Private Sub imgSushi_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'    Unload Me
'    mblnUnload = True
'End Sub
'Private Sub UserForm_Click()
'    Unload Me
'    mblnUnload = True
'End Sub


