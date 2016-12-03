VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSushi 
   Caption         =   "スシ"
   ClientHeight    =   645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
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
    Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hwnd As LongPtr, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Private Declare PtrSafe Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#End If
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private mblnUnload As Boolean

Private Sub UserForm_Initialize()

    Dim OvalSet As Long
    Dim rc As Long
#If VBA7 And Win64 Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If
    Dim x As Single
    Dim y As Single
    
    hwnd = FindWindowA("ThunderDFrame", Me.Caption)

    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) And Not WS_CAPTION

    x = 5
    y = 3
    OvalSet = CreateEllipticRgn(x, y, x + 35, y + 35)
    rc = SetWindowRgn(hwnd, OvalSet, True)
    
End Sub
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


