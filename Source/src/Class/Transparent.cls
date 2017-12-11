VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transparent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If Win64 And VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
    Private hwnd As LongPtr
#Else
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
    Private hwnd As Long
#End If

Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2
Private Const GWL_EXSTYLE As Long = -20
Private Const LWA_COLORKEY = &H1

Sub Init(ByRef f As UserForm)
    
    Dim dwStyle As Long
    
    hwnd = FindWindowA("ThunderDFrame", f.Caption)
    dwStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    dwStyle = dwStyle Or WS_EX_LAYERED
    Call SetWindowLong(hwnd, GWL_EXSTYLE, dwStyle)
   
End Sub


Sub setTransparent(ByVal v As Long)

    Call SetLayeredWindowAttributes(hwnd, 0, v, LWA_ALPHA)

End Sub
Sub setColor(ByVal v As Long)

    Call SetLayeredWindowAttributes(hwnd, v, 255, LWA_COLORKEY)

End Sub
Sub Term()
    
    Dim dwStyle As Long
    
    dwStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    dwStyle = dwStyle Or WS_EX_LAYERED
    Call SetWindowLong(hwnd, GWL_EXSTYLE, dwStyle)
    
    '透明化を戻す-255
    Call SetLayeredWindowAttributes(hwnd, 0, 255, LWA_ALPHA)
    
End Sub



