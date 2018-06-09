VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCrossLineH 
   Caption         =   "frmCrossLineH"
   ClientHeight    =   210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "frmCrossLineH.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmCrossLineH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

#If VBA7 And Win64 Then
    Public hwnd As LongPtr
#Else
    Public hwnd As Long
#End If
Public Transparency As Double

Public Sub Run()

    hwnd = FindWindow("ThunderDFrame", Me.Caption)
    
    If hwnd <> 0& Then
        SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or &H20
        
        'フレーム無
        SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
        
        'キャプションなし
        SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) And Not WS_CAPTION
        
        '半透明化
        SetLayeredWindowAttributes hwnd, 0, Transparency * 0.01 * 255, LWA_ALPHA
        
    End If
    
    Me.Show

End Sub




