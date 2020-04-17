VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCrossLineGauge 
   Caption         =   "ガイド"
   ClientHeight    =   870
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   1992
   OleObjectBlob   =   "frmCrossLineGauge.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCrossLineGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

#If VBA7 And Win64 Then
    Public hWnd As LongPtr
#Else
    Public hWnd As Long
#End If
Public Transparency As Double

Public Sub Run()


    
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    
    
    
    If hWnd <> 0& Then
        'キャプションなし
        SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_SYSMENU
'        SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or &H20

        'フレーム無
'        SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME

        'キャプションなし
'        SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_CAPTION

        '半透明化
'        SetLayeredWindowAttributes hWnd, 0, Transparency * 0.01 * 255, LWA_ALPHA

    End If
    
    Me.Show
    
End Sub



