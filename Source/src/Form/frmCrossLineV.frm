VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCrossLineV 
   Caption         =   "frmCrossLineV"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240.001
   OleObjectBlob   =   "frmCrossLineV.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmCrossLineV"
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

Private Sub UserForm_Resize()

'    Image1.width = Me.width
'    Image1.Height = Me.Height

End Sub
