VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFCS 
   Caption         =   "福島コンピューターシステム株式会社 -  福島に拠点を持つソフトウェア会社です。"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8940
   OleObjectBlob   =   "frmFCS.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
    Dim strBuf As String
    
    strBuf = "郡山駅から徒歩75分(5km)" & vbCrLf
'    strBuf = strBuf & "タクシーでおいでください。" & vbCrLf
    strBuf = strBuf & "自衛隊裏独立系ソフトウェア会社"
    
    lblMessage.Caption = strBuf
    
    
End Sub
