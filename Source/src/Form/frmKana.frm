VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKana 
   Caption         =   "ひらがな⇔カタカナ変換の設定"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10710
   OleObjectBlob   =   "frmKana.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmKana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()
    SaveSetting C_TITLE, "KatakanaConv", "kake", chkKatakanaConv.Value
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    chkKatakanaConv.Value = GetSetting(C_TITLE, "KatakanaConv", "kake", False)
End Sub
