VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGitResult 
   Caption         =   "GitResult"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16080
   OleObjectBlob   =   "frmGitResult.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmGitResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Start(ByVal strMsg As String)

    txtResult.Text = strMsg
    txtResult.SelStart = Len(txtResult.Text)
    txtResult.SelStart = 0
    
    Me.Show

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

