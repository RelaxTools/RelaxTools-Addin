VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "バグ報告"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClip_Click()

    SetClipText txtEdit.Text

End Sub

Private Sub UserForm_Initialize()

    Dim strBuf As String
    
    strBuf = getVersionInfo & vbCrLf & vbCrLf

    strBuf = strBuf & "上記の情報とエラー内容をお知らせください。" & vbCrLf
    strBuf = strBuf & "以下、３つの方法があります。" & vbCrLf & vbCrLf
    
    strBuf = strBuf & "◆GitHub Issue(GitHubのアカウントが必要です)：" & vbCrLf
    strBuf = strBuf & "https://github.com/RelaxTools/RelaxTools-Addin/issues" & vbCrLf & vbCrLf
    
    strBuf = strBuf & "◆掲示板：" & vbCrLf
    strBuf = strBuf & "http://software.opensquare.net/relaxtools/bbs/wforum.cgi" & vbCrLf & vbCrLf
    
    strBuf = strBuf & "◆メール(relaxtools@opensquare.net)でも受け付けます。" & vbCrLf

    txtEdit.Text = strBuf

End Sub
