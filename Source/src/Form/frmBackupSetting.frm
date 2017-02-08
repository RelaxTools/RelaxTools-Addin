VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBackupSetting 
   Caption         =   "簡易世代管理対象ファイルパターン設定"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   OleObjectBlob   =   "frmBackupSetting.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmBackupSetting"
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

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2

Private Const C_FILE_NO As Long = 0
Private Const C_FILE_STR As Long = 1
Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdDel_Click()

    Dim i As Long
    
    For i = 0 To lstResult.ListCount
    
        If lstResult.Selected(i) Then
            lstResult.RemoveItem i
            Exit Sub
        End If
    
    Next

End Sub

Private Sub cmdFolder_Click()

    Dim strFile As String

    'フォルダ名取得
    strFile = rlxSelectFolder()
    
    If Trim(strFile) <> "" Then
        txtFolder.Text = strFile
    End If
    
End Sub

Private Sub cmdOK_Click()

    Dim i As Long
    Dim strFile As String
    Dim strList As String
    
    strFile = LCase(txtFileName.Text)
    
    For i = 0 To lstResult.ListCount - 1
        strList = LCase(lstResult.List(i, C_FILE_STR))
        If strFile = strList Then
            MsgBox "すでに同様のファイルパターンが登録されています。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    Next
    
    i = lstResult.ListCount

    lstResult.AddItem ""
    lstResult.List(i, C_FILE_NO) = i + 1
    lstResult.List(i, C_FILE_STR) = txtFileName.Text

End Sub




Private Sub cmdSave_Click()

    Dim strList As String
    Dim i As Long
    
    If Len(txtFolder.Text) <> 0 Then
        If Not rlxIsFolderExists(txtFolder.Text) Then
            MsgBox "指定されたフォルダは存在しません。", vbOKOnly + vbExclamation, C_TITLE
            txtFolder.SetFocus
            Exit Sub
        End If
    End If
    
    Select Case Val(txtGen.Text)
        Case 1 To 999
        Case Else
            MsgBox "世代数には1～999を入力してください。", vbOKOnly + vbExclamation, C_TITLE
            txtGen.SetFocus
            Exit Sub
    End Select
    
    strList = ""
    For i = 0 To lstResult.ListCount - 1
        If strList = "" Then
            strList = lstResult.List(i, C_FILE_STR)
        Else
            strList = strList & vbTab & lstResult.List(i, C_FILE_STR)
        End If
    Next

    SaveSetting C_TITLE, "Backup", "FileList", strList
    SaveSetting C_TITLE, "Backup", "Folder", txtFolder.Text
    SaveSetting C_TITLE, "Backup", "Gen", txtGen.Text
    
    Unload Me
    
End Sub


'------------------------------------------------------------------------------------------------------------------------
' 選択行を上に移動
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdUp_Click()
     Call moveList(C_UP)
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 選択行を下に移動
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdDown_Click()
     Call moveList(C_DOWN)
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 移動処理
'------------------------------------------------------------------------------------------------------------------------
Private Sub moveList(ByVal lngMode As Long)

    Dim lngCnt As Long
    Dim lngCmp As Long
    
    Dim varTmp As Variant

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngInc As Long

    '１つなら不要
    If lstResult.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstResult.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstResult.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select

    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstResult.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If
            
            lngCmp = lngCnt + lngInc * -1
            
            Dim i As Long
            For i = C_FILE_STR To C_FILE_STR
                varTmp = lstResult.List(lngCnt, i)
                lstResult.List(lngCnt, i) = lstResult.List(lngCmp, i)
                lstResult.List(lngCmp, i) = varTmp
            Next
            
            lstResult.Selected(lngCnt) = False
            lstResult.Selected(lngCnt + lngInc * -1) = True
        End If
    
    Next

End Sub


Private Sub lstResult_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstResult
End Sub

Private Sub UserForm_Initialize()
        
    Dim strResult As String
    Dim strList() As String
    Dim i As Long
    
    txtFileName.Text = Application.ActiveWorkbook.FullName

    strResult = GetSetting(C_TITLE, "Backup", "FileList", "")
    txtFolder.Text = GetSetting(C_TITLE, "Backup", "Folder", "")
    txtGen.Text = GetSetting(C_TITLE, "Backup", "Gen", "99")

    strList = Split(strResult, vbTab)

    For i = 0 To UBound(strList)
        lstResult.AddItem ""
        lstResult.List(i, C_FILE_NO) = i + 1
        lstResult.List(i, C_FILE_STR) = strList(i)
    Next
   
    Me.Top = (Application.Top + Application.Height - Me.Height) - 20
    Me.Left = (Application.Left + Application.Width - Me.Width) - 20
    
    Set MW = basMouseWheel.GetInstance
    MW.Install Me
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = Nothing
End Sub

Private Sub UserForm_Terminate()
    MW.UnInstall
    Set MW = Nothing
End Sub
Private Sub MW_WheelDown(obj As Object)

    If obj.ListCount = 0 Then Exit Sub
    obj.TopIndex = obj.TopIndex + 3
    
End Sub

Private Sub MW_WheelUp(obj As Object)

    Dim lngPos As Long

    If obj.ListCount = 0 Then Exit Sub
    lngPos = obj.TopIndex - 3

    If lngPos < 0 Then
        lngPos = 0
    End If

    obj.TopIndex = lngPos

End Sub
