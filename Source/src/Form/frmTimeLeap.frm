VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimeLeap 
   Caption         =   "TimeLeap"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14055
   OleObjectBlob   =   "frmTimeLeap.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTimeLeap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
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

'
Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1

Private Const C_FILE_NO As Long = 0
Private Const C_FILE_REV As Long = 1
Private Const C_FILE_MODIFY As Long = 2
Private Const C_FILE_NAME As Long = 3
Private Const C_FILE_SIZE As Long = 4
Private Const C_FILE_ORIGINAL As Long = 5
Private Const C_FILE_COMPARE As Long = 6


Private mBarFav As Object


Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2


Private Const C_HEAD As Long = 3
Private Const C_TAIL As Long = 4

Private Const C_FILE_INFO As String = "ファイル情報："
Private mstrBook As String


'
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = Nothing
End Sub
Private Sub lstTimeLeap_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstTimeLeap
End Sub
Private Sub lstTimeLeap_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Select Case KeyCode
'        Case vbKeyV
'            If (Shift And 2) Then
'                Call favPaste
'                Exit Sub
'            End If
'        Case vbKeyC
'            If (Shift And 2) Then
'                Call favCopy
'                Exit Sub
'            End If
'        Case vbKeyA
'            If (Shift And 2) Then
'                Call favAllSelect
'                Exit Sub
'            End If
        Case vbKeyEscape
            Unload Me
            Exit Sub
'        Case vbKeyReturn
'            execOpen False
'            Exit Sub
'        Case vbKeyLeft
'            lstCategory.SetFocus
'            Exit Sub
'        Case vbKeyDelete
'            execDel
'            Exit Sub
    End Select


End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
'        Case vbKeyV
'            If (Shift And 2) Then
'                Call favPaste
'            End If
        Case vbKeyEscape
            Unload Me
'        Case vbKeyReturn
'            execOpen False

    End Select
End Sub


Public Sub execOpen(ByVal blnReadOnly As Boolean)

    Dim strBook As String
    Dim lngCnt As Long

    Dim varFile As Variant

    If lstTimeLeap.ListIndex = -1 Then
        Exit Sub
    End If

    On Error Resume Next
    Me.Hide
    
    For lngCnt = 0 To lstTimeLeap.ListCount - 1

        If lstTimeLeap.Selected(lngCnt) Then

            strBook = lstTimeLeap.List(lngCnt, C_FILE_ORIGINAL)
                        
            On Error Resume Next
            Err.Clear
            Workbooks.Open filename:=strBook, ReadOnly:=blnReadOnly, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True
            If Err.Number <> 0 Then
                MsgBox "ブックを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
            End If
            AppActivate Application.Caption
        End If

    Next

    Unload Me

End Sub
Sub execCompare()

    Dim strBook As String
    Dim strBook2 As String
    Dim lngCnt As Long
    Dim strDiff As String

    Dim varFile As Variant

    If lstTimeLeap.ListIndex <= 0 Then
        Exit Sub
    End If

    strBook = lstTimeLeap.List(0, C_FILE_ORIGINAL)
    strBook2 = lstTimeLeap.List(lstTimeLeap.ListIndex, C_FILE_ORIGINAL)
    
    strDiff = GetSetting(C_TITLE, "TimeLeap", "Diff", "")
    strDiff = Replace(strDiff, "$(SRC)", """" & strBook & """")
    strDiff = Replace(strDiff, "$(DEST)", """" & strBook2 & """")
    
    On Error Resume Next
    Err.Clear
    With CreateObject("WScript.Shell")
        Call .Run(strDiff, 1, True)
    End With
    If Err.Number <> 0 Then
        MsgBox "DIFFの起動に失敗しました。設定を確認してください。", vbOKOnly + vbExclamation, C_TITLE
    End If

End Sub
Sub execCompareBefore()

    Dim strBook As String
    Dim strBook2 As String
    Dim lngCnt As Long
    Dim strDiff As String

    Dim varFile As Variant

    If lstTimeLeap.ListIndex <= 0 Then
        Exit Sub
    End If

    strBook2 = lstTimeLeap.List(lstTimeLeap.ListIndex, C_FILE_ORIGINAL)
    strBook = lstTimeLeap.List(lstTimeLeap.ListIndex - 1, C_FILE_ORIGINAL)
    
    strDiff = GetSetting(C_TITLE, "TimeLeap", "Diff", "")
    strDiff = Replace(strDiff, "$(SRC)", """" & strBook & """")
    strDiff = Replace(strDiff, "$(DEST)", """" & strBook2 & """")
    
    On Error Resume Next
    Err.Clear
    With CreateObject("WScript.Shell")
        Call .Run(strDiff, 1, True)
    End With
    If Err.Number <> 0 Then
        MsgBox "DIFFの起動に失敗しました。設定を確認してください。", vbOKOnly + vbExclamation, C_TITLE
    End If

End Sub
Private Sub lstTimeLeap_Change()

    Dim strMsg As String
    Dim strBook As String
    Dim i As Long
    Dim strCat As String

    If lstTimeLeap.ListIndex = -1 Then
        Exit Sub
    End If

    On Error Resume Next

    strMsg = C_FILE_INFO

    Dim lngCount As Long

    lngCount = 0
    For i = 0 To lstTimeLeap.ListCount - 1
        If lstTimeLeap.Selected(i) Then
            lngCount = lngCount + 1
        End If
    Next

    If lngCount = 1 Then

        Dim shell As Object, folder As Object

        strBook = lstTimeLeap.List(lstTimeLeap.ListIndex, C_FILE_ORIGINAL)

'        If rlxIsFileExists(strBook) Or rlxIsFolderExists(strBook) Then


            strMsg = strMsg & vbCrLf
            strMsg = strMsg & "　フォルダ名：" & rlxGetFullpathFromPathName(strBook) & vbCrLf           ''ファイル名
            strMsg = strMsg & "　ファイル名：" & rlxGetFullpathFromFileName(strBook) & vbCrLf           ''ファイル名

'            If GetSetting(C_TITLE, "Favirite", "Detail", False) Then
'                Set Shell = CreateObject("Shell.Application")
'                Set Folder = Shell.Namespace(rlxGetFullpathFromPathName(strBook))
'                strMsg = strMsg & "　作成者：" & Folder.GetDetailsOf(Folder.ParseName(rlxGetFullpathFromFileName(strBook)), 20) & vbCrLf  ''作成者
'                strMsg = strMsg & "　タイトル：" & Folder.GetDetailsOf(Folder.ParseName(rlxGetFullpathFromFileName(strBook)), 21) & vbCrLf   ''タイトル
'                strMsg = strMsg & "　サブタイトル：" & Folder.GetDetailsOf(Folder.ParseName(rlxGetFullpathFromFileName(strBook)), 22) & vbCrLf   ''サブタイトル
'                Set Folder = Nothing
'                Set Shell = Nothing
'            End If

            Select Case True
                Case rlxIsExcelFile(strBook)
                    strCat = "Excelファイル"
                Case rlxIsPowerPointFile(strBook)
                    strCat = "PowerPointファイル"
                Case rlxIsWordFile(strBook)
                    strCat = "Wordファイル"
                Case Else
                    strCat = "その他"
            End Select

            strMsg = strMsg & "　種類：" & strCat

'        Else
'            strMsg = strMsg & vbCrLf
'            strMsg = strMsg & "　フォルダまたはファイルがありません。" & vbCrLf
'        End If
    Else

        For i = 0 To lstTimeLeap.ListCount - 1

            If lstTimeLeap.Selected(i) Then

                strBook = lstTimeLeap.List(i, C_FILE_ORIGINAL)
                strMsg = strMsg & vbCrLf & "  " & strBook

                Select Case True
                    Case rlxIsExcelFile(strBook)
                        strCat = "Excelファイル"
                    Case rlxIsPowerPointFile(strBook)
                        strCat = "PowerPointファイル"
                    Case rlxIsWordFile(strBook)
                        strCat = "Wordファイル"
                    Case Else
                        strCat = "その他"
                End Select

                strMsg = strMsg & ", " & strCat

            End If
        Next
    End If

    txtDetail.Text = strMsg
End Sub

Public Sub execOverwrite()

    Dim strBook As String
    Dim strBook2 As String
    Dim lngCnt As Long
    Dim a As FileTime
    Dim DateCreated As Date
    Dim DateLastModified As Date
    Dim varFile As Variant
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook

    If lstTimeLeap.ListIndex = -1 Then
        Exit Sub
    End If
    
    If Not WB.Saved Then
        If MsgBox("ブックが変更されています。破棄しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            MsgBox "処理を中断しました。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    End If
    
    If MsgBox("最新のブックを上書きします。元に戻せませんがよろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    On Error GoTo e
    
    strBook = lstTimeLeap.List(lstTimeLeap.ListIndex, C_FILE_ORIGINAL)
    strBook = rlxGetFullpathFromExt(strBook) & "." & Format(Val(Right$(strBook, 3)) + 1, "000")
    
    strBook2 = WB.FullName
    
    Unload Me
    
    '履歴作成
    Call CreateHistory(strBook2)
    
    WB.Close SaveChanges:=False
        
    Set a = New FileTime
    With CreateObject("Scripting.FileSystemObject")
    
        DateCreated = .GetFile(strBook).DateCreated
        DateLastModified = .GetFile(strBook).DateLastModified
        
        .CopyFile strBook, strBook2
        
        a.SetCreationTime strBook2, DateCreated
        a.SetLastWriteTime strBook2, DateLastModified
    
    End With
        
    Workbooks.Open strBook2

    MsgBox "正常に上書きされました。", vbOKOnly + vbInformation, C_TITLE

    Exit Sub
e:
    MsgBox "復元に失敗しました。", vbOKOnly + vbInformation, C_TITLE

End Sub
Public Sub execExport()

    Dim strFile As Variant
    Dim strBook As String
    Dim strExt As String
    
    Dim a As FileTime
    Dim DateCreated As Date
    Dim DateLastModified As Date

    If lstTimeLeap.ListIndex = -1 Then
        Exit Sub
    End If
    
    strBook = lstTimeLeap.List(lstTimeLeap.ListIndex, C_FILE_ORIGINAL)
    
    strExt = rlxGetFullpathFromExt(strBook)
    strExt = "*" & Mid$(strExt, InStrRev(strExt, "."))
    
    strFile = Application.GetSaveAsFilename(InitialFileName:=rlxGetFullpathFromFileName(rlxGetFullpathFromExt(strBook)), Title:="エクスポート", fileFilter:="Excel Files (" & strExt & "), " & strExt & "")
    If strFile = "False" Then
        Exit Sub
    End If
        
    Set a = New FileTime
    
    With CreateObject("Scripting.FileSystemObject")
    
        DateCreated = .GetFile(strBook).DateCreated
        DateLastModified = .GetFile(strBook).DateLastModified
        
        .CopyFile strBook, strFile
        
        a.SetCreationTime strFile, DateCreated
        a.SetLastWriteTime strFile, DateLastModified
        
    End With
    
    MsgBox "エクスポートしました。", vbOKOnly + vbInformation, C_TITLE

End Sub
'

'
Private Sub lstTimeLeap_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If lstTimeLeap.ListIndex = 0 Then
        Exit Sub
    End If


    If Button = 2 Then

        Set mBarFav = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
        With mBarFav

            With .Controls.Add
                .Caption = "開く（読み取り専用）"
                .OnAction = "'basTimeLeap.execOpen(""" & True & """)'"
                .FaceId = 456
            End With
            With .Controls.Add
                .BeginGroup = True
                .Caption = "最新のブックと比較"
                .OnAction = "basTimeLeap.execCompare"
                .FaceId = 19
            End With
            With .Controls.Add
                .Caption = "直前のブックと比較"
                .OnAction = "'basTimeLeap.execCompareBefore'"
                .FaceId = 19
            End With
            With .Controls.Add
                .BeginGroup = True
                .Caption = "エクスポート"
                .OnAction = "'basTimeLeap.execExport'"
                .FaceId = 526
            End With
            With .Controls.Add
                .Caption = "このブックを最新のブックに上書き"
                .OnAction = "'basTimeLeap.execOverwrite'"
                .FaceId = 534
            End With


        End With
        mBarFav.ShowPopup

    End If
End Sub


Private Sub MW_WheelDown(obj As Object)

    On Error GoTo e

    If obj.ListCount = 0 Then Exit Sub
    obj.TopIndex = obj.TopIndex + 3
e:
End Sub

Private Sub MW_WheelUp(obj As Object)

    Dim lngPos As Long

    On Error GoTo e

    If obj.ListCount = 0 Then Exit Sub
    lngPos = obj.TopIndex - 3

    If lngPos < 0 Then
        lngPos = 0
    End If

    obj.TopIndex = lngPos
e:
End Sub


Sub Disp()

    Dim strFolder As String
    Dim objfld As Object
    Dim objFile As Object
    Dim obj As Object
    Dim i As Long
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    Dim col As Collection
    Set col = New Collection
    
    mstrBook = WB.FullName
    
    Me.Caption = "TimeLeap /  " & WB.Name & ""
    
    
    With CreateObject("Scripting.FileSystemObject")
    
        If .FileExists(WB.FullName) Then
            col.Add .GetFile(WB.FullName)
        End If
        
        strFolder = GetSetting(C_TITLE, "TimeLeap", "Folder", GetTimeLeapFolder())
        If Not .FolderExists(strFolder) Then
            .createFolder strFolder
        End If
        Set objfld = .GetFolder(strFolder)
    
        For Each objFile In objfld.files
            If rlxGetFullpathFromExt(objFile.Name) = WB.Name Then
                col.Add objFile
            End If
        Next
    End With
    
    Dim strBeforeSize As String
    Dim strBeforeTime As String
    Dim strAfterSize As String
    Dim strAfterTime As String
    
    strBeforeSize = ""
    strBeforeTime = ""
    
    For Each objFile In col
    
        Dim strSize As String
        strSize = Format$(objFile.size, "#,##0")
        
        lstTimeLeap.AddItem ""
        If i = 0 Then
            lstTimeLeap.List(i, C_FILE_NO) = "-"
        Else
            lstTimeLeap.List(i, C_FILE_NO) = i
        End If
        Select Case i
            Case 0
                lstTimeLeap.List(i, C_FILE_REV) = "最新"
                lstTimeLeap.List(i, C_FILE_NAME) = "最新のブック (最終更新者：" & WB.BuiltinDocumentProperties.Item(7) & ")"
            
            Case 1
                lstTimeLeap.List(i, C_FILE_REV) = Right$(objFile.Name, 3)
                lstTimeLeap.List(i, C_FILE_NAME) = "前回保存した際のブックのクローン。"
            Case Else
                lstTimeLeap.List(i, C_FILE_REV) = Right$(objFile.Name, 3)
                lstTimeLeap.List(i, C_FILE_NAME) = timeAfter(CDate(objFile.DateLastModified), Now) & "のブック"
        End Select
        
        strAfterSize = String$(11 - Len(strSize), " ") & strSize
        strAfterTime = Format(objFile.DateLastModified, "yyyy-mm-dd hh:nn:ss")
        
        If strAfterSize <> strBeforeSize Then
            lstTimeLeap.List(i, C_FILE_SIZE) = strAfterSize
            strBeforeSize = strAfterSize
        Else
            lstTimeLeap.List(i, C_FILE_SIZE) = "          同上"
        End If
        
        If strAfterTime <> strBeforeTime Then
            lstTimeLeap.List(i, C_FILE_MODIFY) = strAfterTime
            strBeforeTime = strAfterTime
        Else
            lstTimeLeap.List(i, C_FILE_MODIFY) = "             同上"
        End If
        
        lstTimeLeap.List(i, C_FILE_ORIGINAL) = objFile.Path
        lstTimeLeap.List(i, C_FILE_COMPARE) = objFile.Path
        i = i + 1
                
    Next
    
    If lstTimeLeap.ListCount > 0 Then
        lstTimeLeap.ListIndex = 0
    End If
    
    Dim lngCnt As Long
    
    If lstTimeLeap.ListCount = 0 Then
        lngCnt = 0
    Else
        lngCnt = lstTimeLeap.ListCount - 1
    End If
    
    lblMsg.Caption = "TimeLeapフォルダには" & lngCnt & "件のブックがあります。"
    
End Sub

Function timeAfter(ByVal s As Date, ByVal d As Date) As String

    Dim lngRet As Long
    Dim strRet As String

    '同じ日の場合
    If Format(s, "yyyymmdd") = Format(d, "yyyymmdd") Then
        lngRet = DateDiff("s", s, d)
        If lngRet > 60 Then
            lngRet = DateDiff("n", s, d)
            If lngRet > 60 Then
                lngRet = DateDiff("h", s, d)
                strRet = CStr(lngRet) & "時間前"
            Else
                strRet = CStr(lngRet) & "分前"
            End If
        Else
            strRet = CStr(lngRet) & "秒前"
        End If
    Else
        d = DateSerial(Year(d), Month(d), Day(d))
        lngRet = DateDiff("d", s, d)
        If lngRet > 31 Then
            lngRet = DateDiff("m", s, d)
            If lngRet > 12 Then
                lngRet = DateDiff("y", s, d)
                strRet = CStr(lngRet) & "年前"
            Else
                If lngRet = 1 Then
                    strRet = "先月"
                Else
                    strRet = CStr(lngRet) & "ヶ月前"
                End If
            End If
        Else
            If lngRet = 1 Then
                strRet = "昨日"
            Else
                strRet = CStr(lngRet) & "日前"
            End If
        End If
    End If

        
    timeAfter = strRet

End Function
Private Sub UserForm_Initialize()
    Disp
    Set MW = basMouseWheel.GetInstance
    MW.Install Me
End Sub

Private Sub UserForm_Terminate()
    
    MW.UnInstall
    Set MW = Nothing

End Sub
