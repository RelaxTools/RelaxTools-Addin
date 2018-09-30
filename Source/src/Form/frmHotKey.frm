VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHotKey 
   Caption         =   "ショートカットキー割り当て"
   ClientHeight    =   8730.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10065
   OleObjectBlob   =   "frmHotKey.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmHotKey"
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

Const C_COM_DATA As Long = 25 '13
Const C_COM_NO As Long = 1
Const C_COM_CATEGORY As Long = 2
Const C_COM_MACRO As Long = 3
Const C_COM_DISP_NAME As Long = 4
Const C_COM_USE As Long = 5

Const C_KEY_DATA As Long = 3
Const C_KEY_NO As Long = 1
Const C_KEY_NAME As Long = 2
Const C_KEY_KEY As Long = 3

Const C_SET_DATA As Long = 3
Const C_SET_NO As Long = 1
Const C_SET_KEY As Long = 2
Const C_SET_DISP_NAME As Long = 3

Const C_SETLIST_NO As Long = 0
Const C_SETLIST_ENABLE As Long = 1
Const C_SETLIST_KEY_NAME As Long = 2
Const C_SETLIST_KEY As Long = 3
Const C_SETLIST_CATEGORY As Long = 4
Const C_SETLIST_MACRO_NAME As Long = 5
Const C_SETLIST_MACRO As Long = 6

Const C_ENABLE As String = " ○ "
Const C_DISABLE As String = " × "


Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1
Private mstrhWnd As String

Private Sub cboCategory_Click()
    Call dispCommand
End Sub



Private Sub cboCategory_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = cboCategory
End Sub

Private Sub cmbShift_Click()
    Call getGuidence
End Sub

Private Sub cmdAdd_Click()
    
    Dim j As Long
    Dim i As Long
    Dim blnFind As Boolean
    Dim strKey As String
     
    blnFind = False
    
    If lstCommand.ListCount = 0 Then
        Exit Sub
    End If
    
    strKey = cmbShift.List(cmbShift.ListIndex, 1) & lstKey.List(lstKey.ListIndex, 2)
    Select Case strKey
        Case "^%{DELETE}", "^+{ESCAPE}"
            MsgBox "システムで使用されているキーは登録できません。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
    End Select
    
    
    If lstSetting.ListCount > 0 Then
        For i = 0 To lstSetting.ListCount - 1
            strKey = cmbShift.List(cmbShift.ListIndex, 1) & lstKey.List(lstKey.ListIndex, 2)
            If lstSetting.List(i, C_SETLIST_KEY) = strKey Then
                blnFind = True
                j = i
                Exit For
            End If
        Next
    End If
    
    If blnFind Then
        If MsgBox("すでにショートカットキーが定義されています。上書きしますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    Else
        j = lstSetting.ListCount
        lstSetting.AddItem ""
    End If
    
    lstSetting.List(j, C_SETLIST_NO) = j + 1
    lstSetting.List(j, C_SETLIST_ENABLE) = C_ENABLE
    lstSetting.List(j, C_SETLIST_KEY_NAME) = cmbShift.List(cmbShift.ListIndex, 0) & "+" & lstKey.List(lstKey.ListIndex, 1)
    lstSetting.List(j, C_SETLIST_KEY) = cmbShift.List(cmbShift.ListIndex, 1) & lstKey.List(lstKey.ListIndex, 2)
    lstSetting.List(j, C_SETLIST_CATEGORY) = lstCommand.List(lstCommand.ListIndex, 1)
    lstSetting.List(j, C_SETLIST_MACRO_NAME) = lstCommand.List(lstCommand.ListIndex, 2)
    lstSetting.List(j, C_SETLIST_MACRO) = lstCommand.List(lstCommand.ListIndex, 3)
    
    For i = 0 To lstSetting.ListCount - 1
        lstSetting.Selected(i) = False
    Next
    
    lstSetting.Selected(j) = True
    
    If lstSetting.ListCount > 0 Then
        cmdDel.enabled = True
    Else
        cmdDel.enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim i As Long
    Dim j As Long

    If lstSetting.ListCount > 0 Then
        i = lstSetting.ListIndex
        lstSetting.RemoveItem i
        If i > lstSetting.ListCount - 1 Then
            i = i - 1
            If i < 0 Then
                i = 0
            End If
        Else
            For j = i To lstSetting.ListCount - 1
                lstSetting.List(j, C_SETLIST_NO) = j + 1
            Next
        End If
        If lstSetting.ListCount > 0 Then
            lstSetting.ListIndex = i
            cmdDel.enabled = True
        Else
            cmdDel.enabled = False
        End If
    End If
End Sub

Private Sub cmdDisable_Click()

    If lstSetting.ListIndex >= 0 Then
        lstSetting.List(lstSetting.ListIndex, C_SETLIST_ENABLE) = C_DISABLE
    End If
    
End Sub

Private Sub cmdEneble_Click()

    If lstSetting.ListIndex >= 0 Then
        lstSetting.List(lstSetting.ListIndex, C_SETLIST_ENABLE) = C_ENABLE
    End If
    
End Sub


Private Sub cmdExport_Click()

    Dim strBuf As String
    Dim strLine As String
    Dim i As Long
    
    Dim vntFileName As Variant
    
    vntFileName = Application.GetSaveAsFilename(InitialFileName:="export.key", fileFilter:="キー定義(*.key),*.key", Title:="キー定義のエクスポート")
    
    If vntFileName = False Then
        Exit Sub
    End If
    
    If rlxIsFileExists(vntFileName) Then
        If MsgBox("すでにファイルが存在します。上書きしますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
    End If
    
    Dim fp As Integer
    Dim sw As Boolean
    
    On Error GoTo e
    
    fp = FreeFile()
    Open vntFileName For Output As fp
    sw = True
    
    Print #fp, "# RelaxTools Addin ショートカットキー定義"
    Dim strVer As String
    strVer = Split(ThisWorkbook.BuiltinDocumentProperties("Comments").Value, vbLf)(0)
    Print #fp, "# Export " & strVer
    Print #fp, "#"
    For i = 0 To lstSetting.ListCount - 1
        Print #fp, "# 【" & lstSetting.List(i, C_SETLIST_KEY_NAME) & "】" & lstSetting.List(i, C_SETLIST_MACRO_NAME)
    Next
    Print #fp, "#"
    Print #fp, "# Author:" & Application.UserName
    Print #fp, "# Date:" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    For i = 0 To lstSetting.ListCount - 1
        
        strLine = """" & lstSetting.List(i, C_SETLIST_NO) & """" & "," & """" & lstSetting.List(i, C_SETLIST_KEY_NAME) & """" & "," & """" & lstSetting.List(i, C_SETLIST_KEY) & """" & "," & """" & lstSetting.List(i, C_SETLIST_CATEGORY) & """" & "," & """" & lstSetting.List(i, C_SETLIST_MACRO_NAME) & """" & "," & """" & lstSetting.List(i, C_SETLIST_MACRO) & """"
        Print #fp, strLine
            
    Next

    sw = False
    Close fp
    
    MsgBox "エクスポートされました", vbOKOnly + vbInformation, C_TITLE
    Exit Sub
e:
    MsgBox "エクスポートでエラーが発生しました", vbOKOnly + vbCritical, C_TITLE
    If sw Then Close

End Sub

Private Sub cmdImport_Click()

    Dim strFile As Variant
    Dim strMsg As String

    On Error GoTo ErrHandle

    strFile = Application.GetOpenFilename("キー定義ファイル(*.key),*.key", , "キー定義のインポート", , False)
    If strFile = False Then
        'ファイル名が指定されなかった場合
        Exit Sub
    End If
    
    'ファイルの存在チェック
    If rlxIsFileExists(strFile) Then
    Else
        MsgBox "ファイルが存在しません。", vbExclamation, C_TITLE
        Exit Sub
    End If

'    If MsgBox("定義の取込を行います。現在の定義に追加しますか？" & vbCrLf & "重複するキー定義は上書きされます。" & vbCrLf & "「いいえ」を選択した場合、キー定義はクリアされます。", vbYesNo + vbQuestion, C_TITLE) = vbYes Then
'    Else
'        lstSetting.Clear
'    End If

    On Error GoTo ErrHandle
    
    Dim fp As Integer
    Dim strBuf As String

    fp = FreeFile()
    
    Open strFile For Input As fp
    
    Do Until Eof(fp)
        
        Line Input #fp, strBuf
        
        If Left(strBuf, 1) = "#" Then
            strMsg = strMsg & vbCrLf & Mid(strBuf, 2)
        End If
        
    Loop
    
    Close fp

    Select Case frmImportKey.Start(strMsg)
        Case vbYes
        Case vbNo
            lstSetting.Clear
        Case Else
            Exit Sub
    End Select

    Dim varField As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    
    fp = FreeFile()
    
    Open strFile For Input As fp
    
    l = 0
    i = lstSetting.ListCount
    Do Until Eof(fp)
pass:
        Line Input #fp, strBuf
        l = l + 1
        If Left(strBuf, 1) = "#" Then
            GoTo pass
        End If
        
        varField = rlxCsvPart(strBuf)
        
        If UBound(varField) = 6 Then
        
            k = existKey(varField)
            Select Case k
                Case -1
                    lstSetting.AddItem ""
                    lstSetting.List(i, C_SETLIST_NO) = i + 1
                    lstSetting.List(i, C_SETLIST_KEY_NAME) = varField(2)
                    lstSetting.List(i, C_SETLIST_KEY) = varField(3)
                    lstSetting.List(i, C_SETLIST_CATEGORY) = varField(4)
                    lstSetting.List(i, C_SETLIST_MACRO_NAME) = varField(5)
                    lstSetting.List(i, C_SETLIST_MACRO) = varField(6)
                    lstSetting.List(i, C_SETLIST_ENABLE) = C_ENABLE
                    i = i + 1
                Case -2
                    MsgBox "存在しないキー(" & varField(3) & ")です。行=" & l, vbOKOnly + vbExclamation, C_TITLE
                    Close fp
                    Exit Sub
                Case -3
                    MsgBox "存在しないマクロ名(" & varField(6) & ")です。行=" & l, vbOKOnly + vbExclamation, C_TITLE
                    Close fp
                    Exit Sub
                Case Else
                    lstSetting.List(k, C_SETLIST_KEY_NAME) = varField(2)
                    lstSetting.List(k, C_SETLIST_KEY) = varField(3)
                    lstSetting.List(k, C_SETLIST_CATEGORY) = varField(4)
                    lstSetting.List(k, C_SETLIST_MACRO_NAME) = varField(5)
                    lstSetting.List(k, C_SETLIST_MACRO) = varField(6)
                    lstSetting.List(k, C_SETLIST_ENABLE) = C_ENABLE
            End Select
        Else
            MsgBox "インポートファイルの形式が不正です。", vbOKOnly + vbExclamation, C_TITLE
            Close fp
            Exit Sub
        End If
    Loop
    
    Close fp
    
    If lstSetting.ListCount > 0 Then
        lstSetting.ListIndex = lstSetting.ListCount - 1
        cmdDel.enabled = True
    Else
        cmdDel.enabled = False
    End If
    
    MsgBox "正常にインポートされました。" & vbCrLf, vbOKOnly + vbInformation, C_TITLE

    Exit Sub
ErrHandle:
    MsgBox "インポートに失敗しました。", vbOKOnly + vbCritical, C_TITLE

End Sub
Private Function existMacro(ByVal strMacro As String) As Boolean

    Dim i As Long
    Dim WS As Worksheet
    
    existMacro = False
    
    Set WS = ThisWorkbook.Worksheets("HELP")
    i = C_COM_DATA

    'マクロシートのロード
    Do Until WS.Cells(i, C_COM_NO).Value = ""

        If WS.Cells(i, C_COM_USE).Value <> "－" Then
        
            If WS.Cells(i, C_COM_MACRO).Value = strMacro Then
                existMacro = True
                Exit Do
            End If
        
        End If
        i = i + 1
    Loop

End Function
Private Function existHotKey(ByVal strKey As String) As Boolean
    
    Dim i As Long
    Dim WS As Worksheet
    
    existHotKey = False
    
    'シフトキーを削除
    Dim lngPos As Long
    
    lngPos = InStr(strKey, "{^}")
    If lngPos > 0 Then
        strKey = "{^}"
    Else
        strKey = Replace(strKey, "^", "", 1, 1)
        strKey = Replace(strKey, "+", "", 1, 1)
        strKey = Replace(strKey, "%", "", 1, 1)
    End If
    
    Set WS = ThisWorkbook.Worksheets("key")
    i = C_KEY_DATA

    'マクロシートのロード
    Do Until WS.Cells(i, C_KEY_NO).Value = ""
        
        If WS.Cells(i, C_KEY_KEY).Value = strKey Then
            existHotKey = True
            Exit Do
        End If
        
        i = i + 1
    Loop

End Function
Private Function existKey(ByRef varBuf As Variant) As Long

    Dim i As Long
    
    existKey = -1
    
    For i = 0 To lstSetting.ListCount - 1
    
        'リストの重複チェック
        If lstSetting.List(i, C_SETLIST_KEY) = varBuf(3) Then
            existKey = i
        End If
        
        'キーの存在チェック
        If existHotKey(varBuf(3)) = False Then
            existKey = -2
            Exit Function
        End If
        
        'マクロの存在チェック
        If existMacro(varBuf(6)) = False Then
            existKey = -3
            Exit Function
        End If
    
    Next

End Function

Private Sub cmdSave_Click()
    
    Dim i As Long
    Dim strBuf As String
    Dim strLine As String
    
    Call removeShortCutKey
    
    strBuf = ""

    For i = 0 To lstSetting.ListCount - 1
        strLine = lstSetting.List(i, C_SETLIST_NO) & vbTab & lstSetting.List(i, C_SETLIST_KEY_NAME) & vbTab & lstSetting.List(i, C_SETLIST_KEY) & vbTab & lstSetting.List(i, C_SETLIST_CATEGORY) & vbTab & lstSetting.List(i, C_SETLIST_MACRO_NAME) & vbTab & lstSetting.List(i, C_SETLIST_MACRO) & vbTab & getFromEnable(lstSetting.List(i, C_SETLIST_ENABLE))
        If Len(strBuf) = 0 Then
            strBuf = strLine
        Else
            strBuf = strBuf & vbVerticalTab & strLine
        End If
    Next
    
    SaveSetting C_TITLE, "ShortCut", "KeyList", strBuf
    
    'ショートカットキーの登録
    Call setShortCutKey
    
    Unload Me

End Sub


Private Sub CommandButton1_Click()

End Sub

Private Sub lstCommand_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    MW.Name = "lstCommand"
    Set MW.obj = lstCommand
End Sub

Private Sub lstKey_Click()
    Call getGuidence
End Sub

Private Sub lstKey_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    MW.Name = "lstKey"
    Set MW.obj = lstKey

End Sub

Private Sub lstSetting_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim i As Long
    Dim strShift As String
    Dim strKey As String
    
    Dim strBuf As String
    If lstSetting.ListCount = 0 Then
        Exit Sub
    End If
    
    strBuf = lstSetting.List(lstSetting.ListIndex, C_SETLIST_KEY)
    
    For i = 1 To Len(strBuf)
        Select Case Mid$(strBuf, i, 1)
            Case "^", "%", "+"
            Case Else
                strShift = Mid$(strBuf, 1, i - 1)
                strKey = Mid$(strBuf, i)
                Exit For
        End Select
    Next
    
    For i = 0 To lstKey.ListCount - 1
        If lstKey.List(i, 2) = strKey Then
            lstKey.ListIndex = i
            lstKey.TopIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To cmbShift.ListCount - 1
        If cmbShift.List(i, 1) = strShift Then
            cmbShift.ListIndex = i
            Exit For
        End If
    Next

    strBuf = lstSetting.List(lstSetting.ListIndex, C_SETLIST_MACRO)
    
    For i = 0 To lstCommand.ListCount - 1
        If lstCommand.List(i, 3) = strBuf Then
            lstCommand.ListIndex = i
            lstCommand.TopIndex = i
            Exit For
        End If
    Next
    
End Sub


Private Sub lstSetting_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    MW.Name = "lstSetting"
    Set MW.obj = lstSetting

End Sub


Private Sub txtKinou_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call dispCommand
    End If
End Sub
'Private Sub UserForm_Activate()
'    MW.Activate
'End Sub
Private Sub UserForm_Initialize()
    
    Dim WS As Worksheet
    Dim i As Long
    Dim j As Long

    cmbShift.AddItem ""
    cmbShift.List(0, 0) = "シフトキーなし"
    cmbShift.List(0, 1) = ""
    cmbShift.AddItem ""
    cmbShift.List(1, 0) = "SHIFT"
    cmbShift.List(1, 1) = "+"
    cmbShift.AddItem ""
    cmbShift.List(2, 0) = "CTRL"
    cmbShift.List(2, 1) = "^"
    cmbShift.AddItem ""
    cmbShift.List(3, 0) = "CTRL+SHIFT"
    cmbShift.List(3, 1) = "^+"
    cmbShift.AddItem ""
    cmbShift.List(4, 0) = "CTRL+ALT"
    cmbShift.List(4, 1) = "^%"
    cmbShift.AddItem ""
    cmbShift.List(5, 0) = "CTRL+ALT+SHIFT"
    cmbShift.List(5, 1) = "^%+"

    cmbShift.ListIndex = 2


    Set WS = ThisWorkbook.Worksheets("HELP")
    i = C_COM_DATA
    j = 0

    Dim strBefore As String
    strBefore = ""
    
    cboCategory.AddItem "すべて"
    'マクロシートのロード
    Do Until WS.Cells(i, C_COM_NO).Value = ""

        If WS.Cells(i, C_COM_USE).Value <> "－" Then
            If WS.Cells(i, C_COM_CATEGORY).Value <> strBefore Then
                cboCategory.AddItem WS.Cells(i, C_COM_CATEGORY).Value
                strBefore = WS.Cells(i, C_COM_CATEGORY).Value
            End If
        End If
        i = i + 1

    Loop
    cboCategory.ListIndex = 0

    Call dispCommand

    
    Set WS = ThisWorkbook.Worksheets("key")
    i = C_KEY_DATA
    j = 0

    'マクロシートのロード
    Do Until WS.Cells(i, C_KEY_NO).Value = ""
        
        lstKey.AddItem ""
        lstKey.List(j, 0) = j + 1
        lstKey.List(j, 1) = WS.Cells(i, C_KEY_NAME).Value
        lstKey.List(j, 2) = WS.Cells(i, C_KEY_KEY).Value
        j = j + 1
        i = i + 1

    Loop
    lstKey.ListIndex = 0
    
    
    Dim strList() As String
    Dim strKey() As String
    Dim strResult As String
    Dim lngMax As Long
    
    strResult = GetSetting(C_TITLE, "ShortCut", "KeyList", "")
    strList = Split(strResult, vbVerticalTab)

    lngMax = UBound(strList)

    For i = 0 To lngMax
        strKey = Split(strList(i) & vbTab & "1", vbTab)
        
        lstSetting.AddItem ""
        lstSetting.List(i, C_SETLIST_NO) = i + 1
        lstSetting.List(i, C_SETLIST_KEY_NAME) = strKey(1)
        lstSetting.List(i, C_SETLIST_KEY) = strKey(2)
        lstSetting.List(i, C_SETLIST_CATEGORY) = strKey(3)
        lstSetting.List(i, C_SETLIST_MACRO_NAME) = strKey(4)
        lstSetting.List(i, C_SETLIST_MACRO) = strKey(5)
        lstSetting.List(i, C_SETLIST_ENABLE) = getEnable(strKey(6))
    Next
    
    If lstSetting.ListCount > 0 Then
        lstSetting.ListIndex = 0
        cmdDel.enabled = True
    Else
        cmdDel.enabled = False
    End If
        
    mstrhWnd = CStr(FindWindow("ThunderDFrame", Me.Caption))
    Set MW = basMouseWheel.Install(mstrhWnd)

'    MW.Install Me
    
End Sub
Function getEnable(ByVal strBuf As String) As String
    If strBuf = "1" Then
        getEnable = C_ENABLE
    Else
        getEnable = C_DISABLE
    End If
End Function
Function getFromEnable(ByVal strBuf As String) As String
    If strBuf = C_ENABLE Then
        getFromEnable = "1"
    Else
        getFromEnable = "0"
    End If
End Function
Sub dispCommand()

    Dim WS As Worksheet
    Dim i As Long
    Dim j As Long

    Set WS = ThisWorkbook.Worksheets("HELP")
    i = C_COM_DATA
    j = 0
    
    lstCommand.Clear

    'マクロシートのロード
    Do Until WS.Cells(i, C_COM_NO).Value = ""
        
        If WS.Cells(i, C_COM_USE).Value <> "－" Then
            If (cboCategory.ListIndex = 0 Or cboCategory.Text = WS.Cells(i, C_COM_CATEGORY).Value) And (txtKinou.Text = "" Or InStr(WS.Cells(i, C_COM_DISP_NAME).Value, txtKinou.Text) > 0) Then
                lstCommand.AddItem ""
                lstCommand.List(j, 0) = j + 1
                lstCommand.List(j, 1) = WS.Cells(i, C_COM_CATEGORY).Value
                lstCommand.List(j, 2) = WS.Cells(i, C_COM_DISP_NAME).Value
                lstCommand.List(j, 3) = WS.Cells(i, C_COM_MACRO).Value
                j = j + 1
            End If
        End If
        i = i + 1

    Loop
    If lstCommand.ListCount > 0 Then
        lstCommand.ListIndex = 0
        cmdAdd.enabled = True
    Else
        cmdAdd.enabled = False
    End If
End Sub
Sub getGuidence()

    Dim WS As Worksheet
    Dim i As Long

    Set WS = ThisWorkbook.Worksheets("ShortCut")
    i = 3
    
    txtGuidence.Text = ""

    'マクロシートのロード
    Do Until WS.Cells(i, 1).Value = ""
        
        If WS.Cells(i, 3).Value = cmbShift.List(cmbShift.ListIndex, 1) & lstKey.List(lstKey.ListIndex, 2) Then
            txtGuidence.Text = "【" & WS.Cells(i, 2).Value & "】" & WS.Cells(i, 4).Value
            Exit Do
        End If
        i = i + 1
    Loop
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    MW.Name = "lstSetting"
    Set MW.obj = Nothing
    
End Sub

Private Sub UserForm_Terminate()
    Set MW = basMouseWheel.UnInstall(mstrhWnd)
End Sub

'Private Sub MW_WheelDown(ByVal Name As String)
'
'    Select Case Name
'        Case "lstKey"
'            If lstKey.ListCount = 0 Then Exit Sub
'            lstKey.TopIndex = lstKey.TopIndex + 3
'        Case "lstCommand"
'            If lstCommand.ListCount = 0 Then Exit Sub
'            lstCommand.TopIndex = lstCommand.TopIndex + 3
'        Case "lstSetting"
'            If lstSetting.ListCount = 0 Then Exit Sub
'            lstSetting.TopIndex = lstSetting.TopIndex + 3
'    End Select
'
'End Sub
'
'Private Sub MW_WheelUp(ByVal Name As String)
'
'    Dim lngPos As Long
'
'    Select Case Name
'        Case "lstKey"
'            If lstKey.ListCount = 0 Then Exit Sub
'            lngPos = lstKey.TopIndex - 3
'        Case "lstCommand"
'            If lstCommand.ListCount = 0 Then Exit Sub
'            lngPos = lstCommand.TopIndex - 3
'        Case "lstSetting"
'            If lstSetting.ListCount = 0 Then Exit Sub
'            lngPos = lstSetting.TopIndex - 3
'    End Select
'
'    If lngPos < 0 Then
'        lngPos = 0
'    End If
'
'    Select Case Name
'        Case "lstKey"
'            lstKey.TopIndex = lngPos
'        Case "lstCommand"
'            lstCommand.TopIndex = lngPos
'        Case "lstSetting"
'            lstSetting.TopIndex = lngPos
'    End Select
'
'End Sub
'Private Sub MW_WheelDown(ByVal Name As String)
'
'    Select Case Name
'        Case "lstKey"
'            If lstKey.ListCount = 0 Then Exit Sub
'            lstKey.TopIndex = lstKey.TopIndex + 3
'        Case "lstCommand"
'            If lstCommand.ListCount = 0 Then Exit Sub
'            lstCommand.TopIndex = lstCommand.TopIndex + 3
'        Case "lstSetting"
'            If lstSetting.ListCount = 0 Then Exit Sub
'            lstSetting.TopIndex = lstSetting.TopIndex + 3
'    End Select
'
'End Sub

Private Sub MW_WheelDown(obj As Object)

    On Error GoTo e

    If obj.ListCount = 0 Then Exit Sub
    obj.TopIndex = obj.TopIndex + 3
e:
End Sub

Private Sub MW_WheelUp(obj As Object)

    On Error GoTo e

    Dim lngPos As Long

    If obj.ListCount = 0 Then Exit Sub
    lngPos = obj.TopIndex - 3

    If lngPos < 0 Then
        lngPos = 0
    End If

    obj.TopIndex = lngPos
e:
End Sub

