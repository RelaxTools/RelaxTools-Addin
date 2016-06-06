VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetPackage 
   Caption         =   "Javaパッケージ配置"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   OleObjectBlob   =   "frmSetPackage.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSetPackage"
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
Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFolder_Click()

    Dim strFile As String

    'フォルダ名取得
    strFile = rlxSelectFolder()
    
    If Trim(strFile) <> "" Then
        txtFolder.Text = strFile
    End If
    
End Sub

Private Sub cmdPackage_Click()

    Dim strFile As String

    'フォルダ名取得
    strFile = rlxSelectFolder()
    
    If Trim(strFile) <> "" Then
        txtPackage.Text = strFile
    End If
    
End Sub

Private Sub cmdRun_Click()

    If Len(Trim(txtFolder.Text)) = 0 Then
        MsgBox "Javaファイルのあるフォルダを入力してください。", vbExclamation, C_TITLE
        Exit Sub
    End If

    If Len(Trim(txtPackage.Text)) = 0 Then
        MsgBox "Javaファイルを配置するフォルダを入力してください。", vbExclamation, C_TITLE
        Exit Sub
    End If

    Call setJavaPackage

End Sub

Private Sub setJavaPackage()

    Dim strDir As String
    Dim fp As Integer
    
    Dim strLine() As String
    
    Dim bytBuf() As Byte
    Dim strBuf() As String
    Dim strPath() As String
    Dim strDest As String
    Dim strSource As String
    Dim strMkDir As String
    
    Dim intCnt As Integer
    Dim lngCount As Long
    
    Dim blnCopySuccess As Boolean
    Dim lngDeleteSuccess As Long
    
    Dim blnNoPackage As Boolean
    
    Const C_RESULT_NO As Long = 0
    Const C_RESULT_FILE As Long = 1
    Const C_RESULT_PACKAGE As Long = 2
    Const C_RESULT_COPY As Long = 3
    Const C_RESULT_DELETE As Long = 4
  
    Const C_DELETE_NONE As Long = 0
    Const C_DELETE_SUCCESS As Long = 1
    Const C_DELETE_FAIL As Long = 2
    
    Dim BASE_FOLDER As String
    Dim DEST_FOLDER As String
    
    BASE_FOLDER = rlxAddFileSeparator(txtFolder.Text)
    DEST_FOLDER = rlxAddFileSeparator(txtPackage.Text)
    
    Dim FS As Object
    Dim D As Object
    Dim f As Object
    
    strDir = Dir(BASE_FOLDER & "*.java")
    If strDir = "" Then
        MsgBox "処理対象のJavaファイルがありません。処理を終了します。", vbExclamation, C_TITLE
        Exit Sub
    End If
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set D = FS.GetFolder(BASE_FOLDER)

    '処理結果リストのクリア
    lstResult.Clear
    lngCount = 0

    'Do Until strDir = ""
    For Each f In D.files
        
        blnNoPackage = True
    
        strDir = f.Name
        
        If LCase(FS.GetExtensionName(f.Name)) <> "java" Then
            GoTo pass
        End If
            
        fp = FreeFile()
        Open BASE_FOLDER & strDir For Binary As fp
        
        '先頭2KBだけ先読みする。
        Const C_MAX_READ As Long = 2048
        Select Case LOF(fp)
            Case 0
                Close fp
                GoTo pass
            Case Is < C_MAX_READ
                ReDim bytBuf(0 To LOF(fp) - 1)
                Get fp, , bytBuf
            Case Else
                ReDim bytBuf(0 To C_MAX_READ - 1)
                Get fp, , bytBuf
        End Select
        Close fp
        
        Dim strAll As String
        Dim lngPos As Long
        Dim i As Long
        
        strAll = StrConv(bytBuf, vbUnicode)
        
        lngPos = InStr(strAll, vbCrLf)
        If lngPos <> 0 Then
            strLine = Split(strAll, vbCrLf)
        Else
            lngPos = InStr(strAll, vbLf)
            If lngPos <> 0 Then
                strLine = Split(strAll, vbLf)
            Else
                strLine = Split(strAll, vbCr)
            End If
        End If
        
        For i = LBound(strLine) To UBound(strLine)
        
        
'        fp = FreeFile()
'        Open BASE_FOLDER & strDir For Input As fp
'
'        Do Until EOF(fp)
'            Line Input #fp, strLine
            
            '「;」と前後スペースを削除してスペースを区切りとして分割
            strBuf = Split(Trim(Replace(strLine(i), ";", "")), " ")
            
            If UBound(strBuf) > 0 Then
                'パラグラフが「package」の場合
                If InStr(strBuf(0), "package") > 0 Then
                
'                    Close fp
                
                    strPath = Split(strBuf(1), ".")
                    strDest = Replace(strBuf(1), ".", "\")
                    
                    strMkDir = ""
                    For intCnt = LBound(strPath) To UBound(strPath)
                        If strMkDir = "" Then
                            strMkDir = strPath(intCnt)
                        Else
                            strMkDir = strMkDir & "\" & strPath(intCnt)
                        End If
                        On Error Resume Next
                        MkDir DEST_FOLDER & strMkDir
                    Next intCnt
                    
                    
                    On Error GoTo 0
                    Err.Clear
                    
                    strSource = BASE_FOLDER & strDir
                    strDest = DEST_FOLDER & strDest & "\" & strDir
                    
                    On Error Resume Next
                    blnCopySuccess = False
                    FileCopy strSource, strDest
                    If Err.Number = 0 Then
                        blnCopySuccess = True
                    End If
                    
                    
                    On Error GoTo 0
                    Err.Clear
                    
                    'チェックボックスが選択されている場合
                    lngDeleteSuccess = C_DELETE_NONE
                    If chkDelete.value Then
                        If Dir$(strDest) <> "" Then
                            '元ファイルを削除する。
                            On Error Resume Next
                            Kill strSource
                            If Err.Number = 0 Then
                                lngDeleteSuccess = C_DELETE_SUCCESS
                            Else
                                lngDeleteSuccess = C_DELETE_FAIL
                            End If
                            On Error GoTo 0
                            Err.Clear
                        End If
                    End If
                    
                    blnNoPackage = False
                    GoTo pass
                End If
            End If
        Next
'        Loop
'        Close fp
pass:
        lstResult.AddItem ""
        lstResult.List(lngCount, C_RESULT_NO) = lngCount + 1
        lstResult.List(lngCount, C_RESULT_FILE) = strDir
        
        Dim strPackageResult As String
        Dim strCopyResult As String
        Dim strDeleteResult As String
        
        If blnNoPackage Then
            strPackageResult = "－"
            strCopyResult = "－"
            strDeleteResult = "－"
        Else
            strPackageResult = "○"
            
            If blnCopySuccess Then
                strCopyResult = "○"
            Else
                strCopyResult = "×"
            End If
            
            Select Case lngDeleteSuccess
                Case C_DELETE_NONE
                    strDeleteResult = "－"
                Case C_DELETE_SUCCESS
                    strDeleteResult = "○"
                Case C_DELETE_FAIL
                    strDeleteResult = "×"
            End Select
            
        End If
        
        lstResult.List(lngCount, C_RESULT_PACKAGE) = strPackageResult
        lstResult.List(lngCount, C_RESULT_COPY) = strCopyResult
        lstResult.List(lngCount, C_RESULT_DELETE) = strDeleteResult
        
        lngCount = lngCount + 1
        
        
        'strDir = Dir
    'Loop
    Next

    MsgBox "配置が完了しました。", vbInformation, C_TITLE

End Sub

Private Sub lstResult_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Set MW.obj = lstResult
End Sub

Private Sub UserForm_Initialize()
    Set MW = basMouseWheel.GetInstance
    MW.Install
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
