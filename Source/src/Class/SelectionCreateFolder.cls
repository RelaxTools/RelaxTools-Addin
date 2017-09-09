VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionCreateFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 宣言
'------------------------------------------------------------------------------------------------------------------------
Private WithEvents SFWork As SelectionFrameWork
Attribute SFWork.VB_VarHelpID = -1

'------------------------------------------------------------------------------------------------------------------------
' メンバ変数宣言部(UOC)
'------------------------------------------------------------------------------------------------------------------------
Public Folder As String
Dim mlngCount As Long

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New SelectionFrameWork
End Sub

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 開放
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set SFWork = Nothing
End Sub

'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork Run
'------------------------------------------------------------------------------------------------------------------------
Public Sub Run()
    SFWork.Run
End Sub
'------------------------------------------------------------------------------------------------------------------------
' ここまでお約束の記述。
'------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------------------
' 前処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionInit(Cancel As Boolean, Undo As Boolean, Func As Boolean)

    '変数の初期化
    mlngCount = 0

    'フォルダ名取得
'    mstrFolder = rlxSelectFolder()
    If Folder = "" Then
        'キャンセル
        Cancel = True
    End If

End Sub

'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal NotHoldFormat As Boolean, Cancel As Boolean)
        
    Dim s As String
    Dim lngRet As Long
    
    On Error GoTo e
    
    s = r.Value
    
    lngRet = createFolder(rlxAddFileSeparator(Folder) & s)
    If lngRet <> 0 Then
        mlngCount = mlngCount + 1
        With r.Interior
            .Color = vbYellow
            .Pattern = xlSolid
        End With
    End If

    Exit Sub
e:
    Call rlxErrMsg(Err)
    Cancel = True

End Sub

'------------------------------------------------------------------------------------------------------------------------
' 終了処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionTerm()

    If mlngCount <> 0 Then
        MsgBox "フォルダを作成しましたが" & mlngCount & "個のフォルダ作成が出来ませんでした。", vbExclamation, C_TITLE
    Else
        MsgBox "フォルダを作成しました。", vbInformation, C_TITLE
    End If

End Sub
Private Function createFolder(ByVal strFolder As String) As Long

    Dim strFolders() As String
    Dim strBuf As String
    Dim f As Variant
    Dim lngMax As Long
    Dim lngCnt As Long
    Dim lngStart As Long
    
    strFolders = Split(strFolder, "\")
    
    strBuf = ""
    lngMax = UBound(strFolders)
    
    Select Case True
        Case Left(strFolder, 2) = "\\"
            lngStart = 3
            strBuf = "\\" & strFolders(2)
        Case Mid$(strFolder, 2, 1) = ":"
            lngStart = 1
            strBuf = strFolders(0)
    End Select
    
    For lngCnt = lngStart To lngMax
        
        strBuf = strBuf & "\" & strFolders(lngCnt)
        On Error Resume Next
        MkDir strBuf
        On Error GoTo 0
    
    Next

    If Not rlxIsFolderExists(strFolder) Then
        createFolder = -1
    Else
        createFolder = 0
    End If


End Function

