Attribute VB_Name = "basCommon"
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
Option Private Module

' 32-bit Function version.
' ドライブ名からネットワークドライブを取得
#If VBA7 And Win64 Then
    'VBA7 = Excel2010以降。赤くコンパイルエラーになって見えますが問題ありません。
    Declare PtrSafe Function WNetGetConnection32 Lib "MPR.DLL" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, lSize As Long) As Long
    Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
    Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As LongPtr, ByVal UINT As Long, ByVal lpszFile As String, ByVal ch As Long) As Long
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Declare PtrSafe Function FlashWindowEx Lib "user32.dll" (pfwi As FLASHWINFO) As LongPtr
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
    Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef lpPictDesc As PICTDESC, ByRef refiid As GUID, ByVal fPictureOwnsHandle As LongPtr, ByRef IPic As IPicture) As Long
    Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
    Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long

    Private Type ChooseColor
        lStructSize As LongPtr
        hWndOwner As LongPtr
        hInstance As LongPtr
        rgbResult As LongPtr
        lpCustColors As String
        Flags As LongPtr
        lCustData As LongPtr
        lpfnHook As LongPtr
        lpTemplateName As String
    End Type
    
    Private Type POINTAPI
        X As Long
        Y As Long
    End Type
    
    Private Type DROPFILES
        pFiles As Long
        pt As POINTAPI
        fNC As Long
        fWide As Long
    End Type
    
    Private Type FLASHWINFO
        cbsize As LongPtr
        hWnd As LongPtr
        dwFlags As Long
        uCount As Long
        dwTimeout As LongPtr
    End Type
    
    Private Type PICTDESC
        cbSizeofStruct As Long
        picType        As Long
        hImage         As LongPtr
        Option1        As LongPtr
        Option2        As LongPtr
    End Type
    Private Type GUID
        Data1          As Long
        Data2          As Integer
        Data3          As Integer
        Data4(7)       As Byte
    End Type
#Else
    Private Declare Function WNetGetConnection32 Lib "MPR.DLL" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, lSize As Long) As Long
    Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function CloseClipboard Lib "user32" () As Long
    Declare Function EmptyClipboard Lib "user32" () As Long
    Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
    Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpszFile As String, ByVal ch As Long) As Long
    Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function FlashWindowEx Lib "user32.dll" (pfwi As FLASHWINFO) As Long
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
    Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef lpPictDesc As PICTDESC, ByRef refiid As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
    Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
    Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
    

Private Type ChooseColor
      lStructSize As Long
      hWndOwner As Long
      hInstance As Long
      rgbResult As Long
      lpCustColors As String
      Flags As Long
      lCustData As Long
      lpfnHook As Long
      lpTemplateName As String
    End Type
    
    Private Type POINTAPI
        X As Long
        Y As Long
    End Type
    
    Private Type DROPFILES
        pFiles As Long
        pt As POINTAPI
        fNC As Long
        fWide As Long
    End Type
    
    Private Type FLASHWINFO
        cbsize As Long
        hWnd As Long
        dwFlags As Long
        uCount As Long
        dwTimeout As Long
    End Type
    
    Private Type PICTDESC
        cbSizeofStruct As Long
        picType        As Long
        hImage         As Long
        Option1        As Long
        Option2        As Long
    End Type
    Private Type GUID
        Data1          As Long
        Data2          As Integer
        Data3          As Integer
        Data4(7)       As Byte
    End Type
    
#End If


Private Const CF_BITMAP      As Long = 2
Private Const CF_PALETTE     As Long = 9

Private Const CC_RGBINIT = &H1                '色のデフォルト値を設定
Private Const CC_LFULLOPEN = &H2              '色の作成を行う部分を表示
Private Const CC_PREVENTFULLOPEN = &H4        '色の作成ボタンを無効にする
Private Const CC_SHOWHELP = &H8               'ヘルプボタンを表示

Private Const NO_ERROR As Long = 0
Private Const lBUFFER_SIZE As Long = 255
Private lpszRemoteName As String
Private cbRemoteName As Long

Public Const IMAGE_BITMAP As Long = 0
Public Const LR_COPYRETURNORG As Long = &H4

Public Const C_EXCEL_VERSION_2016 As Long = 16
Public Const C_EXCEL_VERSION_2013 As Long = 15
Public Const C_EXCEL_VERSION_2010 As Long = 14
Public Const C_EXCEL_VERSION_2007 As Long = 12
Public Const C_EXCEL_VERSION_2003 As Long = 11

'UNDOバッファ
Public Const C_TITLE As String = "RelaxTools-Addin"

Public Const C_MAX_CELLS As Long = 100000
Public pvarSelectionBuffer As Variant
Public pobjSelection As Object

Public Const C_UTF16 As String = "UTF-16(LE)"
Public Const C_UTF8 As String = "UTF-8"
Public Const C_SJIS As String = "MS932(ShiftJIS)"
Public Const C_SJIS_OLD As String = "Shift-JIS"
Public Const C_ERROR As String = "<<ERROR>>"
Public Const CF_TEXT As Long = 1  'テキストデータを読み書きする場合の定数です
Public Const CF_HDROP As Long = 15
Public Const C_REF_TEXT As String = "(参照用)"

Public Const C_ALL As Long = 3
Public Const C_HOLIZON As Long = 1
Public Const C_VERTICAL As Long = 2
'--------------------------------------------------------------
'　コントロールパネルのホイール量取得
'--------------------------------------------------------------
Function scrollPush() As Boolean
    Dim lngValue As Long
    
    Const GetWheelScrollLines = 104
    SystemParametersInfo GetWheelScrollLines, 0, lngValue, 0
    
    scrollPush = (lngValue = GetSetting(C_TITLE, "ScrollLine", "ScrollLine", 1))

End Function
'--------------------------------------------------------------
'　コントロールパネルのホイール量１行
'--------------------------------------------------------------
Sub scrollLine1()

    Const SENDCHANGE = 3
    Const SetWheelScrollLines = 105
    Dim lngValue As Long
    
    lngValue = GetSetting(C_TITLE, "ScrollLine", "ScrollLine", 1)
    SystemParametersInfo SetWheelScrollLines, lngValue, 0, SENDCHANGE

End Sub
'--------------------------------------------------------------
'　コントロールパネルのホイール量３行
'--------------------------------------------------------------
Sub scrollLine3()

    Const SENDCHANGE = 3
    Const SetWheelScrollLines = 105
    Dim lngValue As Long
    
    lngValue = GetSetting(C_TITLE, "ScrollLine", "DefaultLine", 3)
    SystemParametersInfo SetWheelScrollLines, lngValue, 0, SENDCHANGE

End Sub
'--------------------------------------------------------------
'　色を１６進文字列に変換
'--------------------------------------------------------------
Public Function getHexColor(ByVal lngColor As Long) As String
    getHexColor = "&H" & Right$("00000000" & Hex(lngColor), 8)
End Function
'--------------------------------------------------------------
'　１６進文字列を色に変換
'--------------------------------------------------------------
Public Function getLongColor(ByVal strColor As String) As Long
    On Error Resume Next
    getLongColor = CLng(strColor)
End Function
'--------------------------------------------------------------
'　アドレス文字列からオブジェクトに変換
'--------------------------------------------------------------
Public Function getObjectFromAddres(ByVal strAddress As String) As Object

    Dim obj As Object

    #If VBA7 And Win64 Then
        Dim p As LongPtr
        p = CLngPtr(strAddress)
    #Else
        Dim p As Long
        p = CLng(strAddress)
    #End If
  
    CopyMemory obj, p, LenB(p)
    
    Set getObjectFromAddres = obj

End Function
'--------------------------------------------------------------
'　ファイル数カウント
'--------------------------------------------------------------
Public Sub rlxGetFilesCount(ByRef objFs As Object, ByVal strPath As String, ByRef lngFCnt As Long, ByVal blnFile As Boolean, ByVal blnFolder As Boolean, ByVal blnSubFolder As Boolean)

    Dim objfld As Object
    Dim objSub As Object

    Set objfld = objFs.GetFolder(strPath)
    
    If blnFile Then
        lngFCnt = lngFCnt + objfld.files.Count
    End If
    
    If blnFolder Then
        lngFCnt = lngFCnt + objfld.SubFolders.Count
    End If
    
        'フォルダ取得あり
    If blnSubFolder Then
        For Each objSub In objfld.SubFolders
            DoEvents
            rlxGetFilesCount objFs, objSub.Path, lngFCnt, blnFile, blnFolder, blnSubFolder
        Next
        
    End If
End Sub
'--------------------------------------------------------------
'　ファイルセパレータ付加
'--------------------------------------------------------------
Public Function rlxAddFileSeparator(ByVal strFile As String) As String
Attribute rlxAddFileSeparator.VB_Description = "パスとファイルを連結する際に区切り文字(""\\"")を補完します。"
Attribute rlxAddFileSeparator.VB_ProcData.VB_Invoke_Func = " \n19"
    If Right(strFile, 1) = "\" Then
        rlxAddFileSeparator = strFile
    Else
        rlxAddFileSeparator = strFile & "\"
    End If
End Function
'--------------------------------------------------------------
'　フォルダ選択
'--------------------------------------------------------------
Public Function rlxSelectFolder() As String
Attribute rlxSelectFolder.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxSelectFolder.VB_ProcData.VB_Invoke_Func = " \n19"
 
    Dim objShell As Object
    Dim objPath As Object
    Dim WS As Object
    Dim strFolder As String
    
    Set objShell = CreateObject("Shell.Application")
    Set objPath = objShell.BrowseForFolder(&O0, "フォルダを選んでください", &H1 + &H10, "")
    If Not objPath Is Nothing Then
    
        'なぜか「デスクトップ」のパスが取得できない
        If objPath = "デスクトップ" Then
            Set WS = CreateObject("WScript.Shell")
            rlxSelectFolder = WS.SpecialFolders("Desktop")
        Else
            rlxSelectFolder = objPath.Items.Item.Path
        
        End If
    Else
        rlxSelectFolder = ""
    End If
    
End Function
'--------------------------------------------------------------
'　ファイル名取得
'--------------------------------------------------------------
Public Function rlxGetFullpathFromFileName(ByVal strPath As String) As String
Attribute rlxGetFullpathFromFileName.VB_Description = "パス＋ファイル情報からファイル名のみ返却します。"
Attribute rlxGetFullpathFromFileName.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngCnt As Long
    Dim lngMax As Long
    Dim strResult As String
    
    strResult = strPath
    
    lngMax = Len(strPath)
    
    For lngCnt = lngMax To 1 Step -1
    
        Select Case Mid$(strPath, lngCnt, 1)
            Case "\", "/"
                If lngCnt = lngMax Then
                Else
                    strResult = Mid$(strPath, lngCnt + 1)
                End If
                Exit For
        End Select
    
    Next

    rlxGetFullpathFromFileName = strResult

End Function
'--------------------------------------------------------------
'　ファイル名取得(拡張子抜き)
'--------------------------------------------------------------
Public Function rlxGetFullpathFromExt(ByVal strPath As String) As String
Attribute rlxGetFullpathFromExt.VB_Description = "パス＋ファイル情報から拡張子のみ返却します。"
Attribute rlxGetFullpathFromExt.VB_ProcData.VB_Invoke_Func = " \n19"

   Dim lngCnt As Long
    Dim lngMax As Long
    Dim strResult As String
    
    strResult = strPath
    
    lngMax = Len(strPath)
    
    For lngCnt = lngMax To 1 Step -1
    
        If Mid$(strPath, lngCnt, 1) = "." Then
            If lngCnt > 1 Then
                strResult = Mid$(strPath, 1, lngCnt - 1)
                Exit For
            End If
        End If
    
    Next

    rlxGetFullpathFromExt = strResult

End Function
'--------------------------------------------------------------
'　パス情報取得
'--------------------------------------------------------------
Public Function rlxGetFullpathFromPathName(ByVal strPath As String) As String
Attribute rlxGetFullpathFromPathName.VB_Description = "パス＋ファイル情報からディレクトリ情報のみ返却します。"
Attribute rlxGetFullpathFromPathName.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngCnt As Long
    Dim lngMax As Long
    Dim strResult As String
    
    strResult = strPath
    
    lngMax = Len(strPath)
    
    For lngCnt = lngMax To 1 Step -1
    
        Select Case Mid$(strPath, lngCnt, 1)
            Case "\", "/"
                If lngCnt > 1 Then
                    strResult = Mid$(strPath, 1, lngCnt - 1)
                    Exit For
                End If
        End Select
    
    Next

    rlxGetFullpathFromPathName = strResult

End Function
'--------------------------------------------------------------
'　DOSコマンド実行
'--------------------------------------------------------------
Function rlxShellExec(ByVal strCommand As String) As String
Attribute rlxShellExec.VB_Description = "DOSコマンドを実行し、標準出力を返却します。"
Attribute rlxShellExec.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim WSH As Object
    Dim wExec As Object
    
    Set WSH = CreateObject("WScript.Shell")
    
    Set wExec = WSH.exec("%ComSpec% /c " & strCommand)
    Do While wExec.Status = 0
        DoEvents
    Loop
    
    rlxShellExec = wExec.StdOut.ReadAll
    
    Set wExec = Nothing
    Set WSH = Nothing

End Function
'--------------------------------------------------------------
'　アンダーバーがあったらDB項目（大雑把）
'--------------------------------------------------------------
Public Function rlxIsDBField(ByVal strBuf As String) As Boolean
Attribute rlxIsDBField.VB_Description = "DB項目名（半角大文字＋アンダーバー）の場合\ntrueを返却します。"
Attribute rlxIsDBField.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim i As Long
    Dim lngCnt As Long
    
    rlxIsDBField = False
    
    If InStr(strBuf, "_") > 0 Then
        rlxIsDBField = True
        Exit Function
    End If

    lngLen = Len(strBuf)

    For i = 1 To lngLen

        Select Case Mid$(strBuf, i, 1)
            Case "a" To "z"
            Case Else
                lngCnt = lngCnt + 1
        End Select
    Next

    If lngLen = lngCnt Then
        rlxIsDBField = True
    End If


End Function
'--------------------------------------------------------------
'　Javaフィールド名→DB項目名
'--------------------------------------------------------------
Public Function rlxToDBFieldNm(ByVal strJavaField As String) As String
Attribute rlxToDBFieldNm.VB_Description = "Java項目名をDB項目名に変換します。\n ex. dbFieldName → DB_FIELD_NAME"
Attribute rlxToDBFieldNm.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim s As String
  
    Dim i As Long
    Dim Max As Long
    Dim u As Boolean
    Dim strBuf As String
    Dim strResult As String
    Dim strSrc As String
  
    u = False
    
    strSrc = strJavaField
    
    'すでにDB項目なら処理しない
    If rlxIsDBField(strSrc) Then
        rlxToDBFieldNm = strSrc
        Exit Function
    End If
    
    If Len(strSrc) >= 3 Then
        Select Case UCase(Mid$(strSrc, 1, 3))
            Case "GET", "SET"
                strSrc = Mid$(strSrc, 4)
        End Select
    End If
    Max = Len(strSrc)
    strResult = ""

    For i = 1 To Max
    
        strBuf = Mid$(strSrc, i, 1)
        Select Case strBuf
            Case "A" To "Z"
            u = True
        End Select
        
        If u Then
            If strResult <> "" Then
                strBuf = "_" & strBuf
            End If
            u = False
        Else
            strBuf = UCase(strBuf)
        End If
        strResult = strResult & strBuf
    
    Next
    rlxToDBFieldNm = strResult

End Function
'--------------------------------------------------------------
'　DB項目名→Javaフィールド名
'--------------------------------------------------------------
Public Function rlxToJavaFieldNm(ByVal strDBFieldName As String) As String
Attribute rlxToJavaFieldNm.VB_Description = "DB項目名をJava項目名に変換します。\n ex. DB_FIELD_NAME → dbFieldName"
Attribute rlxToJavaFieldNm.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim s As String
  
    Dim i As Long
    Dim Max As Long
    Dim u As Boolean
    Dim strBuf As String
    Dim strResult As String
    Dim strSrc As String
        
    u = False
    
    strSrc = strDBFieldName
    
    If Len(strSrc) >= 3 Then
        Select Case UCase(Mid$(strSrc, 1, 3))
            Case "GET", "SET"
                Select Case Len(strSrc)
                    Case 4
                        strSrc = LCase(Mid$(strSrc, 4, 1))
                    Case Is >= 5
                        strSrc = LCase(Mid$(strSrc, 4, 1)) & Mid$(strSrc, 5)
                End Select
        End Select
    End If
    
    Max = Len(strSrc)
    strResult = ""

    If rlxIsDBField(strSrc) Then
        For i = 1 To Max
        
            strBuf = Mid$(strSrc, i, 1)
            If strBuf = "_" Then
                u = True
            Else
            
                If u Then
                    strBuf = UCase(strBuf)
                    u = False
                Else
                    strBuf = LCase(strBuf)
                End If
                strResult = strResult & strBuf
            End If
        Next
    Else
        strResult = strSrc
    End If
    
    rlxToJavaFieldNm = strResult

End Function
'--------------------------------------------------------------
'　文字列のバイト数を求める。漢字２バイト、半角１バイト。
'--------------------------------------------------------------
Public Function rlxAscLen(ByVal var As Variant) As Long
Attribute rlxAscLen.VB_Description = "文字列のバイト数を求めます。漢字は２バイト、\n半角文字は１バイトと数えます。"
Attribute rlxAscLen.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim ascVar As Variant
    
    ascVar = StrConv(var, vbFromUnicode)


    rlxAscLen = LenB(ascVar)

End Function
'----------------------------------------------------------------------------------
'　文字列の左端から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
'----------------------------------------------------------------------------------
Public Function rlxAscLeft(ByVal var As Variant, ByVal lngsize As Long) As String
Attribute rlxAscLeft.VB_Description = "文字列の左端から指定した文字数分の文字列を返します。\n漢字２バイト、半角１バイト。"
Attribute rlxAscLeft.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim i As Long
    
    Dim strChr As String
    Dim strResult As String
    
    lngLen = Len(var)
    strResult = ""

    For i = 1 To lngLen
    
        strChr = Mid(var, i, 1)
        If rlxAscLen(strResult & strChr) > lngsize Then
            Exit For
        End If
        strResult = strResult & strChr
    
    Next

    rlxAscLeft = strResult

End Function
'----------------------------------------------------------------------------------
'　文字列の右端から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
'----------------------------------------------------------------------------------
Public Function rlxAscRight(ByVal var As Variant, ByVal lngsize As Long) As String
Attribute rlxAscRight.VB_Description = "文字列の右端から指定した文字数分の文字列を返します。\n漢字２バイト、半角１バイト。"
Attribute rlxAscRight.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim ascVar As Variant
    
    ascVar = StrConv(var, vbFromUnicode)

    rlxAscRight = StrConv(RightB(ascVar, lngsize), vbUnicode)

End Function
'----------------------------------------------------------------------------------
'　文字列から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
'----------------------------------------------------------------------------------
Public Function rlxAscMid(ByVal var As Variant, ByVal lngPos As Long, Optional ByVal varSize As Variant) As String
Attribute rlxAscMid.VB_Description = "文字列から指定した文字数分の文字列を返します。\n漢字２バイト、半角１バイト。"
Attribute rlxAscMid.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim ascVar As Variant
    
    ascVar = StrConv(var, vbFromUnicode)

    If IsMissing(varSize) Then
        rlxAscMid = StrConv(MidB(ascVar, lngPos), vbUnicode)
    Else
        rlxAscMid = StrConv(MidB(ascVar, lngPos, varSize), vbUnicode)
    End If

End Function
'--------------------------------------------------------------
'　ドライブ名→UNC名変換
'　ドライブ名(J:等)を指定。エラーの場合ドライブ名をそのまま返却
'--------------------------------------------------------------
Public Function rlxDriveToUNC(ByVal strPath As String) As String
Attribute rlxDriveToUNC.VB_Description = "ネットワークドライブをUNCに変換します。"
Attribute rlxDriveToUNC.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lStatus As Long
    Dim strDrive As String
    
    'デフォルトでパスをセット
    rlxDriveToUNC = strPath
    
    If InStr(strPath, ":") = 2 Then
        strDrive = Left$(strPath, 2)
    Else
        'ドライブ情報が含まれない。
        Exit Function
    End If

    cbRemoteName = lBUFFER_SIZE
    
    lpszRemoteName = lpszRemoteName & Space(lBUFFER_SIZE)
    
    lStatus& = WNetGetConnection32(strDrive, lpszRemoteName, cbRemoteName)
    
    If lStatus& = NO_ERROR Then
        rlxDriveToUNC = Left$(lpszRemoteName, InStr(lpszRemoteName, Chr$(0)) - 1) & Mid$(strPath, 3)
    Else
        'ドライブ情報が含まれるがネットワークドライブではない可能性。
        rlxDriveToUNC = strPath
    End If

End Function
'--------------------------------------------------------------
'　ファイル存在チェック
'--------------------------------------------------------------
Public Function rlxIsFileExists(ByVal strFile As String) As Boolean
Attribute rlxIsFileExists.VB_Description = "ファイルが存在する場合trueを返却します。"
Attribute rlxIsFileExists.VB_ProcData.VB_Invoke_Func = " \n19"
 
    With CreateObject("Scripting.FileSystemObject")
        rlxIsFileExists = .FileExists(strFile)
    End With

End Function
'--------------------------------------------------------------
'　フォルダ存在チェック
'--------------------------------------------------------------
Public Function rlxIsFolderExists(ByVal strFile As String) As Boolean
Attribute rlxIsFolderExists.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxIsFolderExists.VB_ProcData.VB_Invoke_Func = " \n19"
 
    With CreateObject("Scripting.FileSystemObject")
        rlxIsFolderExists = .FolderExists(strFile)
    End With

End Function
'--------------------------------------------------------------
'　テンポラリフォルダ取得
'--------------------------------------------------------------
Public Function rlxGetTempFolder() As String

    On Error Resume Next
    
    Dim strFolder As String
    
    rlxGetTempFolder = ""
    
    With CreateObject("Scripting.FileSystemObject")
    
        strFolder = rlxGetAppDataFolder & "Temp"
        
        If .FolderExists(strFolder) Then
        Else
            .createFolder strFolder
        End If
        
        rlxGetTempFolder = .BuildPath(strFolder, "\")
        
    End With
    

End Function
'--------------------------------------------------------------
'　アプリケーションフォルダ取得
'--------------------------------------------------------------
Public Function rlxGetAppDataFolder() As String

    On Error Resume Next
    
    Dim strFolder As String
    
    rlxGetAppDataFolder = ""
    
    With CreateObject("Scripting.FileSystemObject")
    
        strFolder = .BuildPath(CreateObject("Wscript.Shell").SpecialFolders("AppData"), C_TITLE)
        
        If .FolderExists(strFolder) Then
        Else
            .createFolder strFolder
        End If
        
        rlxGetAppDataFolder = .BuildPath(strFolder, "\")
        
    End With
    

End Function
'--------------------------------------------------------------
'　指定桁での四捨五入(decimal型非対称)
'--------------------------------------------------------------
Public Function rlxRound(ByVal varNumber As Variant, ByVal lngDigit As Long) As Variant

    rlxRound = Int(CDec(varNumber) * (10 ^ lngDigit) + CDec(0.5)) / 10 ^ lngDigit

End Function
'--------------------------------------------------------------
'　指定桁での切捨て(decimal型非対称)
'--------------------------------------------------------------
Public Function rlxRoundDown(ByVal varNumber As Variant, ByVal lngDigit As Long) As Variant

    rlxRoundDown = Int(CDec(varNumber) * (10 ^ lngDigit)) / 10 ^ lngDigit

End Function
'--------------------------------------------------------------
'　指定桁での切上げ(decimal型非対称)
'--------------------------------------------------------------
Public Function rlxRoundUp(ByVal varNumber As Variant, ByVal lngDigit As Long) As Variant

    Dim work As Variant
    Dim work2 As Variant

    work = Int(CDec(varNumber) * (10 ^ lngDigit))
    work2 = CDec(varNumber) * (10 ^ lngDigit)
    
    '小数点以下が存在する場合
    If work = work2 Then
    Else
        work = work + 1
    End If
    
    rlxRoundUp = work / 10 ^ lngDigit

End Function
'--------------------------------------------------------------
'　Luhnアルゴリズム（ISO/IEC 7812-1）
'　クレジットカード番号のチェック
'--------------------------------------------------------------
Function rlxIsLuhn(ByVal strNo As String) As Boolean
Attribute rlxIsLuhn.VB_Description = "Luhnアルゴリズム(クレジットカード番号など）の\nチェックディジットが正しい場合trueを返却します。"
Attribute rlxIsLuhn.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim lngOdd As Long
    Dim lngEvn As Long
    
    Dim i As Long
    
    Dim lngAns As Long
    
    Dim strCheckDigit As String
    
    lngLen = Len(strNo)
    lngOdd = 0
    lngEvn = 0

    If lngLen < 2 Then
        rlxIsLuhn = False
        Exit Function
    End If
    
    If rlxIsNumber(strNo) Then
    Else
        rlxIsLuhn = False
        Exit Function
    End If
    
    For i = 1 To lngLen
    
        If (i Mod 2) = 1 Then
            '奇数桁のみを加算（チェックディジットを除く）
            lngOdd = lngOdd + Val(Mid$(strNo, lngLen - i + 1, 1))
        Else
            '偶数桁のみを加算
            Dim lngWork As Long
            lngWork = Val(Mid$(strNo, lngLen - i + 1, 1)) * 2
            lngEvn = lngEvn + Fix(lngWork / 10) + lngWork Mod 10
        End If
    
    Next

    lngAns = (lngOdd + lngEvn) Mod 10

    If lngAns = 0 Then
        rlxIsLuhn = True
    Else
        rlxIsLuhn = False
    End If

End Function
'--------------------------------------------------------------
'　マイナンバーチェックデジット（個人）
'--------------------------------------------------------------
Function rlxIsMyNumber(ByVal strNo As String) As Boolean

 'マイナンバーチェックデジットチェック
    Dim strBuf As String
    Dim i As Long
    Dim c As Long
    Dim sum As Long
    Dim ans As Long
    Dim cd As Long
    
    rlxIsMyNumber = False
    
    If rlxIsNumber(strNo) Then
    Else
        Exit Function
    End If
    
    If Len(strNo) <> 12 Then
        Exit Function
    End If
    
    sum = 0

    For i = 0 To 11
    
        c = Val(Mid$(strNo, 11 - i + 1, 1))
        
        Select Case i
            Case 1 To 6
                sum = sum + c * (i + 1)
            Case 7 To 11
                sum = sum + c * (i - 5)
            Case 0
                cd = c
        End Select
    
    Next
    
    sum = sum Mod 11
    
    Select Case sum
        Case 0, 1
            ans = 0
        Case Else
            ans = 11 - sum
    End Select

    rlxIsMyNumber = (ans = cd)
    
End Function
'--------------------------------------------------------------
'　マイナンバーチェックデジット(企業)
'--------------------------------------------------------------
Function rlxIsCorpNumber(ByVal strNo As String) As Boolean

    '法人番号チェックデジットチェック
    Dim strBuf As String
    Dim i As Long
    Dim c As Long
    Dim sum As Long
    Dim ans As Long
    Dim cd As Long
    
    rlxIsCorpNumber = False
    
    If rlxIsNumber(strNo) Then
    Else
        Exit Function
    End If
    
    If Len(strNo) <> 13 Then
        Exit Function
    End If
    
    sum = 0

    For i = 1 To 13
    
        c = Val(Mid$(strNo, 13 - i + 1, 1))
        
        Select Case i
            Case 1 To 12
                sum = sum + c * IIf(i Mod 2, 1, 2)
            Case 13
                cd = c
        End Select
    
    Next
    
    sum = sum Mod 9
    
    ans = 9 - sum

    rlxIsCorpNumber = (ans = cd)

    
End Function
'--------------------------------------------------------------
'　モジュラス１０/ウェイト3-1
'--------------------------------------------------------------
Function rlxIsModulus10(ByVal strNo As String) As Boolean
Attribute rlxIsModulus10.VB_Description = "モジュラス10ウェイト3-1/JAN/EAN/ISBN13の\nチェックディジットが正しい場合trueを返却します。"
Attribute rlxIsModulus10.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim lngOdd As Long
    Dim lngEvn As Long
    
    Dim i As Long
    
    Dim lngAns As Long
    
    Dim lngCheckDigit As Long
    
    lngLen = Len(strNo)
    lngOdd = 0
    lngEvn = 0
    
    If lngLen < 2 Then
        rlxIsModulus10 = False
        Exit Function
    End If
    
    If rlxIsNumber(strNo) Then
    Else
        rlxIsModulus10 = False
        Exit Function
    End If

    For i = 1 To lngLen
    
        If i = 1 Then
            lngCheckDigit = Val(Mid$(strNo, lngLen - i + 1, 1))
        Else
            If (i Mod 2) = 1 Then
                '奇数桁のみを加算（チェックディジットを除く）
                lngOdd = lngOdd + Val(Mid$(strNo, lngLen - i + 1, 1))
            Else
                '偶数桁のみを加算
                lngEvn = lngEvn + Val(Mid$(strNo, lngLen - i + 1, 1))
            End If
        End If
    Next

    '奇数の加算と偶数の加算を３倍したものを加算。下１桁を１０から引く
    lngAns = 10 - (lngOdd + lngEvn * 3) Mod 10

    If lngAns = lngCheckDigit Then
        rlxIsModulus10 = True
    Else
        rlxIsModulus10 = False
    End If

End Function
'--------------------------------------------------------------
'　モジュラス１１ウェイト10-2
'--------------------------------------------------------------
Function rlxIsModulus11_10_2(ByVal strNo As String) As Boolean
Attribute rlxIsModulus11_10_2.VB_Description = "モジュラス11ウェイト10-2/ISBN10の\nチェックディジットが正しい場合trueを返却します。"
Attribute rlxIsModulus11_10_2.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim lngWork As Long
    Dim lngWeight As Long
    
    Dim i As Long
    
    Dim lngAns As Long
    
    Dim lngCheckDigit As Long
    
    lngLen = Len(strNo)
    lngWork = 0
    
    If lngLen < 2 Then
        rlxIsModulus11_10_2 = False
        Exit Function
    End If

    For i = 1 To lngLen
    
        If i = 1 Then
            lngCheckDigit = xVal(Mid$(strNo, lngLen - i + 1, 1))
        Else
            Select Case (i Mod 9)
                Case 2
                    lngWeight = 2
                Case 3
                    lngWeight = 3
                Case 4
                    lngWeight = 4
                Case 5
                    lngWeight = 5
                Case 6
                    lngWeight = 6
                Case 7
                    lngWeight = 7
                Case 8
                    lngWeight = 8
                Case 0
                    lngWeight = 9
                Case 1
                    lngWeight = 10
            End Select
            lngWork = lngWork + (Val(Mid$(strNo, lngLen - i + 1, 1)) * i)
        End If
    Next

    lngAns = (11 - (lngWork Mod 11)) Mod 11


    If lngAns = lngCheckDigit Then
        rlxIsModulus11_10_2 = True
    Else
        rlxIsModulus11_10_2 = False
    End If

End Function
'--------------------------------------------------------------
'　ISBNコードでチェックディジットがＸになった場合の変換。
'--------------------------------------------------------------
Private Function xVal(ByVal strNo) As Long
    If LCase(strNo) = "x" Then
        xVal = 10
    Else
        xVal = Val(strNo)
    End If
End Function
'--------------------------------------------------------------
'　モジュラス１１ウェイト2-7
'--------------------------------------------------------------
Function rlxIsModulus11_2_7(ByVal strNo As String) As Boolean
Attribute rlxIsModulus11_2_7.VB_Description = "モジュラス11/地方公共団体コードの\nチェックディジットが正しい場合trueを返却します。"
Attribute rlxIsModulus11_2_7.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim lngWork As Long
    Dim lngWeight As Long
    
    Dim i As Long
    
    Dim lngAns As Long
    
    Dim lngCheckDigit As Long
    
    lngLen = Len(strNo)
    lngWork = 0
    
    If lngLen < 2 Then
        rlxIsModulus11_2_7 = False
        Exit Function
    End If
    
    If rlxIsNumber(strNo) Then
    Else
        rlxIsModulus11_2_7 = False
        Exit Function
    End If

    For i = 1 To lngLen
    
        If i = 1 Then
            lngCheckDigit = Val(Mid$(strNo, lngLen - i + 1, 1))
        Else
            Select Case (i Mod 6)
                Case 2
                    lngWeight = 2
                Case 3
                    lngWeight = 3
                Case 4
                    lngWeight = 4
                Case 5
                    lngWeight = 5
                Case 0
                    lngWeight = 6
                Case 1
                    lngWeight = 7
            End Select
            lngWork = lngWork + (Val(Mid$(strNo, lngLen - i + 1, 1)) * lngWeight)
        End If
    Next

    lngAns = (11 - (lngWork Mod 11))

    If lngAns = lngCheckDigit Then
        rlxIsModulus11_2_7 = True
    Else
        rlxIsModulus11_2_7 = False
    End If

End Function
'--------------------------------------------------------------
'　モジュラス11/地方公共団体コード
'--------------------------------------------------------------
Function rlxIsModulus11_Pref(ByVal strNo As String) As Boolean
Attribute rlxIsModulus11_Pref.VB_Description = "モジュラス11ウェイト2-7/NW-7/免許証番号1～11の\nチェックディジットが正しい場合trueを返却します。"
Attribute rlxIsModulus11_Pref.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim lngWork As Long
    Dim lngMod As Long
    
    Dim i As Long
    
    Dim lngAns As Long
    
    Dim lngCheckDigit As Long
    
    lngLen = Len(strNo)
    lngWork = 0
    
    If lngLen < 2 Then
        rlxIsModulus11_Pref = False
        Exit Function
    End If
    
    If rlxIsNumber(strNo) Then
    Else
        rlxIsModulus11_Pref = False
        Exit Function
    End If
    
    For i = 1 To lngLen
    
        If i = 1 Then
            lngCheckDigit = Val(Mid$(strNo, lngLen - i + 1, 1))
        Else
            lngWork = lngWork + (Val(Mid$(strNo, lngLen - i + 1, 1)) * i)
        End If
    Next

    lngMod = lngWork Mod 11
    Select Case lngMod
        Case 0
            lngAns = 1
        Case 1
            lngAns = 0
        Case Else
            lngAns = 11 - lngMod
    End Select

    If lngAns = lngCheckDigit Then
        rlxIsModulus11_Pref = True
    Else
        rlxIsModulus11_Pref = False
    End If

End Function
'--------------------------------------------------------------
'　数字チェック
'--------------------------------------------------------------
Function rlxIsNumber(ByVal strNo As String) As Boolean
Attribute rlxIsNumber.VB_Description = "数字のみの場合trueを返却します。"
Attribute rlxIsNumber.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim i As Long
    
    rlxIsNumber = True
    
    lngLen = Len(strNo)
    
    For i = 1 To lngLen
    
        Select Case Mid(strNo, i, 1)
            Case "0" To "9"
            Case Else
                rlxIsNumber = False
                Exit Function
        End Select
    Next

End Function
'--------------------------------------------------------------
'　英字チェック
'--------------------------------------------------------------
Function rlxIsAlphabet(ByVal strNo As String) As Boolean
Attribute rlxIsAlphabet.VB_Description = "英字のみの場合trueを返却します。"
Attribute rlxIsAlphabet.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim i As Long
    
    rlxIsAlphabet = True
    
    lngLen = Len(strNo)
    
    For i = 1 To lngLen
    
        Select Case Mid(strNo, i, 1)
            Case "A" To "Z"
            Case "a" To "z"
            Case Else
                rlxIsAlphabet = False
                Exit Function
        End Select
    Next

End Function
'--------------------------------------------------------------
'　英数字チェック
'--------------------------------------------------------------
Function rlxIsAlphaAndNum(ByVal strNo As String) As Boolean
Attribute rlxIsAlphaAndNum.VB_Description = "英数字のみの場合trueを返却します。"
Attribute rlxIsAlphaAndNum.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngLen As Long
    Dim i As Long
    
    rlxIsAlphaAndNum = True
    
    lngLen = Len(strNo)
    
    For i = 1 To lngLen
    
        Select Case Mid(strNo, i, 1)
            Case "0" To "9"
            Case "A" To "Z"
            Case "a" To "z"
            Case Else
                rlxIsAlphaAndNum = False
                Exit Function
        End Select
    Next

End Function
'--------------------------------------------------------------
'  ＨＴＭＬ文字列のサニタイジングを行う。
'--------------------------------------------------------------
Public Function rlxHtmlSanitizing(ByVal strBuf As String) As String
Attribute rlxHtmlSanitizing.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxHtmlSanitizing.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim strRep As String

    strRep = Replace(strRep, "&", "&amp;")
    strRep = Replace(strBuf, """", "&quot;")
    strRep = Replace(strRep, "<", "&lt;")
    rlxHtmlSanitizing = Replace(strRep, ">", "&gt;")

End Function
'--------------------------------------------------------------
'  コレクションのソート
'--------------------------------------------------------------
Public Sub rlxSortCollection(ByRef col As Collection)

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim key1 As String
    Dim key2 As String
    Dim col2 As Collection
    Dim strKey() As String
    Dim wk As String

    'Collectionが空ならなにもしない
    If col Is Nothing Then
        Exit Sub
    End If

    'Collectionの要素数が０または１の場合ソート不要
    If col.Count <= 1 Then
        Exit Sub
    End If

    n = col.Count
    ReDim strKey(1 To n)

    For i = 1 To n
        strKey(i) = col.Item(i).Name
    Next

    '挿入ソート
    For i = 2 To n

        wk = strKey(i)

        If strKey(i - 1) > wk Then

            j = i

            Do

                strKey(j) = strKey(j - 1)

                j = j - 1

                If j = 1 Then
                    Exit Do
                End If

            Loop While strKey(j - 1) > wk
            strKey(j) = wk

        End If
    Next

    Set col2 = New Collection

    For i = 1 To n
        col2.Add col.Item(strKey(i)), col.Item(strKey(i)).Name
    Next

    Set col = col2
    Set col2 = Nothing

End Sub
''--------------------------------------------------------------
''  ディクショナリのソート
''--------------------------------------------------------------
'Public Sub rlxSortDictionary(ByRef col As Object)
'
'    Dim i As Long
'    Dim j As Long
'    Dim n As Long
'    Dim key1 As String
'    Dim key2 As String
'    Dim col2 As Object
'    Dim strKey() As String
'    Dim wk As String
'
'    'Collectionが空ならなにもしない
'    If col Is Nothing Then
'        Exit Sub
'    End If
'
'    'Collectionの要素数が０または１の場合ソート不要
'    If col.Count <= 1 Then
'        Exit Sub
'    End If
'
'    n = col.Count
'    ReDim strKey(1 To n)
'
'    i = i + 1
'    Dim v As Variant
'    For Each v In col
'        strKey(i) = v
'        i = i + 1
'    Next
'
'    '挿入ソート
'    For i = 2 To n
'
'        wk = strKey(i)
'
'        If strKey(i - 1) > wk Then
'
'            j = i
'
'            Do
'
'                strKey(j) = strKey(j - 1)
'
'                j = j - 1
'
'                If j = 1 Then
'                    Exit Do
'                End If
'
'            Loop While strKey(j - 1) > wk
'            strKey(j) = wk
'
'        End If
'    Next
'
'    Set col2 = CreateObject("Scripting.Dictionary")
'
'    For i = 1 To n
'        col2.Add col.Item(strKey(i)).Name, col.Item(strKey(i))
'    Next
'
'    Set col = col2
'    Set col2 = Nothing
'
'End Sub

'クリップボードにテキストデータを書き込むプロシージャ
Public Sub SetClipText(strData As String)

#If VBA7 And Win64 Then
  Dim lngHwnd As LongPtr, lngMEM As LongPtr
  Dim lngDataLen As LongPtr
  Dim lngRet As LongPtr
#Else
  Dim lngHwnd As Long, lngMEM As Long
  Dim lngDataLen As Long
  Dim lngRet As Long
#End If
  Dim blnErrflg As Boolean
  Const GMEM_MOVEABLE = 2

  blnErrflg = True
  
  'クリップボードをオープン
  If OpenClipboard(0&) <> 0 Then
  
    'クリップボードを空にする
    If EmptyClipboard() <> 0 Then
    
        'グローバルメモリに書き込む領域を確保してそのハンドルを取得
        lngDataLen = LenB(strData) + 1
        
        lngHwnd = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
        
        If lngHwnd <> 0 Then
      
            'グローバルメモリをロックしてそのポインタを取得
            lngMEM = GlobalLock(lngHwnd)
            
            If lngMEM <> 0 Then
        
                '書き込むテキストをグローバルメモリにコピー
                If lstrcpy(lngMEM, strData) <> 0 Then
                    'クリップボードにメモリブロックのデータを書き込み
                    lngRet = SetClipboardData(CF_TEXT, lngHwnd)
                    blnErrflg = False
                End If
                'グローバルメモリブロックのロックを解除
                lngRet = GlobalUnlock(lngHwnd)
            End If
        End If
    End If
    'クリップボードをクローズ(これはWindowsに制御が
    '戻らないうちにできる限り速やかに行う)
    lngRet = CloseClipboard()
  End If

  If blnErrflg Then MsgBox "クリップボードに情報が書き込めません", vbOKOnly, C_TITLE

End Sub
'クリップボードからテキストデータを取得するプロシージャ
Public Function GetClipText() As String

#If VBA7 And Win64 Then
  Dim lngHwnd As LongPtr, lngMEM As LongPtr
  Dim lngDataLen As LongPtr
  Dim lngRet As LongPtr
#Else
  Dim lngHwnd As Long, lngMEM As Long
  Dim lngDataLen As Long
  Dim lngRet As Long
#End If

    Const MAXSIZE = 4096
    Dim MyString As String
  
    'クリップボードをオープン
    If OpenClipboard(0&) <> 0 Then
    
        lngMEM = GetClipboardData(CF_TEXT)
        
        If lngMEM <> 0 Then
        
            lngHwnd = GlobalLock(lngMEM)
            
            If lngHwnd <> 0 Then
            
                MyString = Space$(MAXSIZE)
                
                lngRet = lstrcpy(MyString, lngHwnd)
                lngRet = GlobalUnlock(lngMEM)
                
                MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
            
            End If
        
        End If
        
        lngRet = CloseClipboard()
    
    End If
    
    GetClipText = MyString

End Function

'クリップボードにファイルコピー情報を書き込むプロシージャ
Public Sub SetCopyClipText(strBuf() As String)

#If VBA7 And Win64 Then
    Dim lngHwnd As LongPtr, lngMEM As LongPtr
    Dim lngDataLen As LongPtr
    Dim lngRet As LongPtr
#Else
    Dim lngHwnd As Long, lngMEM As Long
    Dim lngDataLen As Long
    Dim lngRet As Long
#End If

    Dim blnErrflg As Boolean
    Const GMEM_MOVEABLE = 2
    
    Dim df As DROPFILES
    
    Dim strData As String
    Dim i As Long
    
    For i = LBound(strBuf) To UBound(strBuf)
        strData = strData & strBuf(i) & vbNullChar
    Next
    strData = strData & vbNullChar

    blnErrflg = True
  
    'クリップボードをオープン
    If OpenClipboard(0&) <> 0 Then
  
        'クリップボードを空にする
        If EmptyClipboard() <> 0 Then
    
            'グローバルメモリに書き込む領域を確保してそのハンドルを取得
            lngDataLen = LenB(strData) + LenB(df) + 1024
            
            lngHwnd = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
            
            If lngHwnd <> 0 Then
            
                'グローバルメモリをロックしてそのポインタを取得
                lngMEM = GlobalLock(lngHwnd)
                
                If lngMEM <> 0 Then
                
                    df.pFiles = LenB(df)
            
                    '書き込むテキストをグローバルメモリにコピー
                    CopyMemory ByVal lngMEM, df, LenB(df)
                    CopyMemory ByVal (lngMEM + LenB(df)), ByVal strData, LenB(strData)
                    
                    'クリップボードにメモリブロックのデータを書き込み
                    lngRet = SetClipboardData(CF_HDROP, lngHwnd)
                    blnErrflg = False
                
                    'グローバルメモリブロックのロックを解除
                    lngRet = GlobalUnlock(lngHwnd)
                    
                End If
                
            End If
            
            
            'テキストも一緒に書き込んでおく
            strData = ""
            For i = LBound(strBuf) To UBound(strBuf)
                strData = strData & strBuf(i) & vbCrLf
            Next
            
            'グローバルメモリに書き込む領域を確保してそのハンドルを取得
            lngDataLen = LenB(strData) + 1
            
            lngHwnd = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
            
            If lngHwnd <> 0 Then
            
                'グローバルメモリをロックしてそのポインタを取得
                lngMEM = GlobalLock(lngHwnd)
                
                If lngMEM <> 0 Then
            
                    '書き込むテキストをグローバルメモリにコピー
                    If lstrcpy(lngMEM, strData) <> 0 Then
                        'クリップボードにメモリブロックのデータを書き込み
                        lngRet = SetClipboardData(CF_TEXT, lngHwnd)
                        blnErrflg = False
                    End If
                    'グローバルメモリブロックのロックを解除
                    lngRet = GlobalUnlock(lngHwnd)
                End If
            End If
            
        End If
        
        'クリップボードをクローズ(これはWindowsに制御が
        '戻らないうちにできる限り速やかに行う)
        lngRet = CloseClipboard()
    End If
    
    If blnErrflg Then MsgBox "クリップボードに情報が書き込めません", vbOKOnly, C_TITLE

End Sub
Function rlxSetLimit(ByVal l As Long, ByVal h As Long, ByVal lngVal As Long) As Long
Attribute rlxSetLimit.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxSetLimit.VB_ProcData.VB_Invoke_Func = " \n19"

    If lngVal < l Then
        lngVal = l
    End If
    
    If lngVal > h Then
        lngVal = h
    End If

    rlxSetLimit = lngVal

End Function
'--------------------------------------------------------------
'　インデント設定
'--------------------------------------------------------------
Sub setIndent(ByRef r As Range, ByVal lngIndent As Long)
    If lngIndent <> 0 Then
        If r.IndentLevel = 0 And lngIndent = -1 Then
        Else
'            r.InsertIndent lngIndent
            r.IndentLevel = r.IndentLevel + lngIndent
        End If
    End If
End Sub
'--------------------------------------------------------------
'　ローマ数字→数字
'--------------------------------------------------------------
Public Function rlxArabic(ByVal strRoman As String) As Long
Attribute rlxArabic.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxArabic.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngRet As Long

    Select Case LCase(strRoman)
        Case "i"
            lngRet = 1
        Case "ii"
            lngRet = 2
        Case "iii"
            lngRet = 3
        Case "iv"
            lngRet = 4
        Case "v"
            lngRet = 5
        Case "vi"
            lngRet = 6
        Case "vii"
            lngRet = 7
        Case "viii"
            lngRet = 8
        Case "ix"
            lngRet = 9
        Case "x"
            lngRet = 10
        Case "xi"
            lngRet = 11
        Case "xii"
            lngRet = 12
        Case "xiii"
            lngRet = 13
        Case "xiv"
            lngRet = 14
        Case "xv"
            lngRet = 15
        Case "xvi"
            lngRet = 16
        Case "xvii"
            lngRet = 17
        Case "xviii"
            lngRet = 18
        Case "xix"
            lngRet = 19
        Case "xx"
            lngRet = 20
    End Select

    rlxArabic = lngRet

End Function
'--------------------------------------------------------------
'　数字ローマ→数字
'--------------------------------------------------------------
Public Function rlxRoman(ByVal lngRoman As Long) As String
Attribute rlxRoman.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxRoman.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim strRet As String

    Select Case lngRoman
        Case 1
            strRet = "i"
        Case 2
            strRet = "ii"
        Case 3
            strRet = "iii"
        Case 4
            strRet = "iv"
        Case 5
            strRet = "v"
        Case 6
            strRet = "vi"
        Case 7
            strRet = "vii"
        Case 8
            strRet = "viii"
        Case 9
            strRet = "ix"
        Case 10
            strRet = "x"
        Case 11
            strRet = "xi"
        Case 12
            strRet = "xii"
        Case 13
            strRet = "xiii"
        Case 14
            strRet = "xiv"
        Case 15
            strRet = "xv"
        Case 16
            strRet = "xvi"
        Case 17
            strRet = "xvii"
        Case 18
            strRet = "xviii"
        Case 19
            strRet = "xix"
        Case 20
            strRet = "xx"
    End Select
    
    rlxRoman = strRet

End Function
'--------------------------------------------------------------
'　カラーダイアログ表示
'--------------------------------------------------------------
Public Function rlxGetColorDlg(lngDefColor As Long) As Long
Attribute rlxGetColorDlg.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetColorDlg.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim lngBackColor As Long
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim strColor As String
    
    strColor = Right$("000000" & Hex(lngDefColor), 6)
    lngRed = CLng("&H" & Mid$(strColor, 5, 2))
    lngGreen = CLng("&H" & Mid$(strColor, 3, 2))
    lngBlue = CLng("&H" & Mid$(strColor, 1, 2))
    
    If ActiveWorkbook Is Nothing Then
        rlxGetColorDlg = -2
        Exit Function
    End If
    
    lngBackColor = ActiveWorkbook.Colors(1)
    If Application.Dialogs(xlDialogEditColor).Show(1, lngRed, lngGreen, lngBlue) Then
        rlxGetColorDlg = ActiveWorkbook.Colors(1)
        ActiveWorkbook.Colors(1) = lngBackColor
    Else
        rlxGetColorDlg = -1
    End If

End Function
'--------------------------------------------------------------
'　クリップボードからファイル名を取得
'--------------------------------------------------------------
Public Function rlxGetFileNameFromCli() As String
Attribute rlxGetFileNameFromCli.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetFileNameFromCli.VB_ProcData.VB_Invoke_Func = " \n19"

#If VBA7 And Win64 Then
    Dim hData As LongPtr
#Else
    Dim hData As Long
#End If
    Dim files As Long
    Dim r As Long
    Dim i As Long
    Dim strFilePath As String
    Dim ret As String
    Const DEF_FILE_PATH_MAX_SIZE As Long = 1024 + 1
    
    If OpenClipboard(0) <> 0 Then
   
        hData = GetClipboardData(CF_HDROP)
        If Not IsNull(hData) Then
            files = DragQueryFile(hData, &HFFFFFFFF, 0, 0)
            For i = 0 To files - 1 Step 1
                strFilePath = String(DEF_FILE_PATH_MAX_SIZE, vbNullChar)
                Call DragQueryFile(hData, i, strFilePath, DEF_FILE_PATH_MAX_SIZE)
                If i = 0 Then
                    ret = Mid$(strFilePath, 1, InStr(strFilePath, vbNullChar) - 1)
                Else
                    ret = ret & vbCrLf & Mid$(strFilePath, 1, InStr(strFilePath, vbNullChar) - 1)
                End If
            Next
        End If
        r = CloseClipboard()
    
    End If
        rlxGetFileNameFromCli = ret
    
End Function
'--------------------------------------------------------------
'　Excelファイル判定
'--------------------------------------------------------------
Function rlxIsExcelFile(ByVal strFile As String) As Boolean
Attribute rlxIsExcelFile.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxIsExcelFile.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim varExt As Variant
    Dim i As Long
    rlxIsExcelFile = False
    
    varExt = Array(".XLSX", ".XLSM", ".XLS", ".XLB")

    For i = LBound(varExt) To UBound(varExt)
    
        If InStr(UCase(strFile), varExt(i)) > 0 Then
            rlxIsExcelFile = True
            Exit For
        End If
    
    Next

End Function
'--------------------------------------------------------------
'　PowerPointファイル判定
'--------------------------------------------------------------
Function rlxIsPowerPointFile(ByVal strFile As String) As Boolean

    Dim varExt As Variant
    Dim i As Long
    rlxIsPowerPointFile = False
    
    varExt = Array(".PPT", ".PPTX")

    For i = LBound(varExt) To UBound(varExt)
    
        If InStr(UCase(strFile), varExt(i)) > 0 Then
            rlxIsPowerPointFile = True
            Exit For
        End If
    
    Next

End Function
'--------------------------------------------------------------
'　Wordファイル判定
'--------------------------------------------------------------
Function rlxIsWordFile(ByVal strFile As String) As Boolean

    Dim varExt As Variant
    Dim i As Long
    rlxIsWordFile = False
    
    varExt = Array(".DOC", ".DOCX")

    For i = LBound(varExt) To UBound(varExt)
    
        If InStr(UCase(strFile), varExt(i)) > 0 Then
            rlxIsWordFile = True
            Exit For
        End If
    
    Next

End Function
'--------------------------------------------------------------
'　タイトルバー点滅
'--------------------------------------------------------------
Sub rlxFlashWindow()

#If VBA7 And Win64 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
    Dim udtFLASHWINFO As FLASHWINFO
    
    Const FLASH_STOP = &H0
    Const FLASH_CAPTION = &H1
    Const FLASH_TRAY = &H2
    Const FLASH_ALL = FLASH_CAPTION Or FLASH_TRAY
    Const FLASH_TIMER = &H4
    Const FLASH_TIMERNOFG = &HC

    hWnd = FindWindow("XLMAIN", Application.Caption)
    
    '点滅の設定
    With udtFLASHWINFO
        .cbsize = Len(udtFLASHWINFO)
        .hWnd = hWnd
        .dwFlags = FLASH_ALL
        .uCount = 5
        .dwTimeout = 100
    End With

    '点滅実行
    Call FlashWindowEx(udtFLASHWINFO)
    
End Sub
'--------------------------------------------------------------
'　エラーメッセージ表示
'--------------------------------------------------------------
Sub rlxErrMsg(ByRef objErr As Object)

    Select Case objErr.Number
        Case 0
        Case 1004
            MsgBox "エラーです。シート保護などを確認してください。", vbCritical + vbOKOnly, C_TITLE
        Case Else
            MsgBox objErr.Description & "(" & Err.Number & ")", vbCritical + vbOKOnly, C_TITLE
    End Select

End Sub
'----------------------------------------------------------------------
' クリップボードのビットマップデータから Picture オブジェクトを作成
'----------------------------------------------------------------------
Public Function CreatePictureFromClipboard(o As Object) As StdPicture
  
#If VBA7 And Win64 Then
    Dim hImg      As LongPtr
    Dim hPalette As LongPtr
    Dim hCopy As LongPtr
#Else
    Dim hImg      As Long
    Dim hPalette As Long
    Dim hCopy As Long
#End If
    
    Dim uPictDesc As PICTDESC
    Dim uGUID     As GUID
    
    Set CreatePictureFromClipboard = Nothing
  
'    Dim c As New Collection
    
    On Error GoTo e
    
    'クリップボードの保存
'    SaveClipData c
    If OpenClipboard(0&) <> 0 Then
        Call EmptyClipboard
        Call CloseClipboard
    End If
    
    '指定シェイプをビットマップでクリップボードに貼り付け
10  o.CopyPicture Appearance:=xlScreen, Format:=xlBitmap

    Call CopyClipboardSleep
        
    If IsClipboardFormatAvailable(CF_BITMAP) <> 0 Then
    
        If OpenClipboard(0&) <> 0 Then
            
            hImg = GetClipboardData(CF_BITMAP)
        
            If hImg <> 0 Then
          
                hCopy = CopyImage(hImg, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
                
                With uPictDesc
                    .cbSizeofStruct = Len(uPictDesc)
                    .picType = 1
                    .hImage = hCopy
                    .Option1 = 0&
                End With
                
                With uGUID
                    .Data1 = &H20400
                    .Data4(0) = &HC0
                    .Data4(7) = &H46
                End With
                
    '            Call OleCreatePictureIndirect(uPictDesc, uGUID, 0&, CreatePictureFromClipboard)
                Call OleCreatePictureIndirect(uPictDesc, uGUID, True, CreatePictureFromClipboard)
            
                Call EmptyClipboard
                
            End If
            
            Call CloseClipboard
        End If
        
    End If
    
    'クリップボードの復元
'    RestoreClipData c
    Exit Function
e:
    If Erl = 10 Then
        Resume
    End If
    
End Function
'--------------------------------------------------------------
'クリップボードにデータを保存するプロシージャ
'--------------------------------------------------------------
Public Sub SaveClipData(c As Collection)

#If VBA7 And Win64 Then
    Dim lngHwnd As LongPtr
    Dim lngMEM As LongPtr
    Dim lngDst As LongPtr
    Dim lngSrc As LongPtr
    Dim lngDataLen As LongPtr
    Dim lngRet As LongPtr
#Else
    Dim lngHwnd As Long
    Dim lngMEM As Long
    Dim lngDst As Long
    Dim lngSrc As Long
    Dim lngDataLen As Long
    Dim lngRet As Long
#End If
    Const GMEM_MOVEABLE = 2
    Dim lngFormatID As Long
    Dim s As ClipDataDTO

    'クリップボードをオープン
    If OpenClipboard(0&) <> 0 Then
  
        lngFormatID = EnumClipboardFormats(0)
        
        Do Until lngFormatID = 0
        
            'クリップボードに指定の形式が存在するか
            If IsClipboardFormatAvailable(lngFormatID) <> 0 Then
            
                lngMEM = GetClipboardData(lngFormatID)
        
                If lngMEM <> 0 Then
                
                    lngDataLen = GlobalSize(lngMEM)
                    
                    If lngDataLen <> 0 Then
                
                        'グローバルメモリに書き込む領域を確保してそのハンドルを取得
                        lngHwnd = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
                        
                        If lngHwnd <> 0 Then
                            
                            'グローバルメモリをロックしてそのポインタを取得
                            lngDst = GlobalLock(lngHwnd)
                            lngSrc = GlobalLock(lngMEM)
                    
                            CopyMemory ByVal lngDst, ByVal lngSrc, lngDataLen
                            
                            Call GlobalUnlock(lngMEM)
                            Call GlobalUnlock(lngHwnd)
                            
                            Set s = New ClipDataDTO
                            
                            s.lngFormat = lngFormatID
                            s.lngHandle = lngHwnd
                            
                            c.Add s
                            
                            Set s = Nothing
                
                        End If
                        
                    End If
                    
                End If
                
            End If
        
            lngFormatID = EnumClipboardFormats(lngFormatID)
            'Exit Do
        Loop
        
        Call EmptyClipboard

        'クリップボードをクローズ(これはWindowsに制御が
        '戻らないうちにできる限り速やかに行う)
        lngRet = CloseClipboard()
        
    End If

End Sub
'--------------------------------------------------------------
'クリップボードにデータを復元するプロシージャ
'--------------------------------------------------------------
Public Sub RestoreClipData(c As Collection)

#If VBA7 And Win64 Then
    Dim lngMEM As LongPtr
    Dim lngRet As LongPtr
#Else
    Dim lngMEM As Long
    Dim lngRet As Long
#End If

    Const GMEM_MOVEABLE = 2
  
    Dim s As ClipDataDTO
    
    If c.Count = 0 Then
        Exit Sub
    End If

    'クリップボードをオープン
    If OpenClipboard(0&) <> 0 Then
  
        'クリップボードを空にする
        If EmptyClipboard() <> 0 Then
    
            For Each s In c
        
                If s.lngHandle <> 0 Then
        
                    'グローバルメモリをロックしてそのポインタを取得
                    lngMEM = GlobalLock(s.lngHandle)
              
                    If lngMEM <> 0 Then
                    
                        'クリップボードにメモリブロックのデータを書き込み
                        lngRet = SetClipboardData(s.lngFormat, s.lngHandle)
                    
                        'グローバルメモリブロックのロックを解除
                        lngRet = GlobalUnlock(s.lngHandle)
                        
                        'lngRet = GlobalFree(s.lngHandle)
                    
                    End If
                  
                End If
          
            Next
        End If
    
    End If
    
    lngRet = CloseClipboard()

End Sub
'--------------------------------------------------------------
'クリップボードをクリアする
'--------------------------------------------------------------
Public Sub ClearClipboard()

    If OpenClipboard(0&) <> 0 Then
        Call EmptyClipboard
        Call CloseClipboard
    End If

End Sub
'--------------------------------------------------------------
'文字化け対応StrConv(vbUnicode, vbFromUnicodeは使えません)
'--------------------------------------------------------------
Public Function StrConvU(ByVal strSource As String, conv As VbStrConv) As String

    Dim i As Long
    Dim strBuf As String
    Dim c As String
    Dim strRet As String
    Dim strBefore As String
    Dim strChr As String
    Dim strNext As String

    strRet = ""
    strBuf = ""
    strBefore = ""
    strNext = ""

    For i = 1 To Len(strSource)

        c = Mid$(strSource, i, 1)
        
        If i = Len(strSource) Then
            strNext = ""
        Else
            strNext = Mid$(strSource, i + 1, 1)
        End If

        Select Case c
            '全角の￥
            Case "￥"
                If (conv And vbNarrow) > 0 Then
                    strChr = "\"
                    strRet = strRet & StrConv(strBuf, conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
           
            '半角の\
            Case "\"
                If (conv And vbWide) > 0 Then
                    strChr = "￥"
                    strRet = strRet & StrConv(strBuf, conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
            '全角の濁点、半濁点
            Case "゜", "゛"
                If (conv And vbNarrow) > 0 Then
                    If c = "゜" Then
                        strChr = "ﾟ"
                    Else
                        strChr = "ﾞ"
                    End If
                    strRet = strRet & StrConv(strBuf, conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
                
            '半角の半濁点
            Case "ﾟ"
                '１つ前の文字
                Select Case strBefore
                    Case "ﾊ" To "ﾎ"
                        strBuf = strBuf & c
                    Case Else
                        If (conv And vbWide) > 0 Then
                             strChr = "゜"
                            strRet = strRet & StrConv(strBuf, conv) & strChr
                            strBuf = ""
                        Else
                            strBuf = strBuf & c
                        End If
                End Select
                
            '半角の濁点
            Case "ﾞ"
                '１つ前の文字
                Select Case strBefore
                    Case "ｳ", "ｶ" To "ｺ", "ｻ" To "ｿ", "ﾀ" To "ﾄ", "ﾊ" To "ﾎ"
                        strBuf = strBuf & c
                    Case Else
                        If (conv And vbWide) > 0 Then
                            strChr = "゛"
                            strRet = strRet & StrConv(strBuf, conv) & strChr
                            strBuf = ""
                        Else
                            strBuf = strBuf & c
                        End If
                End Select
            'ヴ
            Case "ヴ"
                If (conv And vbHiragana) > 0 Then
                    Dim b() As Byte
                    ReDim b(0 To 1)
                    b(0) = &H94
                    b(1) = &H30
                    strChr = b
                    strRet = strRet & StrConv(strBuf, conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
            'う゛
            Case "う"
                If strNext = "゛" And (conv And vbKatakana) > 0 Then
                    strChr = "ヴ"
                    strRet = strRet & StrConv(strBuf, conv) & strChr
                    strBuf = ""
                    i = i + 1
                Else
                    strBuf = strBuf & c
                End If

            'ヶヵ
            Case "ヶ", "ヵ"
                If (conv And vbHiragana) > 0 Then
                    strRet = strRet & StrConv(strBuf, conv) & c
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If

            'その他
            Case Else
                '第二水準等StrConvで文字化けするものを退避
                If Asc(c) = 63 And c <> "?" Then
                    strRet = strRet & StrConv(strBuf, conv) & c
                    strBuf = ""
                Else
                    'う”
                    If Unicode(c) = &H3094 Then
                        If conv = vbKatakana Then
                            strRet = strRet & StrConv(strBuf, conv) & "ヴ"
                            strBuf = ""
                        Else
                            strRet = strRet & StrConv(strBuf, conv) & c
                            strBuf = ""
                        End If
                    Else
                        strBuf = strBuf & c
                    End If
                End If
        End Select
        
        '１個前の文字
        strBefore = c

    Next

    If strBuf <> "" Then
        strRet = strRet & StrConv(strBuf, conv)
    End If

    StrConvU = strRet

End Function
Private Function Unicode(ByVal strBuf As String) As Long
    Dim bytBuf() As Byte
    
    If Len(strBuf) <> 0 Then
        bytBuf = strBuf
        Unicode = CLng(bytBuf(1)) * &H100 + bytBuf(0)
    End If
End Function
'--------------------------------------------------------------
'  フォルダの作成
'--------------------------------------------------------------
Sub rlxCreateFolder(ByVal strPath As String)

    Dim v As Variant
    Dim s As Variant
    
    Dim f As String
    
    v = Split(strPath, "\")

    On Error Resume Next
    For Each s In v
    
        If f = "" Then
            f = s
            MkDir f & "\"
        Else
            f = f & "\" & s
            MkDir f
        End If
    
    Next
    Err.Clear

End Sub
'--------------------------------------------------------------
'  Excelファイルを開き直す
'--------------------------------------------------------------
Sub CloseAndOpen()

    Dim strFile As String
    Dim blnReadOnly As Boolean
    
    strFile = ActiveWorkbook.FullName
    
    If rlxIsFileExists(strFile) Then
    
        Application.ScreenUpdating = False
        
        blnReadOnly = ActiveWorkbook.ReadOnly
        
        ActiveWorkbook.Close
        
        Workbooks.Open strFile, ReadOnly:=blnReadOnly
        
        Application.ScreenUpdating = True
    
    End If

End Sub
'--------------------------------------------------------------
'  現在アクティブなブックをRenameする
'--------------------------------------------------------------
Sub RenameActiveBook()

    Dim WB As Workbook
    Dim strFile As String
    Dim strBook As String
    Dim strPath As String
    Dim strBuf As String
    Dim strNew As String
    
    Dim strExt As String
    Dim strName As String
    
    Set WB = ActiveWorkbook
    strBook = WB.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "ブックが保存されていません。保存してから実行してください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    strFile = rlxGetFullpathFromFileName(strBook)
    strPath = rlxGetFullpathFromPathName(strBook)
    
    '読み取り専用の場合は名前を変更させない
    If ActiveWorkbook.ReadOnly Then
        MsgBox "読み取り専用の場合には名前を変更できません。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    strName = rlxGetFullpathFromExt(strFile)
    
    '「アクティブなブックの名前を変更」にてファイル名の最初のピリオド以下を拡張子として扱う問題の修正 #59
    strExt = Mid(strFile, InStrRev(strFile, "."))
    
    strBuf = InputBox("変更後のブック名を入力してください。", "アクティブなブックの名前を変更", strName)
    If Trim(strBuf) = "" Then
        Exit Sub
    End If
    
    '変更後のブックを合成
    strNew = rlxAddFileSeparator(strPath) & strBuf & strExt
    
    Dim s As Workbook
    Dim strNewName As String
    strNewName = rlxGetFullpathFromFileName(strNew)
    For Each s In Workbooks
        If LCase(s.Name) = LCase(strNewName) Then
            MsgBox "同名のブックを開いているため名前の変更ができません。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
    Next
    
    If rlxIsFileExists(strNew) Then
        MsgBox "すでに変更後のブックが存在します。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Err.Clear
    WB.SaveAs filename:=strNew, local:=True
    If Err.Number <> 0 Then
        Application.ScreenUpdating = True
        MsgBox "エラーが発生しました。名前の変更に失敗しました。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    Else
        Kill strBook
    End If
    Application.ScreenUpdating = True
    
End Sub
'--------------------------------------------------------------
'  全角トリム
'--------------------------------------------------------------
Function TrimZen(ByVal strBuf As String) As String
 
    Dim lngLen As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    
    lngLen = Len(strBuf)
    
    lngStart = 1
    Do Until lngStart > lngLen
        Select Case Mid$(strBuf, lngStart, 1)
            Case Is <= " "
            Case Is = "　"
            Case Else
                Exit Do
        End Select
        lngStart = lngStart + 1
    Loop
    
    lngEnd = lngLen
    Do Until lngEnd < 1
        Select Case Mid$(strBuf, lngEnd, 1)
            Case Is <= " "
            Case Is = "　"
            Case Else
                Exit Do
        End Select
        lngEnd = lngEnd - 1
    Loop
    
    If lngEnd > 0 Or lngStart <= lngLen Then
        TrimZen = Mid$(strBuf, lngStart, (lngEnd - lngStart) + 1)
    Else
        TrimZen = ""
    End If

End Function
'--------------------------------------------------------------
'  マイドキュメントフォルダ移動
'--------------------------------------------------------------
Sub SetMyDocument()
    On Error Resume Next
    ChDir CreateObject("Wscript.Shell").SpecialFolders("MyDocuments")
End Sub

'--------------------------------------------------------------
'  サロゲートペア対応Len
'--------------------------------------------------------------
Function LenEx(ByVal strBuf As String) As Long

    Dim bytBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long
    
    lngLen = 0
    
    If Len(strBuf) = 0 Then
        LenEx = 0
        Exit Function
    End If
    
    bytBuf = strBuf
    
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            '上位サロゲート
            Case &HD800& To &HDBFF&
                lngLen = lngLen + 1
            '下位サロゲート
            Case &HDC00& To &HDFFF&
                'カウントしない
            Case Else
                lngLen = lngLen + 1
        End Select
    
    Next
    
    LenEx = lngLen

End Function
'------------------------------------------------------------------------------------------------------------------------
' 下位バイト取得
'------------------------------------------------------------------------------------------------------------------------
Function LByte(ByVal lngValue As Long) As Long
    LByte = lngValue And &HFF&
End Function
'------------------------------------------------------------------------------------------------------------------------
' 上位バイト取得
'------------------------------------------------------------------------------------------------------------------------
Function UByte(ByVal lngValue As Long) As Long
    UByte = RShift((lngValue And &HFF00&), 8)
End Function
'------------------------------------------------------------------------------------------------------------------------
' 右シフト
'------------------------------------------------------------------------------------------------------------------------
Function RShift(ByVal lngValue As Long, ByVal lngKeta As Long) As Long
    RShift = lngValue \ (2 ^ lngKeta)
End Function
'------------------------------------------------------------------------------------------------------------------------
' 左シフト
'------------------------------------------------------------------------------------------------------------------------
Function LShift(ByVal lngValue As Long, ByVal lngKeta As Long) As Long
    LShift = lngValue * (2 ^ lngKeta)
End Function
'------------------------------------------------------------------------------------------------------------------------
' 非同期実行
'------------------------------------------------------------------------------------------------------------------------
Function UnSyncRun(ByVal strMacro As String, Optional ByVal lngSec As Long = 0) As Long
    Application.OnTime DateAdd("s", lngSec, Now), MacroHelper.BuildPath(strMacro)
End Function
Function getHtmlRGB(ByVal lngColor As Variant) As String

    Dim strBuf As String
    If IsNull(lngColor) Then
        getHtmlRGB = "#000000"
    Else
    
        strBuf = Right$("000000" & Hex$(lngColor), 6)
    
        getHtmlRGB = "#" & Mid$(strBuf, 5, 2) & Mid$(strBuf, 3, 2) & Mid$(strBuf, 1, 2)
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------
' CopyPictureがJavaアプリやクリップボードツールなどで失敗する対策
'------------------------------------------------------------------------------------------------------------------------
Public Sub CopyClipboardSleep()
    DoEvents
    Sleep Val(GetSetting(C_TITLE, "Option", "ClipboardSleep", 0))
End Sub

'--------------------------------------------------------------
'  UNICODE対応ひらがな→カタカナ変換
'--------------------------------------------------------------
Function ToKatakana(ByVal strBuf As String, Optional ByVal flag As Boolean = False) As String

    Dim bytBuf() As Byte
    Dim retBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long
    Dim lngConv As Long
    Dim lngOpt As Long
    
    lngLen = 0
    
    If Len(strBuf) = 0 Then
        ToKatakana = ""
        Exit Function
    End If
    
    bytBuf = strBuf
    retBuf = strBuf
    
    If flag Then
        lngOpt = &H3096&
    Else
        lngOpt = &H3094&
    End If
    
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'ひらがな
            Case &H3041& To lngOpt, &H309D&, &H309E&
            
                lngConv = lngBuf + &H60&
                retBuf(i) = LByte(lngConv)
                retBuf(i + 1) = UByte(lngConv)
            
        End Select
    
    Next
    
    ToKatakana = retBuf()

End Function
'--------------------------------------------------------------
'  UNICODE対応カタカナ→ひらがな変換
'--------------------------------------------------------------
Function ToHiragana(ByVal strBuf As String, Optional ByVal flag As Boolean = False) As String

    Dim bytBuf() As Byte
    Dim retBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long
    Dim lngConv As Long
    Dim lngOpt As Long
    
    lngLen = 0
    
    If Len(strBuf) = 0 Then
        ToHiragana = ""
        Exit Function
    End If
    
    bytBuf = strBuf
    retBuf = strBuf
    
    If flag Then
        lngOpt = &H30F6&
    Else
        lngOpt = &H30F4&
    End If
    
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'カタカナ
            Case &H30A1& To lngOpt, &H30FD&, &H30FE&
            
                lngConv = lngBuf - &H60&
                retBuf(i) = LByte(lngConv)
                retBuf(i + 1) = UByte(lngConv)
            
        End Select
    
    Next
    
    ToHiragana = retBuf()

End Function
'シート名不可文字チェック
Function IsErrSheetNameChar(ByVal strBuf As String) As Boolean

    Dim lngChar As Long
    Dim i As Long
    Dim bytBuf() As Byte
    
    IsErrSheetNameChar = False
    
    '空はエラー
    If Len(strBuf) = 0 Then
        IsErrSheetNameChar = True
        Exit Function
    End If
    
    '履歴は予約語なので
    If strBuf = "履歴" Then
        IsErrSheetNameChar = True
        Exit Function
    End If
    
    bytBuf = strBuf
    
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngChar = LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngChar
            'エラーになる文字'*/:?[\]’＇＊／：？［＼］￥
            Case &H0&, &H27&, &H2A&, &H2F&, &H3A&, &H3F&, &H5B&, &H5C&, &H5D&, &H2019&, &HFF07&, &HFF0A&, &HFF0F&, &HFF1A&, &HFF1F&, &HFF3B&, &HFF3C&, &HFF3D&, &HFFE5&
                IsErrSheetNameChar = True
                Exit Function
                
        End Select
    
    Next

End Function
Function getVersionInfo()

    Dim strVer As String
    Dim strTitle As String

    strTitle = ThisWorkbook.BuiltinDocumentProperties("Title").Value
    strVer = ThisWorkbook.BuiltinDocumentProperties("Comments").Value
    
    Dim strBuf As String
    Dim i As Long
    Dim obj As Object

    strBuf = strTitle
    Dim s() As String
    s = Split(strVer, vbLf)
    strBuf = strBuf & " " & s(0) & vbCrLf
    
    strBuf = strBuf & "Microsoft "
    Select Case True
        Case InStr(Application.OperatingSystem, "5.00") > 0
            strBuf = strBuf & "Windows 2000"
        Case InStr(Application.OperatingSystem, "5.01") > 0
            strBuf = strBuf & "Windows XP"
        Case InStr(Application.OperatingSystem, "6.00") > 0
            strBuf = strBuf & "Windows Vista"
        Case InStr(Application.OperatingSystem, "6.01") > 0
            strBuf = strBuf & "Windows 7"
        Case InStr(Application.OperatingSystem, "6.02") > 0
            strBuf = strBuf & "Windows 8 or 8.1"
        Case Else
            strBuf = strBuf & "Windows 10 or Later"
    End Select
    If Isx64 Then
        strBuf = strBuf & " (64bit)"
    Else
        strBuf = strBuf & " (32bit)"
    End If
    
    strBuf = strBuf & vbCrLf

    strBuf = strBuf & "Microsoft Excel "

    Select Case Val(Application.Version)
        Case Is = 0
            strBuf = strBuf & "不明"
        Case Is <= 11
            strBuf = strBuf & "2003以前"
        Case 12
            strBuf = strBuf & "2007"
        Case 14
            strBuf = strBuf & "2010"
        Case 15
            strBuf = strBuf & "2013"
        Case 16
            strBuf = strBuf & "2016"
        Case Else
            strBuf = strBuf & "2016より未来のバージョン"
    End Select
    strBuf = strBuf & " Build " & Application.Build
#If Win64 Then
    strBuf = strBuf & " (64bit)"
#Else
    strBuf = strBuf & " (32bit)"
#End If

    getVersionInfo = strBuf

End Function
Private Function Isx64() As Boolean

    On Error GoTo xp

    Dim colItems As Object
    Dim itm As Object
    Dim ret As Boolean

    ret = False '初期化

    Set colItems = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery("Select * From Win32_OperatingSystem")

    For Each itm In colItems
        If InStr(itm.OSArchitecture, "64") Then
            ret = True
            Exit For
        End If
    Next

    Isx64 = ret

    Exit Function
xp:
    Isx64 = False

End Function
'マルチプロセス実行用
Public Sub MultiProsess(ByVal strMacro As String)

    frmMulti.Show

'    With CreateObject("WScript.Shell")
'        .Run (.SpecialFolders("AppData") & "\" & C_TITLE & "\" & "RunMacro.vbs """ & strMacro & """")
'    End With

'    Dim strAddin As String
'
'    strAddin = CreateObject("WScript.Shell").SpecialFolders("Appdata") & "\Microsoft\Addins\Relaxtools.xlam"
    
    Err.Clear
    
    On Error Resume Next
    With CreateObject("Excel.Application")
        .Workbooks.Open ThisWorkbook.FullName
        .Run strMacro
    End With
    
    
    Unload frmMulti
    
    If Err.Number <> 0 Then
        MsgBox "Grep検索(Multi Process版)の起動に失敗しました。", vbCritical + vbOKOnly, C_TITLE
    End If
    
    
End Sub
'**
' コピーアドレスの取得
'**
Public Function getObjectLink() As String

#If VBA7 And Win64 Then
  Dim lngHwnd As LongPtr, lngMEM As LongPtr
  Dim lngDataLen As LongPtr
  Dim lngRet As LongPtr
#Else
  Dim lngHwnd As Long, lngMEM As Long
  Dim lngDataLen As Long
  Dim lngRet As Long
#End If

    Const MAXSIZE = 4096
    Dim MyString As String
    Dim size As Long
    Dim data() As Byte
    Dim i As Long
  
    'クリップボードをオープン
    If OpenClipboard(0&) <> 0 Then
    
'        lngMEM = GetClipboardData(RegisterClipboardFormat("Link"))
        lngMEM = GetClipboardData(RegisterClipboardFormat("ObjectLink"))
        
        If lngMEM <> 0 Then
        
            size = CLng(GlobalSize(lngMEM))
            lngHwnd = GlobalLock(lngMEM)
            
            If lngHwnd <> 0 Then
                
                ReDim data(0 To size - 1)
                Call CopyMemory(data(0), ByVal lngHwnd, size)
                
                lngRet = GlobalUnlock(lngMEM)
                
                For i = 0 To size - 1
                    If data(i) = 0 Then
                        data(i) = &H9
                    End If
                Next i
                MyString = StrConv(data(), vbUnicode)
                
            End If
        
        End If
        
        lngRet = CloseClipboard()
    
    End If
    
    getObjectLink = MyString

End Function
Function SpecialCellsEx(v As Range) As Range

    Dim r As Range
    Dim c As Range
    Dim f As Range

    On Error Resume Next

    Set c = v.SpecialCells(xlCellTypeConstants)
    Set f = v.SpecialCells(xlCellTypeFormulas)

    Set r = Nothing

    If Not c Is Nothing Then
        Set r = c
    End If

    If Not f Is Nothing Then
        If r Is Nothing Then
            Set r = f
        Else
            Set r = v.Application.Union(r, f)
        End If
    End If

    Set SpecialCellsEx = r

End Function
Function FileToURL(ByVal arg As String) As String

    FileToURL = "file:///" & Replace(arg, "#", "%23")

End Function
