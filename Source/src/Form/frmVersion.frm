VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVersion 
   Caption         =   "バージョン情報"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10245
   OleObjectBlob   =   "frmVersion.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmVersion"
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

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub lblGitHub_Click()
    Call GotoGitHub
End Sub

Private Sub lblUrl_Click()

    Dim WSH As Object
    
    Set WSH = CreateObject("WScript.Shell")
    
    Call WSH.Run(C_URL)
    
    Set WSH = Nothing

End Sub


Private Sub txtDebug_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = txtDebug
End Sub
Private Sub UserForm_Activate()
    MW.Activate
End Sub
Private Sub UserForm_Initialize()

    Dim strVer As String
    Dim strTitle As String

    strTitle = ThisWorkbook.BuiltinDocumentProperties("Title").value
    strVer = ThisWorkbook.BuiltinDocumentProperties("Comments").value
    'MsgBox strTitle & "            " & vbCrLf & strVer, vbInformation, C_TITLE
    
    
    lblTitle.Caption = strTitle
    lblComment.Caption = strTitle & " " & strVer
    lblUrl.Caption = C_URL
    lblGitHub.Caption = C_GITHUB_URL
    
    Dim strBuf As String
    Dim col As New Collection
    Dim i As Long
    Dim obj As Object

'    col.Add "Scripting.FileSystemObject"
'    col.Add "Shell.Application"
'    col.Add "WScript.Shell"
'    col.Add "VBScript.RegExp"
'    col.Add "System.Collections.ArrayList"
'    col.Add "System.Text.StringBuilder"
'    col.Add "System.Text.UTF8Encoding"
'    col.Add "System.Security.Cryptography.MD5CryptoServiceProvider"
'    col.Add "System.Security.Cryptography.SHA1Managed"
'    col.Add "System.Security.Cryptography.SHA256Managed"
'    col.Add "System.Security.Cryptography.SHA384Managed"
'    col.Add "System.Security.Cryptography.SHA512Managed"
'
'    strBuf = strTitle
'    Dim s() As String
'    s = Split(strVer, vbLf)
'    strBuf = strBuf & " " & s(0) & vbCrLf
'
'    strBuf = strBuf & "Microsoft "
'    Select Case True
'        Case InStr(Application.OperatingSystem, "5.00") > 0
'            strBuf = strBuf & "Windows 2000"
'        Case InStr(Application.OperatingSystem, "5.01") > 0
'            strBuf = strBuf & "Windows XP"
'        Case InStr(Application.OperatingSystem, "6.00") > 0
'            strBuf = strBuf & "Windows Vista"
'        Case InStr(Application.OperatingSystem, "6.01") > 0
'            strBuf = strBuf & "Windows 7"
'        Case InStr(Application.OperatingSystem, "6.02") > 0
'            strBuf = strBuf & "Windows 8 or 8.1"
'        Case Else
'            strBuf = strBuf & "Windows 10 or Later"
'    End Select
'    If Isx64 Then
'        strBuf = strBuf & " (64bit)" & vbCrLf
'    Else
'        strBuf = strBuf & " (32bit)" & vbCrLf
'    End If
'
'    strBuf = strBuf & "Microsoft Excel "
'
'    Select Case Val(Application.Version)
'        Case Is = 0
'            strBuf = strBuf & "不明"
'        Case Is <= 11
'            strBuf = strBuf & "2003以前"
'        Case 12
'            strBuf = strBuf & "2007"
'        Case 14
'            strBuf = strBuf & "2010"
'        Case 15
'            strBuf = strBuf & "2013"
'        Case 16
'            strBuf = strBuf & "2016"
'        Case Else
'            strBuf = strBuf & "2013より未来のバージョン"
'    End Select
'    strBuf = strBuf & " Build " & Application.Build
'#If Win64 Then
'    strBuf = strBuf & " (64bit)" & vbCrLf
'#Else
'    strBuf = strBuf & " (32bit)" & vbCrLf
'#End If
'    strBuf = strBuf & "" & vbCrLf
'    For i = 1 To col.count
'
'        strBuf = strBuf & col.Item(i) & ":"
'        On Error Resume Next
'        err.Clear
'        Set obj = CreateObject(col.Item(i))
'        If err.Number <> 0 Or obj Is Nothing Then
'            strBuf = strBuf & "NG"
'        Else
'            strBuf = strBuf & "OK"
'        End If
'        Set obj = Nothing
'        On Error GoTo 0
'        strBuf = strBuf & vbCrLf
'
'    Next
    
   
    strBuf = strBuf & " [RelaxTools-Addin] v4" & vbCrLf
    strBuf = strBuf & "" & vbCrLf
    strBuf = strBuf & " Copyright (c) 2009 Yasuhiro Watanabe" & vbCrLf
    strBuf = strBuf & " https://github.com/RelaxTools/RelaxTools-Addin" & vbCrLf
    strBuf = strBuf & " author:relaxtools@opensquare.net" & vbCrLf
    strBuf = strBuf & "" & vbCrLf
    strBuf = strBuf & " The MIT License (MIT)" & vbCrLf
    strBuf = strBuf & "" & vbCrLf
    strBuf = strBuf & " Permission is hereby granted, free of charge, to any person obtaining a copy" & vbCrLf
    strBuf = strBuf & " of this software and associated documentation files (the ""Software""), to deal" & vbCrLf
    strBuf = strBuf & " in the Software without restriction, including without limitation the rights" & vbCrLf
    strBuf = strBuf & " to use, copy, modify, merge, publish, distribute, sublicense, and/or sell" & vbCrLf
    strBuf = strBuf & " copies of the Software, and to permit persons to whom the Software is" & vbCrLf
    strBuf = strBuf & " furnished to do so, subject to the following conditions:" & vbCrLf
    strBuf = strBuf & "" & vbCrLf
    strBuf = strBuf & " The above copyright notice and this permission notice shall be included in all" & vbCrLf
    strBuf = strBuf & " copies or substantial portions of the Software." & vbCrLf
    strBuf = strBuf & "" & vbCrLf
    strBuf = strBuf & " THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR" & vbCrLf
    strBuf = strBuf & " IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY," & vbCrLf
    strBuf = strBuf & " FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE" & vbCrLf
    strBuf = strBuf & " AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER" & vbCrLf
    strBuf = strBuf & " LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM," & vbCrLf
    strBuf = strBuf & " OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE" & vbCrLf
    strBuf = strBuf & " SOFTWARE." & vbCrLf
    
    txtDebug.Text = strBuf
    txtDebug.SelStart = Len(txtDebug.Text)
    txtDebug.SelStart = 0

'    txtDebug.SetFocus
'    SendKeys "^A"
    
    Set MW = basMouseWheel.GetInstance
    MW.Install
    
End Sub


Private Sub MW_WheelDown(obj As Object)
    
    Dim lngPos As Long
    
    On Error GoTo e
    lngPos = obj.CurLine + 3
    If lngPos >= obj.LineCount Then
        lngPos = obj.LineCount - 1
    End If
    obj.CurLine = lngPos
e:
End Sub

Private Sub MW_WheelUp(obj As Object)

    Dim lngPos As Long
    
    On Error GoTo e
    lngPos = obj.CurLine - 3
    If lngPos < 0 Then
        lngPos = 0
    End If
    obj.CurLine = lngPos
e:
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = Nothing
End Sub

Private Sub UserForm_Terminate()
    MW.UnInstall
    Set MW = Nothing
End Sub
