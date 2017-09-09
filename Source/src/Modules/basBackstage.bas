Attribute VB_Name = "basBackstage"
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
'--------------------------------------------------------------------
' ラベルを表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub getVersion(control As IRibbonControl, ByRef Screentip)

    Dim strVer As String

    strVer = ThisWorkbook.BuiltinDocumentProperties("Title").Value
    strVer = strVer & " " & ThisWorkbook.BuiltinDocumentProperties("Comments").Value
    
    Screentip = strVer

End Sub
Public Sub getExcelVersion(control As IRibbonControl, ByRef Screentip)

    Dim strBuf As String

     strBuf = getVersionInfo & vbCrLf & vbCrLf

    strBuf = strBuf & "上記の情報とエラー内容をお知らせください。" & vbCrLf
    strBuf = strBuf & "以下、３つの方法があります。" & vbCrLf & vbCrLf
    
    strBuf = strBuf & "◆GitHub Issue(GitHubのアカウントが必要です)：" & vbCrLf
    strBuf = strBuf & "https://github.com/RelaxTools/RelaxTools-Addin/issues" & vbCrLf & vbCrLf
    
    strBuf = strBuf & "◆掲示板：" & vbCrLf
    strBuf = strBuf & "http://software.opensquare.net/relaxtools/bbs/wforum.cgi" & vbCrLf & vbCrLf
    
    strBuf = strBuf & "◆メール(relaxtools@opensquare.net)でも受け付けます。" & vbCrLf
    
    Screentip = strBuf

End Sub
Public Sub getLisence1(control As IRibbonControl, ByRef Screentip)
    Dim strBuf As String
   
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
    
    Screentip = strBuf
    
End Sub
Public Sub getLisence2(control As IRibbonControl, ByRef Screentip)
    Dim strBuf As String
   
    strBuf = strBuf & " THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR" & vbCrLf
    strBuf = strBuf & " IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY," & vbCrLf
    strBuf = strBuf & " FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE" & vbCrLf
    strBuf = strBuf & " AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER" & vbCrLf
    strBuf = strBuf & " LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING " & vbCrLf
    strBuf = strBuf & " FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER " & vbCrLf
    strBuf = strBuf & " DEALINGS IN THE SOFTWARE." & vbCrLf
    
    Screentip = strBuf
    
End Sub
Public Sub OnHide(contextObject As Object) ' Backstage ビューが閉じた時の処理
    Call StopSushi
End Sub
