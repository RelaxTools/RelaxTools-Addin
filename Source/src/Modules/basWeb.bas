Attribute VB_Name = "basWeb"
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

Public Const C_GITHUB_URL As String = "https://github.com/RelaxTools/RelaxTools-Addin"
Public Const C_URL As String = "http://software.opensquare.net/relaxtools/"
Public Const C_REGEXP_URL As String = "http://software.opensquare.net/relaxtools/about/foruse/regexp/"
Public Const C_STAMP_URL As String = "http://software.opensquare.net/relaxtools/about/foruse/stamp/"
Public Const C_CAMPAIGN_URL As String = "http://software.opensquare.net/relaxtools/support-2/campaign/"
Public Const C_FCS_URL As String = "https://www.fcs.co.jp/?relaxtools"

Public Const C_RELEASE_URL As String = "https://github.com/RelaxTools/RelaxTools-Addin/releases"
Public Const C_MADO_URL As String = "http://forest.watch.impress.co.jp/library/software/relaxtools/"
Public Const C_BBS_URL As String = "http://software.opensquare.net/relaxtools/bbs/wforum.cgi"
Public Const C_ISSUE_URL As String = "https://github.com/RelaxTools/RelaxTools-Addin/issues"
Public Const C_MAIL_URL As String = "mailto:relaxtools@opensquare.net"


Public Sub GotoGitHub()
    Call GoURL(C_GITHUB_URL)
End Sub
Public Sub GotoOfficialWeb()
    Call GoURL(C_URL)
End Sub
Public Sub GotoCampaign()
    Call GoURL(C_CAMPAIGN_URL)
End Sub
Public Sub GotoRegExpHelp()
    Call GoURL(C_REGEXP_URL)
End Sub
Public Sub GotoRelease()
    Call GoURL(C_RELEASE_URL)
End Sub
Public Sub GotoMado()
    Call GoURL(C_MADO_URL)
End Sub
Public Sub GotoBBS()
    Call GoURL(C_BBS_URL)
End Sub
Public Sub GotoIssue()
    Call GoURL(C_ISSUE_URL)
End Sub
Public Sub GotoMail()
    Call GoURL(C_MAIL_URL)
End Sub
Public Sub GotoFCS()
    Call GoURL(C_FCS_URL)
End Sub
Public Sub GoURL(ByVal strURL As String)
    With CreateObject("WScript.Shell")
        .Run (strURL)
    End With
End Sub
