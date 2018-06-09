Attribute VB_Name = "basSetting"
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
Sub getCrossLineSetting(ByRef lngType As Long, ByRef blnFillVisible As Boolean, ByRef lngFillColor As Long, ByRef dblFillTransparency As Double, ByRef lngLineVisible As Long, ByRef lngLineColor As Long, ByRef sngLineWeight As Single, ByRef strOnAction As String, ByRef blnGuid As Boolean, ByRef lngFontColor As Long, ByRef blnEdit As Boolean, ByRef blnLineWidth As Boolean, ByRef lngFBackColor As Long, ByRef dblGuidTransparency As Double)

    lngType = Val(GetSetting(C_TITLE, "CrossLine", "Type", C_ALL))
'    blnFillVisible = GetSetting(C_TITLE, "CrossLine", "FillVisible", True)
    lngFillColor = CLng(GetSetting(C_TITLE, "CrossLine", "FillColor", "&H0050B000"))
    dblFillTransparency = Val(GetSetting(C_TITLE, "CrossLine", "FillTransparency", "50"))
        
'    lngLineVisible = Val(GetSetting(C_TITLE, "CrossLine", "LineVisible", "0"))
'    lngLineColor = CLng(GetSetting(C_TITLE, "CrossLine", "LineColor", "&H0050B000"))
'    sngLineWeight = Val(GetSetting(C_TITLE, "CrossLine", "LineWeight", "2"))
    
'    strOnAction = GetSetting(C_TITLE, "CrossLine", "OnAction", "rlxCrossClick")
'
    blnGuid = GetSetting(C_TITLE, "CrossLine", "Guid", True)
'    blnEdit = GetSetting(C_TITLE, "CrossLine", "Edit", False)
'    blnLineWidth = GetSetting(C_TITLE, "CrossLine", "LineWidth", False)
'
'    lngFontColor = CLng(GetSetting(C_TITLE, "CrossLine", "FontColor", "&H0050B000"))
'    lngFBackColor = CLng(GetSetting(C_TITLE, "CrossLine", "BackColor", "&H00FFFFFF"))
'
'    dblGuidTransparency = Val(GetSetting(C_TITLE, "CrossLine", "GuidTransparency", "100"))

End Sub
Sub setCrossLineSetting(ByVal strType As String, ByVal strFillVisible As String, ByVal strFillColor As String, ByVal strFillTransparency As String, ByVal strLineVisible As String, ByVal strLineColor As String, ByVal strLineWeight As String, ByVal blnGuid As Boolean, ByVal strFontColor As String, ByVal blnEdit As Boolean, ByVal blnLineWidth As Boolean, ByVal strFBackColor As String, ByVal strGuidTransparency As String)

    Call SaveSetting(C_TITLE, "CrossLine", "Type", strType)
'    Call SaveSetting(C_TITLE, "CrossLine", "FillVisible", strFillVisible)
    Call SaveSetting(C_TITLE, "CrossLine", "FillColor", strFillColor)
    Call SaveSetting(C_TITLE, "CrossLine", "FillTransparency", strFillTransparency)
        
'    Call SaveSetting(C_TITLE, "CrossLine", "LineVisible", strLineVisible)
'    Call SaveSetting(C_TITLE, "CrossLine", "LineColor", strLineColor)
'    Call SaveSetting(C_TITLE, "CrossLine", "LineWeight", strLineWeight)
'
'    Call SaveSetting(C_TITLE, "CrossLine", "FontColor", strFontColor)
    
    Call SaveSetting(C_TITLE, "CrossLine", "Guid", blnGuid)
'    Call SaveSetting(C_TITLE, "CrossLine", "Edit", blnEdit)
'    Call SaveSetting(C_TITLE, "CrossLine", "LineWidth", blnLineWidth)
'    Call SaveSetting(C_TITLE, "CrossLine", "BackColor", strFBackColor)
'
'    Call SaveSetting(C_TITLE, "CrossLine", "GuidTransparency", strGuidTransparency)
    
End Sub
Sub getCopyScreenSetting(ByRef blnFillVisible As Boolean, ByRef lngFillColor As Long, ByRef blnLine As Boolean)

    blnFillVisible = GetSetting(C_TITLE, "CopyScreen", "FillVisible", True)
    lngFillColor = CLng(GetSetting(C_TITLE, "CopyScreen", "FillColor", "&H00FFFFFF"))
    blnLine = GetSetting(C_TITLE, "CopyScreen", "Line", True)
    
End Sub
Sub setCopyScreenSetting(ByVal blnFillVisible As Boolean, ByVal strFillColor As String, ByVal blnLine As Boolean)

    Call SaveSetting(C_TITLE, "CopyScreen", "FillVisible", blnFillVisible)
    Call SaveSetting(C_TITLE, "CopyScreen", "FillColor", strFillColor)
    Call SaveSetting(C_TITLE, "CopyScreen", "Line", blnLine)
        
End Sub

