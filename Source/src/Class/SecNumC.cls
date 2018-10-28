VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SecNumC"
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
Private Const C_SECTION_REGEXP As String = "^[①-⑳][ ]"
Private Const C_SECTION_SEPARATER As String = " "
Private Const C_SECTION_NAME As String = "①, ②, ③ ･･･ ⑳"
Private Const C_SECTION_CLASS As String = "SecNumC"
Public Property Get Class() As String
    Class = C_SECTION_CLASS
End Property

Public Function SectionLevelName(ByVal lngLevel As Long) As String
    
    SectionLevelName = " ① ────"

End Function
Public Property Get SectionName() As String
    SectionName = C_SECTION_NAME
End Property
Public Property Get SectionSeparator() As String
    SectionSeparator = C_SECTION_SEPARATER
End Property
Public Property Get SectionRegExp() As String
    SectionRegExp = C_SECTION_REGEXP
End Property
Public Property Get SectionNumber(ByVal v As String) As String

    Dim objRegx As Object
    Dim objMatch As Object
    
    Set objRegx = CreateObject("VBScript.RegExp")
        
    objRegx.Pattern = C_SECTION_REGEXP
    objRegx.IgnoreCase = False
    objRegx.Global = False

    Set objMatch = objRegx.Execute(v)
    
    If objMatch.Count > 0 Then
        SectionNumber = objMatch(0).Value
    Else
        SectionNumber = ""
    End If
    
    Set objMatch = Nothing
    Set objRegx = Nothing
    
End Property
Public Property Get NextNumber(ByVal v As String, ByVal lngFromLevel As Long, ByVal lngIndentLevel As Long) As String
       
    Dim i As Long
    Dim strNewNum As String
    Dim strNumber As String
    Dim lngAsc As Long
    
    strNumber = Trim(SectionNumber(v))
    
    If strNumber = "" Or lngFromLevel <> lngIndentLevel Then
        strNewNum = "①" & C_SECTION_SEPARATER
    Else
        lngAsc = Asc(strNumber) + 1
        If lngAsc > Asc("⑳") Then
            lngAsc = Asc("①")
        End If
        strNewNum = Chr(lngAsc) & C_SECTION_SEPARATER
    End If
    
    NextNumber = strNewNum

End Property




