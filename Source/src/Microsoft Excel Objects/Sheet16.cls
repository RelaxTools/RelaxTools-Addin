VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
'Private Const C_TITLE As String = "Grep置換"
Private Const C_SEARCH_NO As Long = 1
Private Const C_SEARCH_BOOK As Long = 2

'Private Const C_WORD_FILE As String = ".DOC"
'Private Const C_EXCEL_FILE As String = ".XLS"
'Private Const C_PPT_FILE As String = ".PPT"


Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim WB As Workbook
    Dim WS As Worksheet
    Dim strBook As String

    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If

    strBook = Cells(ActiveCell.row, C_SEARCH_BOOK).Value
    If Len(strBook) = 0 Then
        Exit Sub
    End If

    On Error Resume Next
        

    Dim WSH As Object
    
    Set WSH = CreateObject("WScript.Shell")
    
    Call WSH.Run(strBook)
    
    Set WSH = Nothing

'    Select Case True
'        Case InStr(UCase(strBook), C_EXCEL_FILE) > 0
'
'            Dim XL As Excel.Application
'            Set XL = New Excel.Application
'            XL.visible = True
'
'            Call XL.Workbooks.Open(filename:=strBook)
'
'            Set XL = Nothing
'
'        Case InStr(UCase(strBook), C_WORD_FILE) > 0
'
'            Dim WD As Object
'            Set WD = CreateObject("Word.Application")
'            WD.visible = True
'
'            Call WD.Documents.Open(filename:=strBook)
'            AppActivate WD.Caption
'
'            Set WD = Nothing
'
'        Case InStr(UCase(strBook), C_PPT_FILE) > 0
'
'            Dim PP As Object
'
'            Set PP = CreateObject("PowerPoint.Application")
'            PP.visible
'
'            Call PP.Presentations.Open(filename:=strBook)
'
'            AppActivate PP.Caption
'
'            Set PP = Nothing
'
'    End Select
        
    
End Sub



