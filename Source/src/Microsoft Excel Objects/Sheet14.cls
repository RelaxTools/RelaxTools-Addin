VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
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

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Const C_COMP_NO As Long = 1
    Const C_COMP_RESULT As Long = 2
    Const C_COMP_SRCSTR As Long = 3
    Const C_COMP_DSTSTR As Long = 4
    Const C_COMP_BOOK As Long = 5
    Const C_COMP_SHEET As Long = 6
    Const C_COMP_ADDRESS As Long = 7

    Dim Wb As Workbook
    Dim WS As Worksheet
    Dim strBook As String
    Dim strSheet As String
    Dim strAddress As String

    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If

    strBook = Cells(ActiveCell.row, C_COMP_BOOK).value
    If Len(strBook) = 0 Then
        Exit Sub
    End If
    strSheet = Cells(ActiveCell.row, C_COMP_SHEET).value
    If Len(strSheet) = 0 Then
        Exit Sub
    End If
    strAddress = Cells(ActiveCell.row, C_COMP_ADDRESS).value
    If Len(strAddress) = 0 Then
        Exit Sub
    End If

    On Error Resume Next

    Set Wb = Workbooks(strBook)
    Wb.Activate

    Set WS = Wb.Worksheets(strSheet)
    WS.Select
    
    WS.Range(strAddress).Select
    
End Sub


