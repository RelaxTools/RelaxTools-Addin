VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionAllMinus"
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


Private mRange As Range

'------------------------------------------------------------------------------------------------------------------------
' SelectionAllFrameWork 宣言
'------------------------------------------------------------------------------------------------------------------------
Private WithEvents SFWork As SelectionAllFrameWork
Attribute SFWork.VB_VarHelpID = -1

'------------------------------------------------------------------------------------------------------------------------
' SelectionAllFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set SFWork = New SelectionAllFrameWork
End Sub

'------------------------------------------------------------------------------------------------------------------------
' SelectionAllFrameWork 開放
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set SFWork = Nothing
End Sub

'------------------------------------------------------------------------------------------------------------------------
' SelectionAllFrameWork Run
'------------------------------------------------------------------------------------------------------------------------
Public Sub Run()
    SFWork.Run
End Sub

Private Sub SFWork_SelectionInit(Cancel As Boolean, Undo As Boolean)
    Undo = True
End Sub

Private Sub SFWork_SelectionMain(r As Range, Cancel As Boolean)

    Dim lngCnt As Long
    Dim lngRow As Long
    Dim lngCol As Long

    Dim strBuf As String
    Dim strHead As String
    Dim strBottom As String
    Dim strNum As String
    Dim lngPos As Long
    Dim lngLen As Long
    Dim i As Long
    
    Dim blnStart As Boolean
    Dim lngStartPos As Long
    Dim lngEndPos As Long

    On Error Resume Next

    lngCol = r.Column
    lngRow = r.Row
    
    strBuf = r.FormulaLocal
    
    If IsDate(strBuf) Then
        Cells(lngRow, lngCol).FormulaLocal = DateAdd("d", -1, CDate(strBuf))
    Else
        blnStart = False
        lngStartPos = 1
        lngEndPos = -1
        
        For i = Len(strBuf) To 1 Step -1
            Select Case Mid$(strBuf, i, 1)
                Case "0" To "9"
                    If blnStart Then
                    Else
                        blnStart = True
                        lngEndPos = i
                    End If
                Case Else
                    If blnStart Then
                        lngStartPos = i + 1
                        Exit For
                    End If
            End Select
        Next
        
        Select Case lngEndPos
            Case -1
            Case Else
                strHead = Mid$(strBuf, 1, lngStartPos - 1)
                strNum = Mid$(strBuf, lngStartPos, lngEndPos - lngStartPos + 1)
                strBottom = Mid$(strBuf, lngEndPos + 1, Len(strBuf) - lngEndPos)
                
        End Select
        
        Dim lngNo As Long
        
        If Val(strNum) - 1 < 0 Then
            strNum = String(Len(strNum), "9")
            Cells(lngRow, lngCol).FormulaLocal = strHead & Format$(Val(strNum), String(Len(strNum), "0")) & strBottom
        Else
            Cells(lngRow, lngCol).FormulaLocal = strHead & Format$(Val(strNum) - 1, String(Len(strNum), "0")) & strBottom
        End If
    End If
End Sub





