VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRedmine 
   Caption         =   "表のTextile変換"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   OleObjectBlob   =   "frmRedmine.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmRedmine"
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
Const C_SPLIT As String = "|"
Const C_HEAD As String = "_"

Const C_LEFT As String = "<"
Const C_RIGHT As String = ">"
Const C_CENTER As String = "="
Const C_TOP As String = "^"
Const C_BOTTOM As String = "~"

Const C_COLSPAN As String = "\0"
Const C_ROWSPAN As String = "/0"

Private blnReCall As Boolean

Private Sub chkHead_Change()
    Call TextileConv
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Call TextileConv
End Sub
Private Sub UserForm_Activate()
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText.Text)
End Sub
Private Sub TextileConv()

    Dim strBuf As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim r As Range

    For lngRow = 1 To Selection.Rows.count
    
        strBuf = strBuf & C_SPLIT
        
        For lngCol = 1 To Selection.Columns.count
            
            Set r = Selection(lngRow, lngCol)
            
            'マージセルの場合
            If r.MergeCells Then
            
                If r.MergeArea(1, 1).Address = r.Address Then

                    If lngRow = 1 And chkHead.value Then
                        strBuf = strBuf & C_HEAD
                    End If
                    
                    If r.MergeArea.Columns.count > 1 Then
                        strBuf = strBuf & "\" & r.MergeArea.Columns.count
                    End If
                    
                    If r.MergeArea.Rows.count > 1 Then
                        strBuf = strBuf & "/" & r.MergeArea.Rows.count
                    End If
                    
                    strBuf = strBuf & getAttr(r) & C_SPLIT
                
                End If
            
            Else
                
                If lngRow = 1 And chkHead.value Then
                    strBuf = strBuf & C_HEAD
                End If
                    
                strBuf = strBuf & getAttr(r) & C_SPLIT
            
            End If
        
        Next
        
        strBuf = strBuf & vbCrLf
        
    Next
    
    txtText.Text = strBuf

End Sub

Private Function getAttr(ByRef r As Range) As String

    Dim strH As String
    Dim strV As String
    Dim strValue As String

    Select Case r.HorizontalAlignment
        Case xlLeft
            strH = "" 'C_LEFT
        Case xlRight
            strH = C_RIGHT
        Case xlCenter
            strH = C_CENTER
        Case Else
            Select Case True
                Case r.NumberFormatLocal = "@"
                    strH = "" 'C_LEFT
                Case IsNumeric(r.value)
                    strH = C_RIGHT
                Case Else
                    strH = "" 'C_LEFT
            End Select
    End Select

    Select Case r.VerticalAlignment
        Case xlTop
            strV = C_TOP
        Case xlBottom
            strV = C_BOTTOM
        Case xlCenter
            strV = ""
        Case Else
    End Select
    
    strValue = r.value
            
    If VarType(r.value) = vbString Then
    
        If r.HasFormula Then
        Else
            strValue = CharacterStyle(r)
        End If
    Else
        Select Case True
            Case r.Font.Strikethrough
                strValue = "-" & strValue & "-"
            Case r.Font.Italic
                strValue = "_" & strValue & "_"
            Case r.Font.Underline <> xlUnderlineStyleNone
                strValue = "+" & strValue & "+"
            Case Else
        End Select
        If r.Font.Bold Then
            strValue = "*" & strValue & "*"
        End If
    End If

    getAttr = strH & strV & ". " & strValue

End Function
Function CharacterStyle(ByRef r As Range) As String

    Dim i As Long
    Dim blnBold As Boolean
    Dim blnStrike As Boolean
    Dim blnItalic As Boolean
    Dim blnUnder As Boolean
    
    Dim strBuf As String
    Dim strTag As String
    Dim blnStart As Boolean
    Dim blnEnd As Boolean
    
    For i = 1 To r.Characters.count
    
        blnStart = False
        blnEnd = False
        strTag = ""
        If r.Characters(i, 1).Font.Strikethrough Then
            If blnStrike Then
            Else
                blnStrike = True
                strTag = strTag & "-"
                blnStart = True
            End If
        Else
            If blnStrike Then
                strTag = strTag & "-"
                blnStrike = False
                blnEnd = True
            End If
        End If
            
        If r.Characters(i, 1).Font.Italic Then
            If blnItalic Then
            Else
                blnItalic = True
                strTag = strTag & "_"
                blnStart = True
            End If
        Else
            If blnItalic Then
                strTag = strTag & "_"
                blnItalic = False
                blnEnd = True
            End If
        End If
            
        If r.Characters(i, 1).Font.Underline <> xlUnderlineStyleNone Then
            If blnUnder Then
            Else
                blnUnder = True
                strTag = strTag & "+"
                blnStart = True
            End If
        Else
            If blnUnder Then
                strTag = strTag & "+"
                blnUnder = False
                blnEnd = True
            End If
        End If
            
        If r.Characters(i, 1).Font.Bold Then
            If blnBold Then
            Else
                blnBold = True
                strTag = strTag & "*"
                blnStart = True
            End If
        Else
            If blnBold Then
                strTag = strTag & "*"
                blnBold = False
                blnEnd = True
            End If
        End If
        
        Select Case True
            Case blnStart
                strBuf = strBuf & " " & strTag & r.Characters(i, 1).Text
            Case blnEnd
                strBuf = strBuf & strTag & " " & r.Characters(i, 1).Text
            Case Else
                strBuf = strBuf & r.Characters(i, 1).Text
        End Select
    
    Next
    
    If blnStrike Then
        strBuf = strBuf & "-"
        blnStrike = False
    End If
    
    If blnItalic Then
        strBuf = strBuf & "_"
        blnItalic = False
    End If
    
    If blnUnder Then
        strBuf = strBuf & "+"
        blnUnder = False
    End If
    
    If blnBold Then
        strBuf = strBuf & "*"
        blnBold = False
    End If
    
    CharacterStyle = strBuf

End Function
