VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMarkdown 
   Caption         =   "表のMarkdown変換"
   ClientHeight    =   6828
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11568
   OleObjectBlob   =   "frmMarkdown.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMarkdown"
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
Const C_COLON As String = ":"

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Activate()
    
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText.Text)
End Sub

Private Sub UserForm_Initialize()
    Call MarkdownConv
End Sub
Private Sub MarkdownConv()

    Dim strBuf As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim r As Range

    For lngRow = 1 To Selection.Rows.Count
    
        strBuf = strBuf & C_SPLIT
        
        If lngRow = 2 Then
            For lngCol = 1 To Selection.Columns.Count
                Set r = Selection(1, lngCol)
                strBuf = strBuf & getHead(r) & C_SPLIT
            Next
            strBuf = strBuf & vbCrLf & C_SPLIT
        End If
        
        For lngCol = 1 To Selection.Columns.Count
            
            Set r = Selection(lngRow, lngCol)
            
            'マージセルの場合
            If r.MergeCells Then
            
                If r.MergeArea(1, 1).Address = r.Address Then
                    
                    strBuf = strBuf & getAttr(r) & C_SPLIT
                Else
                    
                    strBuf = strBuf & getAttr(r) & C_SPLIT
                
                End If
            
            Else
                    
                strBuf = strBuf & getAttr(r) & C_SPLIT
            
            End If
        
        Next
        
        strBuf = strBuf & vbCrLf
        
    Next
    
    txtText.Text = strBuf

End Sub
Private Function getHead(ByRef r As Range) As String
    
    Dim strLeft As String
    Dim strRight As String
    
    Select Case r.HorizontalAlignment
        Case xlLeft
            strLeft = ""
            strRight = ""
        Case xlRight
            strLeft = ""
            strRight = C_COLON
        Case xlCenter
            strLeft = C_COLON
            strRight = C_COLON
        Case Else
            Select Case True
                Case r.NumberFormatLocal = "@"
                    strLeft = ""
                    strRight = ""
                Case IsNumeric(r.Value)
                    strLeft = ""
                    strRight = C_COLON
                Case Else
                    strLeft = ""
                    strRight = ""
            End Select
    End Select
    
    getHead = strLeft & "---" & strRight
End Function
Private Function getAttr(ByRef r As Range) As String

    Dim strValue As String
    
    strValue = r.Text
            
    If VarType(r.Value) = vbString Then
    
        strValue = CharacterStyle(r)
    
    Else
        If r.Font.Italic Then
                strValue = "__" & strValue & "__"
        End If
        If r.Font.Bold Then
            strValue = "**" & strValue & "**"
        End If
    End If

    'Markdown変換でセル内改行を<br>に変換するようにした #56
    getAttr = Replace(Replace(strValue, vbCrLf, "<br>"), vbLf, "<br>")

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
    
    For i = 1 To r.Characters.Count
    
        blnStart = False
        blnEnd = False
        strTag = ""
            
        If r.Characters(i, 1).Font.Italic Then
            If blnItalic Then
            Else
                blnItalic = True
                strTag = strTag & "__"
                blnStart = True
            End If
        Else
            If blnItalic Then
                strTag = strTag & "__"
                blnItalic = False
                blnEnd = True
            End If
        End If

        If r.Characters(i, 1).Font.Bold Then
            If blnBold Then
            Else
                blnBold = True
                strTag = strTag & "**"
                blnStart = True
            End If
        Else
            If blnBold Then
                strTag = strTag & "**"
                blnBold = False
                blnEnd = True
            End If
        End If
        
        Select Case True
            Case blnStart
                strBuf = strBuf & strTag & r.Characters(i, 1).Text
            Case blnEnd
                strBuf = strBuf & strTag & r.Characters(i, 1).Text
            Case Else
                strBuf = strBuf & r.Characters(i, 1).Text
        End Select
    
    Next
    
    If blnItalic Then
        strBuf = strBuf & "__"
        blnItalic = False
    End If
    
    If blnBold Then
        strBuf = strBuf & "**"
        blnBold = False
    End If
    
    CharacterStyle = strBuf

End Function

