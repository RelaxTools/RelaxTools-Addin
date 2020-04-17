VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHtml 
   Caption         =   "表のHTML変換"
   ClientHeight    =   6930
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11568
   OleObjectBlob   =   "frmHtml.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmHtml"
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
Const C_HTML_START As String = "<html>"
Const C_HTML_END As String = "</html>"
Const C_BODY_START As String = "<body>"
Const C_BODY_END As String = "</body>"
Const C_TABLE_START As String = "<table>"
Const C_TABLE_BORDER_START As String = "<table border=""1"">"
Const C_TABLE_END As String = "</table>"
Const C_TR_START As String = "  <tr>"
Const C_TR_END As String = "  </tr>"
Const C_TD_START_FROM As String = "    <td"
Const C_TD_START_TO As String = ">"
Const C_TD_END As String = "</td>"
Const C_TH_START_FROM As String = "    <th"
Const C_TH_START_TO As String = ">"
Const C_TH_END As String = "</th>"
Const C_BR As String = "<br>"

Private blnReCall As Boolean
Private Sub UserForm_Activate()
    txtHtml.SelStart = 0
    txtHtml.SelLength = Len(txtHtml.Text)
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()

    On Error Resume Next

    Dim strTmpHtml As String
    Dim FS As Object
    Const C_FILE_NAME As String = "preview.html"


    Set FS = CreateObject("Scripting.FileSystemObject")

    strTmpHtml = rlxGetTempFolder() & C_FILE_NAME
    
    Set FS = Nothing


    Dim fp As Integer

    fp = FreeFile()
    Open strTmpHtml For Output As fp
    Close fp

    fp = FreeFile()
    Open strTmpHtml For Binary As fp
    
    Put #fp, , txtHtml.Text

    Close fp

    Dim WSH As Object
    
    Set WSH = CreateObject("WScript.Shell")
    
    WSH.Run strTmpHtml
    
    Set WSH = Nothing
    
    

End Sub

Private Sub optColor1_Click()
    Call htmlConv(True)
End Sub

Private Sub optColor2_Click()
    Call htmlConv(True)
End Sub

Private Sub optLine1_Click()
    Call htmlConv(True)
End Sub

Private Sub optLine2_Click()
    Call htmlConv(True)
End Sub

Private Sub optTag1_Click()
    Call htmlConv(True)
End Sub

Private Sub optTag2_Click()
    Call htmlConv(True)
End Sub

Private Sub optWidth1_Click()
    Call htmlConv(True)
End Sub

Private Sub optWidth2_Click()
    Call htmlConv(True)
End Sub

Private Sub UserForm_Initialize()


    blnReCall = True

    optLine1.Value = True
    optTag1.Value = True
    optWidth1.Value = True
    optColor1.Value = True

    blnReCall = False

    Call htmlConv(False)

End Sub
Private Function AddTag(ByRef strBuf As String, ByVal strTag As String) As String

    AddTag = strBuf & strTag & vbCrLf

End Function

Private Sub htmlConv(ByVal flg As Variant)

    Dim strBuf As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim r As Range
    Dim lngPoint() As Long
    Dim lngWidth() As Long
    Dim lngTotal As Long
    Dim lngTotal2 As Long
    
    Dim lngMax As Long
    Dim lngMaxPos As Long
    Dim lngMin As Long
    Dim lngMinPos As Long
    
    Dim lngSum As Long
    Dim lngPos As Long
    
    Dim strTdStartFrom As String
    Dim strTdStartTo As String
    Dim strTdEnd As String
    
    If blnReCall Then
        Exit Sub
    End If
    
'    If flg Then
'        Me.Hide
'        frmInformation.Message = "HTML変換中です。しばらくお待ちください。"
'        frmInformation.Show
'        DoEvents
'    End If

    Dim blnExLine As Boolean
    Dim blnExTag As Boolean
    Dim blnExWidth As Boolean
    Dim blnExColor As Boolean
    
    ReDim lngPoint(1 To Selection.Columns.Count)
    ReDim lngWidth(1 To Selection.Columns.Count)
    
    blnExLine = optLine1.Value
    blnExTag = optTag1.Value
    blnExWidth = optWidth2.Value
    blnExColor = optColor2.Value
    
    lngTotal = 0
    For lngCol = 1 To Selection.Columns.Count
        lngPoint(lngCol) = Selection.Columns(lngCol).width
        lngTotal = lngTotal + lngPoint(lngCol)
    Next

    lngTotal2 = 0
    lngMax = 0
    lngMaxPos = 0
    lngMin = 101
    lngMinPos = 0
    For lngCol = 1 To Selection.Columns.Count
        lngWidth(lngCol) = Fix((lngPoint(lngCol) / lngTotal) * 100)
        lngTotal2 = lngTotal2 + lngWidth(lngCol)
        If lngMin > lngWidth(lngCol) Then
            lngMin = lngWidth(lngCol)
            lngMinPos = lngCol
        End If
        If lngMax < lngWidth(lngCol) Then
            lngMax = lngWidth(lngCol)
            lngMaxPos = lngCol
        End If
    Next

    Dim lngMinCnt As Long
    Dim lngMaxCnt As Long
    
    lngMaxCnt = 0
    lngMinCnt = 0
    
    For lngCol = 1 To Selection.Columns.Count
        If lngWidth(lngMinPos) = lngWidth(lngCol) Then
            lngMinCnt = lngMinCnt + 1
        End If
        If lngWidth(lngMaxPos) = lngWidth(lngCol) Then
            lngMaxCnt = lngMaxCnt + 1
        End If
    Next

    If lngMinCnt < lngMaxCnt Then
        lngWidth(lngMinPos) = lngWidth(lngMinPos) + (100 - lngTotal2)
    Else
        lngWidth(lngMaxPos) = lngWidth(lngMaxPos) + (100 - lngTotal2)
    End If

    strBuf = AddTag(strBuf, C_HTML_START)
    strBuf = AddTag(strBuf, C_BODY_START)
    If blnExLine Then
        strBuf = AddTag(strBuf, C_TABLE_BORDER_START)
    Else
        strBuf = AddTag(strBuf, C_TABLE_START)
    End If
    
    If blnExTag Then
        strTdStartFrom = C_TD_START_FROM
        strTdStartTo = C_TD_START_TO
        strTdEnd = C_TD_END
    Else
        strTdStartFrom = C_TH_START_FROM
        strTdStartTo = C_TH_START_TO
        strTdEnd = C_TH_END
    End If
    
    For lngRow = 1 To Selection.Rows.Count
    
        strBuf = AddTag(strBuf, C_TR_START)
    
        For lngCol = 1 To Selection.Columns.Count
            
            Set r = Selection(lngRow, lngCol)
            
            'マージセルの場合
            If r.MergeCells Then
            
                If r.MergeArea(1, 1).Address = r.Address Then
                
                    If lngRow = 1 Then
                    
                        strBuf = strBuf & strTdStartFrom
                        
                        lngSum = 0
                        For lngPos = lngCol To lngCol + r.MergeArea.Columns.Count - 1
                            lngSum = lngSum + lngWidth(lngPos)
                        Next
                        If blnExWidth Then
                            strBuf = strBuf & " width=""" & lngSum & "%"""
                        End If
                    Else
                        strBuf = strBuf & C_TD_START_FROM
                    End If
                    
                    strBuf = strBuf & " style="""
                    If blnExColor Then
#If VBA7 Then
                            strBuf = strBuf & "background-color:" & getHtmlRGB(r.DisplayFormat.Interior.Color) & ";"
#Else
                            strBuf = strBuf & "background-color:" & getHtmlRGB(r.Interior.Color) & ";"
#End If
                    End If
                    
'                    strBuf = strBuf & " style=""text-align:" & getAlign(r) & ";"""
                    strBuf = strBuf & "text-align:" & getAlign(r) & ";"""
                    
                    If r.MergeArea.Columns.Count > 1 Then
                        strBuf = strBuf & " colspan=""" & r.MergeArea.Columns.Count & """"
                    End If
                    
                    If r.MergeArea.Rows.Count > 1 Then
                        strBuf = strBuf & " rowspan=""" & r.MergeArea.Rows.Count & """"
                    End If
                    
                    If lngRow = 1 Then
                        strBuf = strBuf & strTdStartTo & getText(blnExColor, r) & strTdEnd & vbCrLf
                    Else
                        strBuf = strBuf & C_TD_START_TO & getText(blnExColor, r) & C_TD_END & vbCrLf
                    End If
                    
                End If
            
            Else
                If lngRow = 1 Then
                    strBuf = strBuf & strTdStartFrom
                
                    If blnExWidth Then
                        strBuf = strBuf & " width=""" & lngWidth(lngCol) & "%"""
                    End If
                
                    strBuf = strBuf & " style="""
                    If blnExColor Then
#If VBA7 Then
                            strBuf = strBuf & "background-color:" & getHtmlRGB(r.DisplayFormat.Interior.Color) & ";"
#Else
                            strBuf = strBuf & "background-color:" & getHtmlRGB(r.Interior.Color) & ";"
#End If
                    End If
                    
                    strBuf = strBuf & "text-align:" & getAlign(r) & ";"""
                    strBuf = strBuf & strTdStartTo & getText(blnExColor, r) & strTdEnd & vbCrLf
                
                Else
                
                    strBuf = strBuf & C_TD_START_FROM
                    
                    strBuf = strBuf & " style="""
                    If blnExColor Then
#If VBA7 Then
                            strBuf = strBuf & "background-color:" & getHtmlRGB(r.DisplayFormat.Interior.Color) & ";"
#Else
                            strBuf = strBuf & "background-color:" & getHtmlRGB(r.Interior.Color) & ";"
#End If
                    End If
                    
                    strBuf = strBuf & "text-align:" & getAlign(r) & ";"""
                    strBuf = strBuf & C_TD_START_TO & getText(blnExColor, r) & C_TD_END & vbCrLf
                End If
            End If
        
        Next
        
        strBuf = AddTag(strBuf, C_TR_END)
        
    Next

    strBuf = AddTag(strBuf, C_TABLE_END)
    strBuf = AddTag(strBuf, C_BODY_END)
    strBuf = AddTag(strBuf, C_HTML_END)
    
    txtHtml.Text = strBuf
    
    txtHtml.SelStart = Len(txtHtml.Text)
    txtHtml.SelStart = 0
    txtHtml.SelLength = Len(txtHtml.Text)

'If flg Then
'    Unload frmInformation
'    Me.Show
'End If
    
    txtHtml.SetFocus
'    SendKeys "^A"
    
End Sub
Private Function getText(ByVal blnExColor As Boolean, ByRef r As Range) As String


    Dim lngColor As Variant
    Dim i As Long
    Dim blnBold As Boolean
    Dim strBuf As String
    
    blnBold = False
    
    If blnExColor Then
''        If VarType(r.Value) = vbString Then
'        On Error Resume Next
'        i = r.Characters.count
'        If err.Number = 0 Then
''            For i = 1 To Len(r.Value)
'            For i = 1 To r.Characters.count
'                '<span>
'                Select Case i
'                    Case 1
'#If VBA7 Then
'                            strBuf = "<span style=""color:" & getRGB(r.DisplayFormat.Characters(i, 1).Font.Color) & """ >"
'                            lngColor = r.DisplayFormat.Characters(i, 1).Font.Color
'#Else
'                            strBuf = "<span style=""color:" & getRGB(r.Characters(i, 1).Font.Color) & """ >"
'                            lngColor = r.Characters(i, 1).Font.Color
'#End If
'                    Case Else
'#If VBA7 Then
'                            If lngColor <> r.DisplayFormat.Characters(i, 1).Font.Color Then
'                                strBuf = strBuf & "</span><span style=""color:" & getRGB(r.DisplayFormat.Characters(i, 1).Font.Color) & """ >"
'                                lngColor = r.DisplayFormat.Characters(i, 1).Font.Color
'                            End If
'#Else
'                            If lngColor <> r.Characters(i, 1).Font.Color Then
'                                strBuf = strBuf & "</span><span style=""color:" & getRGB(r.Characters(i, 1).Font.Color) & """ >"
'                                lngColor = r.Characters(i, 1).Font.Color
'                            End If
'#End If
'
'                End Select
'
'
'                strBuf = strBuf & htmlSanitizing(r.Characters(i, 1).Text)
''                strBuf = strBuf & htmlSanitizing(Mid$(r.Value, i, 1))
'
'                '</span>
'                Select Case i
'                    Case r.Characters.count
''                    Case Len(r.Value)
'                        strBuf = strBuf & "</span>"
'                End Select
'
'            Next
'        Else
#If VBA7 Then
                strBuf = "<span style=""color:" & getHtmlRGB(r.DisplayFormat.Font.Color) & """>" & rlxHtmlSanitizing(r.Text) & "</span>"
#Else
                strBuf = "<span style=""color:" & getHtmlRGB(r.Font.Color) & """>" & rlxHtmlSanitizing(r.Text) & "</span>"
#End If
'        End If
    Else
        strBuf = rlxHtmlSanitizing(r.Text)
    End If
    
    getText = replaceNl(strBuf)

End Function
Private Function replaceNl(ByVal strBuf As String) As String

    strBuf = Replace(strBuf, vbCrLf, C_BR)
    strBuf = Replace(strBuf, vbCr, C_BR)
    strBuf = Replace(strBuf, vbLf, C_BR)
    
    replaceNl = strBuf
    
End Function

Private Function getAlign(ByRef r As Range) As String

    Dim strBuf As String

    Select Case r.HorizontalAlignment
        Case xlLeft
            strBuf = "left"
        Case xlRight
            strBuf = "right"
        Case xlCenter
            strBuf = "center"
        Case Else
            Select Case True
                Case r.NumberFormatLocal = "@"
                    strBuf = "left"
                Case IsNumeric(r.Value)
                    strBuf = "right"
                Case Else
                    strBuf = "left"
            End Select
    End Select

    getAlign = strBuf

End Function
