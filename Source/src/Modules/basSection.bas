Attribute VB_Name = "basSection"
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

Public mColSection As New Collection
Public mColAllSection As Collection
Public Const C_FONT_DEFAULT As String = "ＭＳ ゴシック"

Public Sub createAllSectionObject()

    Set mColAllSection = New Collection

'    Dim i As Long
'    For i = 1 To mColAllSection.count
'        mColAllSection.Remove i
'    Next

    'すべてのセクション
    mColAllSection.Add New SecNumOne
    mColAllSection.Add New SecNumOneP
    mColAllSection.Add New SecNumPoint
    mColAllSection.Add New SecNumPoint2
    mColAllSection.Add New SecNumPoint3
    mColAllSection.Add New SecNumPointZen
    mColAllSection.Add New SecNumPoint2Zen
    mColAllSection.Add New SecNumPoint3Zen
    mColAllSection.Add New SecNum1S
    mColAllSection.Add New SecNum1E
    mColAllSection.Add New SecNum1K
    mColAllSection.Add New SecNum1
    mColAllSection.Add New SecNumA
    mColAllSection.Add New SecNumA2
    mColAllSection.Add New SecNumK
    mColAllSection.Add New SecNumK2
    mColAllSection.Add New SecNumK3
    mColAllSection.Add New SecNumK4
    mColAllSection.Add New SecNumK5
    mColAllSection.Add New SecNumK6
    mColAllSection.Add New SecNumI
    mColAllSection.Add New SecNumI2
    mColAllSection.Add New SecNumI4
    mColAllSection.Add New SecNumC

End Sub
Function rlxIsSectionNo(ByVal strBuf As String) As Boolean
Attribute rlxIsSectionNo.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxIsSectionNo.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim i As Long
    Dim lngCnt As Long
    Dim strSecNo As String
    Dim blnFind As Boolean
    
    blnFind = False

    lngCnt = Len(strBuf)
    If lngCnt = 0 Then
        Exit Function
    End If

    If mColAllSection Is Nothing Then
        Call createAllSectionObject
    End If

    For i = 1 To mColAllSection.count
    
        strSecNo = mColAllSection(i).SectionNumber(strBuf)
        If Len(strSecNo) > 0 Then
            blnFind = True
            Exit For
        End If
        
    Next
    
    rlxIsSectionNo = blnFind

End Function
Function rlxGetSectionNoAny(ByVal strBuf As String) As String
Attribute rlxGetSectionNoAny.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetSectionNoAny.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim i As Long
    Dim lngCnt As Long
    Dim strSecNo As String
    Dim blnFind As Boolean
    
    blnFind = False

    lngCnt = Len(strBuf)
    If lngCnt = 0 Then
        Exit Function
    End If
    
    strSecNo = ""

    If mColAllSection Is Nothing Then
        Call createAllSectionObject
    End If
    
    For i = 1 To mColAllSection.count
    
        strSecNo = mColAllSection(i).SectionNumber(strBuf)
        If Len(strSecNo) > 0 Then
            Exit For
        End If
        
    Next

    rlxGetSectionNoAny = strSecNo

End Function
Function rlxGetSectionNo(ByVal strBuf As String, ByVal lngIndentLevel As Long) As String
Attribute rlxGetSectionNo.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetSectionNo.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim obj As Object

    Set obj = rlxGetSectionObject(lngIndentLevel).classObj

    rlxGetSectionNo = obj.SectionNumber(strBuf)
    
    Set obj = Nothing
    
End Function

Function rlxHasSectionNo(ByVal strBuf As String, ByVal lngIndentLevel As Long) As Boolean
Attribute rlxHasSectionNo.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxHasSectionNo.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim i As Long
    Dim lngCnt As Long
    
    rlxHasSectionNo = False
    
    lngCnt = Len(strBuf)
    If lngCnt = 0 Then
        Exit Function
    End If
    
    If Len(rlxGetSectionNo(strBuf, lngIndentLevel)) > 0 Then
        rlxHasSectionNo = True
    End If
    
End Function

Function rlxGetSectionNext(ByVal strBuf As String, ByVal lngFromLevel As Long, ByVal lngIndentLevel As Long) As String
Attribute rlxGetSectionNext.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetSectionNext.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim obj As Object

    Set obj = rlxGetSectionObject(lngIndentLevel).classObj

    '次番号の取得
    rlxGetSectionNext = obj.NextNumber(strBuf, lngFromLevel, lngIndentLevel)
    
    Set obj = Nothing

End Function

Sub setSectionNo(ByRef r As Range, ByVal strNewNo As String)

    Dim lngPos As Long
    Dim obj As Object

    If VarType(r.Value) = vbString Then
        r.Characters(0, 0).Insert strNewNo
    Else
        r.Value = strNewNo & r.Value
    End If
    

    Set obj = rlxGetSectionObject(r.IndentLevel)

    'フォント有効の場合
    If obj.useFormat Then
        r.Font.Name = obj.fontName
        r.Font.Size = obj.fontSize
        r.Font.Bold = obj.fontBold
        r.Font.Italic = obj.fontItalic
        r.Font.Underline = obj.fontUnderLine
    End If

    Set obj = Nothing
    
End Sub

Sub delSectionNo(ByRef r As Range)

    Dim strSecNo As String
    Dim lngPos As Long
    Dim obj As Object

    '現在の段落番号を取得（レベルにかかわらない）
    strSecNo = rlxGetSectionNoAny(r.Value)
    If VarType(r.Value) = vbString Then
        If Len(strSecNo) > 0 Then
            r.Characters(1, Len(strSecNo)).Delete
        End If
    Else
        If Len(strSecNo) > 0 Then
            r.Value = Mid$(r.Value, Len(strSecNo) + 1)
        End If
    End If

    Set obj = rlxGetSectionObject(r.IndentLevel)

    'フォント有効の場合
    If obj.useFormat2 Then
        r.Font.Name = obj.fontName2
        r.Font.Size = obj.fontSize2
        r.Font.Bold = obj.fontBold2
        r.Font.Italic = obj.fontItalic2
        r.Font.Underline = obj.fontUnderLine2
    End If

    Set obj = Nothing

End Sub

Function rlxGetSectionObject(ByVal lngLevel As Long) As Object
Attribute rlxGetSectionObject.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetSectionObject.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim o As Object
    Dim key As String
    
    On Error Resume Next
    
    If mColSection Is Nothing Then
        Set mColSection = rlxInitSectionSetting()
    End If
    
    key = Format$((lngLevel Mod mColSection.count) + 1, "00")
    
    Set rlxGetSectionObject = mColSection(key)

End Function
Function rlxInitSectionSetting() As Collection
Attribute rlxInitSectionSetting.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxInitSectionSetting.VB_ProcData.VB_Invoke_Func = " \n19"

    On Error Resume Next

    Dim i As Long
    Dim j As Long
    Dim strPos As String
    Dim C_FONT_SIZE_DEFAULT As String
    Dim strClass As String
    
'    C_FONT_DEFAULT = Application.StandardFont
    C_FONT_SIZE_DEFAULT = Application.StandardFontSize
    
    strPos = GetSetting(C_TITLE, "Section", "pos", "x")
    If strPos = "x" Then
    
        strPos = "1"
    
        Dim Col As Collection
        Dim ss As SectionStructDTO
        
        For i = 1 To 6
        
            Set Col = New Collection
            
            For j = 1 To 10
            
                strClass = getDefault(Format$(i, "00") & Format$(j, "00"), 2)
                If strClass <> "" Then
                    Set ss = New SectionStructDTO
                    Set ss.classObj = rlxCreateSectionObject(strClass)
                    ss.useFormat = False
                    ss.fontName = C_FONT_DEFAULT
                    ss.fontSize = C_FONT_SIZE_DEFAULT
                    ss.fontBold = False
                    ss.fontItalic = False
                    ss.fontUnderLine = False
                    
                    ss.useFormat2 = False
                    ss.fontName2 = C_FONT_DEFAULT
                    ss.fontSize2 = C_FONT_SIZE_DEFAULT
                    ss.fontBold2 = False
                    ss.fontItalic2 = False
                    ss.fontUnderLine2 = False
                    
                    Col.Add ss, Format$(j, "00")
                    Set ss = Nothing
                End If
                
            Next
            
            setSectionSetting Format$(i, "00"), Col
            Set Col = Nothing
        
        Next
    End If
    
    Set rlxInitSectionSetting = rlxGetSectionSetting(Format$(Val(strPos), "00"))
    
End Function
Function rlxGetSectionSetting(ByVal strNo As String) As Collection
Attribute rlxGetSectionSetting.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetSectionSetting.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim strClass As String
    Dim ss As SectionStructDTO
    Dim Col As Collection
    Dim i As Long
    
'    Dim C_FONT_DEFAULT As String
    Dim C_FONT_SIZE_DEFAULT As String
'    C_FONT_DEFAULT = Application.StandardFont
    C_FONT_SIZE_DEFAULT = Application.StandardFontSize
    
    
    Set Col = New Collection
    
    i = 1
    Do While True
        strClass = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "class", "")
        If strClass = "" Then
            Exit Do
        Else
            Set ss = New SectionStructDTO
            Set ss.classObj = rlxCreateSectionObject(strClass)
            ss.useFormat = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "useFormat", False)
            ss.fontName = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontName", C_FONT_DEFAULT)
            ss.fontSize = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontSize", C_FONT_SIZE_DEFAULT)
            ss.fontBold = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontBold", False)
            ss.fontItalic = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontItalic", False)
            ss.fontUnderLine = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontUnderLine", False)
            
            ss.useFormat2 = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "useFormat2", False)
            ss.fontName2 = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontName2", C_FONT_DEFAULT)
            ss.fontSize2 = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontSize2", C_FONT_SIZE_DEFAULT)
            ss.fontBold2 = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontBold2", False)
            ss.fontItalic2 = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontItalic2", False)
            ss.fontUnderLine2 = GetSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontUnderLine2", False)
            
            Col.Add ss, Format$(i, "00")
            Set ss = Nothing
        End If
        i = i + 1
    Loop
    
    Set rlxGetSectionSetting = Col
    
    Set Col = Nothing
    
End Function
Sub setSectionSetting(ByVal strNo As String, ByRef Col As Collection)

    Dim i As Long
    On Error Resume Next
    For i = 1 To 99
        err.Clear
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "class")
        If err.Number <> 0 Then
            err.Clear
            On Error GoTo 0
            Exit For
        End If
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "useFormat")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontName")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontSize")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontBold")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontItalic")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontUnderLine")
        
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "useFormat2")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontName2")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontSize2")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontBold2")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontItalic2")
        Call DeleteSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontUnderLine2")
    Next
    
    For i = 1 To Col.count
        
        If Col(i).classObj Is Nothing Then
            Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "class", "")
        Else
            Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "class", Col(i).classObj.Class)
        End If
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "useFormat", Col(i).useFormat)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontName", Col(i).fontName)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontSize", Col(i).fontSize)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontBold", Col(i).fontBold)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontItalic", Col(i).fontItalic)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontUnderLine", Col(i).fontUnderLine)
        
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "useFormat2", Col(i).useFormat2)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontName2", Col(i).fontName2)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontSize2", Col(i).fontSize2)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontBold2", Col(i).fontBold2)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontItalic2", Col(i).fontItalic2)
        Call SaveSetting(C_TITLE, "Section", "Section" & strNo & Format$(i, "00") & "fontUnderLine2", Col(i).fontUnderLine2)
        
    Next
    
End Sub
Private Function getDefault(ByVal strBuf As String, ByVal lngCol As Long) As Variant

    Dim i As Long
    Dim strRet As Variant

    i = 2
    strRet = ""

    Do Until ThisWorkbook.Worksheets("Section").Cells(i, 1).Value = ""

        If UCase(strBuf) = ThisWorkbook.Worksheets("Section").Cells(i, 1).Value Then
            strRet = ThisWorkbook.Worksheets("Section").Cells(i, lngCol).Value
            If UCase(strRet) = "FALSE" Then
                strRet = False
            End If
            If UCase(strRet) = "TRUE" Then
                strRet = True
            End If
            Exit Do
        End If

        i = i + 1

    Loop

    getDefault = strRet

End Function
Function rlxCreateSectionObject(ByVal className As String) As Object
Attribute rlxCreateSectionObject.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxCreateSectionObject.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim ret As Object
    Dim obj As Object
    Set ret = Nothing
    
    If mColAllSection Is Nothing Then
        Call createAllSectionObject
    End If
    
    For Each obj In mColAllSection
    
        If className = obj.Class Then
            Set ret = obj
            Exit For
        End If
    
    Next

    Set rlxCreateSectionObject = ret

End Function
'--------------------------------------------------------------
'　目次作成(段落番号)
'--------------------------------------------------------------
Sub createContents()
    
    On Error GoTo ErrHandle

    frmContents.Show

    Exit Sub
ErrHandle:
    MsgBox "エラーが発生しました。", vbOKOnly, C_TITLE

End Sub
