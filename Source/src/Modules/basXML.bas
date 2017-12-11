Attribute VB_Name = "basXML"
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

Public Function FormatXML(ByVal strBuf As String) As String

    Dim lngMax As Long
    Dim strChar As String
    Dim i As Long

    lngMax = Len(strBuf)

    Dim c As Collection
    Dim blnComment As Boolean
    Dim blnCData As Boolean
    Dim blnQuat As Boolean
    Dim s As XMLElement
    Dim strElement As String

    blnQuat = False
    blnComment = False
    blnCData = False

    Set c = New Collection

    strElement = ""


    '字句解析
    For i = 1 To lngMax

        strChar = Mid(strBuf, i, 1)

        Select Case True
            Case Mid(strBuf, i, 3) = "-->"
                blnComment = False
                strElement = strElement & strChar

            Case Mid(strBuf, i, 3) = "]]>"
                blnCData = False
                strElement = strElement & strChar

            Case strChar = """"
                blnQuat = Not (blnQuat)
                strElement = strElement & strChar

            Case Else
                If blnQuat Or blnComment Or blnCData Then
                    strElement = strElement & strChar
                Else
                    Select Case strChar
                        Case "<"
                            If IsDataEmpty(strElement) Then
                            Else
                                Set s = New XMLElement
                                s.Element = strElement
                                s.ElementType = data
                                s.ElementName = ""
                                c.Add s

                            End If

                            Select Case True
                                Case Mid(strBuf, i, 4) = "<!--"
                                    blnComment = True

                                Case Mid(strBuf, i, 9) = "<![CDATA["
                                    blnCData = True

                            End Select
                            strElement = strChar
                        Case ">"
                            strElement = strElement & strChar
                            Set s = New XMLElement

                            s.Element = strElement
                            Select Case True
                                Case Right(strElement, 3) = "-->"
                                    s.ElementType = Comment
                                    s.ElementName = TrimTag(strElement)

                                Case Right(strElement, 3) = "]]>"
                                    s.ElementType = CData
                                    s.ElementName = TrimTag(strElement)

                                Case Left(strElement, 2) = "<?"
                                    s.ElementType = Header
                                    s.ElementName = TrimTag(strElement)

                                Case Left(strElement, 2) = "</"
                                    s.ElementType = EndTag
                                    s.ElementName = TrimTag(strElement)

                                Case Right(strElement, 2) = "/>"
                                    s.ElementType = StandAlone
                                    s.ElementName = TrimTag(strElement)

                                Case Left(strElement, 1) = "<"
                                    s.ElementType = StartTag
                                    s.ElementName = TrimTag(strElement)

                                Case Else
                                    s.ElementType = data
                                    s.ElementName = ""

                            End Select
                            c.Add s
                            strElement = ""
                        Case Else
                            strElement = strElement & strChar
                    End Select
                End If
        End Select
    Next

    Dim strResult As String

    '構文解析

    lngMax = c.count

    Dim curElement As XMLElement
    Dim nextElement As XMLElement
    Dim strIndent As String
    Dim lngIndent As Long

    lngIndent = 0
    strResult = ""

    For i = 1 To lngMax

        Set curElement = c(i)
'        Cells(i, 1).Value = curElement.Element
'        Cells(i, 2).Value = curElement.ElementType
'        Cells(i, 3).Value = curElement.ElementName

        If i = lngMax Then
            Set nextElement = New XMLElement
        Else
            Set nextElement = c(i + 1)
        End If

        Select Case curElement.ElementType
            Case EnumElementType.Comment, EnumElementType.CData, EnumElementType.Header
                strResult = strResult & curElement.Element & vbCrLf

            Case EnumElementType.StandAlone
                strResult = strResult & GetIndentStr(lngIndent) & curElement.Element & vbCrLf

            Case EnumElementType.StartTag
                Select Case nextElement.ElementType
                    Case data
                        strResult = strResult & GetIndentStr(lngIndent) & curElement.Element
                    Case EndTag
                        If curElement.ElementName <> nextElement.ElementName Then
                            strResult = strResult & GetIndentStr(lngIndent) & curElement.Element & vbCrLf
                            lngIndent = lngIndent + 1
                        Else
                            strResult = strResult & GetIndentStr(lngIndent) & curElement.Element & nextElement.Element & vbCrLf
                            i = i + 1
                        End If
                    Case Else
                        strResult = strResult & GetIndentStr(lngIndent) & curElement.Element & vbCrLf
                        lngIndent = lngIndent + 1
                End Select

            Case EnumElementType.EndTag
                lngIndent = lngIndent - 1

                'XMLのミスなどでインデントがマイナスになるのを防ぐ
                If lngIndent < 0 Then
                    lngIndent = 0
                End If
                strResult = strResult & GetIndentStr(lngIndent) & curElement.Element & vbCrLf

            Case EnumElementType.data
                strResult = strResult & curElement.Element
                Select Case nextElement.ElementType
                    Case EndTag
                        strResult = strResult & nextElement.Element & vbCrLf
                        i = i + 1
                End Select
        End Select

    Next

    FormatXML = strResult

End Function
Private Function GetIndentStr(ByVal lngIndent As Long) As String

    Dim blnTab As Boolean
    Dim lngSeed As Long

    If CBool(GetSetting(C_TITLE, "XML", "Tab", False)) Then
        GetIndentStr = String$(lngIndent, vbTab)
    Else
        lngSeed = GetSetting(C_TITLE, "XML", "Seed", 2)
        GetIndentStr = Space(lngIndent * lngSeed)
    End If
End Function

Private Function IsDataEmpty(ByVal strBuf As String) As Boolean

    Dim i As Long
    Dim lngCnt As Long

    lngCnt = 0

    For i = 1 To Len(strBuf)
        Select Case Mid(strBuf, i, 1)
            Case " "
            Case vbTab
            Case vbCr
            Case vbLf
            Case Else
                lngCnt = lngCnt + 1
        End Select
    Next

    IsDataEmpty = (lngCnt <= 0)

End Function
Private Function TrimTag(ByVal strBuf As String) As String

    Dim i As Long
    Dim strResult As String
    Dim strChar As String

    strResult = ""

    Select Case True
        Case Left(strBuf, 2) = "<?"
            strBuf = Mid(strBuf, 3)
        Case Left(strBuf, 2) = "</"
            strBuf = Mid(strBuf, 3)
        Case Left(strBuf, 1) = "<"
            strBuf = Mid(strBuf, 2)
    End Select

    Select Case True
        Case Right(strBuf, 2) = "?>"
            strBuf = Mid(strBuf, 1, Len(strBuf) - 2)
        Case Right(strBuf, 2) = "/>"
            strBuf = Mid(strBuf, 1, Len(strBuf) - 2)
        Case Right(strBuf, 1) = ">"
            strBuf = Mid(strBuf, 1, Len(strBuf) - 1)
    End Select

    Dim lngPos As Long
    lngPos = InStr(strBuf, " ")
    If lngPos > 0 Then
        strResult = Mid(strBuf, 1, lngPos - 1)
    Else
        strResult = strBuf
    End If

    TrimTag = strResult

End Function



