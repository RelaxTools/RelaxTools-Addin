Attribute VB_Name = "basStampSakura"
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
'--------------------------------------------------------------
'　付箋貼り付け
'--------------------------------------------------------------
Sub pasteSakura(ByVal strId As String, ByVal Index As Long)

    Dim r As Shape
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    Set r = ThisWorkbook.Worksheets("sakura").Shapes("picSakura" & Format(Index, "00"))

    r.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Call CopyClipboardSleep
 
    ActiveSheet.Paste

    Selection.ShapeRange.Width = 25 * C_RASIO
    Selection.ShapeRange.Height = 25 * C_RASIO

    Application.ScreenUpdating = True

End Sub
'--------------------------------------------------------------
'　イメージファイル作成
'--------------------------------------------------------------
Function getImageSakura(ByVal strId As String, ByVal Index As Long) As StdPicture
    
    Set getImageSakura = Nothing
    
    On Error Resume Next
    
    Dim r As Shape
    Set r = ThisWorkbook.Worksheets("sakura").Shapes("picSakura" & Format(Index, "00"))
    
    Dim b As Shape
    Dim o As Object
    
    Set b = ThisWorkbook.Worksheets("sakura").Shapes("shpBack")
    
    b.Top = r.Top
    b.Left = r.Left
    b.Height = r.Width
    b.Width = r.Width
    
    b.ZOrder msoSendToBack
    
    Set o = ThisWorkbook.Worksheets("sakura").Shapes.Range(Array(r.name, b.name)).Group
    
    Set getImageSakura = CreatePictureFromClipboard(o)
    
    o.Ungroup
    
    
End Function
