VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBinary 
   Caption         =   "セルのUNICODE表示"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655.001
   OleObjectBlob   =   "frmBinary.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmBinary"
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

Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1

Private objLabel() As Object

Private varBuf() As Variant


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub MW_WheelDown(obj As Object)
    Dim lngValue As Long
    
    On Error GoTo e

    lngValue = scrBar.Value + 3
    If lngValue > scrBar.Max Then
        scrBar.Value = scrBar.Max
    Else
        scrBar.Value = lngValue
    End If
    Disp
e:
End Sub

Private Sub MW_WheelUp(obj As Object)
    Dim lngValue As Long
    
    On Error GoTo e

    lngValue = scrBar.Value - 3
    If lngValue < scrBar.Min Then
        scrBar.Value = scrBar.Min
    Else
        scrBar.Value = lngValue
    End If
    Disp
e:
End Sub

Private Sub scrBar_Change()
    Disp
End Sub

Private Sub scrBar_Scroll()
    Disp
End Sub

'Private Sub UserForm_Activate()
'    If MW Is Nothing Then
'    Else
'        MW.Activate
'    End If
'End Sub

Private Sub UserForm_Initialize()
            
    Dim i As Long
    Dim j As Long
    Dim lbl As MSForms.Label
    Dim lngTop As Long
    Dim varLeft As Variant
    Dim varWidth As Variant
    
    varLeft = Array(lblHead01.Left, lblHead02.Left, lblHead03.Left, lblHead04.Left, lblHead05.Left, lblHead06.Left, lblHead07.Left, lblHead08.Left, lblHead09.Left, lblHead18.Left)
    varWidth = Array(lblHead01.width, lblHead02.width, lblHead03.width, lblHead04.width, lblHead05.width, lblHead06.width, lblHead07.width, lblHead08.width, lblHead09.width, lblHead18.width)
    
    lngTop = 4
    ReDim objLabel(1 To 16, 1 To 10)
    
    For i = 1 To 16
        
        lngTop = lngTop + 16
        
        For j = 1 To 10
        
            Set lbl = Controls.Add("Forms.Label.1", "Lavel" & Format(i, "00") & Format(j, "00"), False)
            
            lbl.AutoSize = False
'            lbl.Font.Charset = 128
'            lbl.Font.Name = "ＭＳ ゴシック"
'            lbl.Font.Size = 9
            lbl.WordWrap = False
            
            lbl.Top = lngTop
            lbl.Left = varLeft(j - 1)
            lbl.width = varWidth(j - 1)
'            lbl.Height = 16
            lbl.BackColor = &HFFFFFF
            lbl.SpecialEffect = fmSpecialEffectEtched
            
            If j <> 10 Then
                lbl.TextAlign = fmTextAlignCenter
            Else
                lbl.TextAlign = fmTextAlignLeft
            End If
            
'            lbl.visible = True
            Set objLabel(i, j) = lbl
            
            Set lbl = Nothing
            
        Next
    Next
    
    If Len(ActiveCell.Value) = 0 Then
        scrBar.enabled = False
        Exit Sub
    End If
    
    Dim bytBuf() As Byte
    Dim bytChr() As Byte

    bytBuf = ActiveCell.Value
    bytChr = ActiveCell.Value
    
    'バイトオーダの変更
    Dim bytSwap As Byte
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
        bytSwap = bytBuf(i)
        bytBuf(i) = bytBuf(i + 1)
        bytBuf(i + 1) = bytSwap
    Next
    
    
    
    
    j = 0
    Dim lngMax As Long
    Dim lngLine As Long
    
    lngMax = UBound(bytBuf) + 1
    lngLine = rlxRoundUp(lngMax / 16, 0)
    
    
    ReDim varBuf(1 To lngLine, 1 To 10)
    
    Dim bytStr() As Byte
    Dim k As Long
    Dim m As Long
    
    k = 0
    For i = 1 To lngLine
        For j = 1 To 10
        
            Select Case j
                Case 1
                    varBuf(i, j) = FixHex(k, 8)
                    ReDim bytStr(0 To 15)
                    m = 0
                Case 10
                    varBuf(i, j) = ReplaceStr(bytStr)
                Case Else
                    If UBound(bytBuf) < k Then
                        varBuf(i, j) = ""
                    Else
                        varBuf(i, j) = FixHex(bytBuf(k), 2) & FixHex(bytBuf(k + 1), 2)
                        bytStr(m) = bytChr(k)
                        m = m + 1
                        bytStr(m) = bytChr(k + 1)
                        m = m + 1
                        k = k + 2
                    End If
            End Select
  
        
        Next
    Next
    
    
    scrBar.Min = 1
    Dim d As Long
    d = lngLine
    If d <= 0 Then
        scrBar.Max = 1
    Else
        scrBar.Max = d
    End If
    scrBar.LargeChange = 16
    scrBar.SmallChange = 1
    Disp
    
    Set MW = basMouseWheel.GetInstance
    MW.Install Me
    Set MW.obj = scrBar

End Sub
Sub Disp()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = (scrBar.Value - 1)
    
    For i = 1 To 16
        For j = 1 To 10
            If UBound(varBuf, 1) < i + k Then
                objLabel(i, j).Caption = ""
                objLabel(i, j).visible = False
            Else
                objLabel(i, j).Caption = varBuf(i + k, j)
                objLabel(i, j).visible = True
            End If
        Next
    Next


End Sub
Function ReplaceStr(ByVal strBuf As String) As String

    ReplaceStr = Replace(Replace(strBuf, vbLf, "[LF]"), vbCr, "[CR]")

End Function
Sub bytCopy(ByRef bytSource() As Byte, ByRef bytDest() As Byte, ByVal lngPos As Long, ByVal lngsize As Long)

    Dim i As Long
    Dim j As Long
    
    i = lngPos
    j = 0
    
    ReDim bytDest(0 To lngsize - 1)
    
    Do Until lngPos + lngsize <= i
        If UBound(bytSource) < i Then
            Exit Do
        End If
        bytDest(j) = bytSource(i)
        i = i + 1
        j = j + 1
    Loop


End Sub
Function FixHex(ByVal lngAddress As Long, ByVal lngLen As Long) As String
    FixHex = Right$(String$(lngLen, "0") & Hex(lngAddress), lngLen)
End Function
'----------------------------------------------------------------------------------
'　文字列の左端から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
'----------------------------------------------------------------------------------
Private Function AscLeft(ByVal var As Variant, ByVal lngsize As Long) As String

    Dim lngLen As Long
    Dim i As Long
    
    Dim strChr As String
    Dim strResult As String
    
    lngLen = Len(var)
    strResult = ""

    For i = 1 To lngLen
    
        strChr = Mid(var, i, 1)
        strResult = strResult & strChr
        If rlxAscLen(strResult) > lngsize Then
            Exit For
        End If
    
    Next

    AscLeft = strResult

End Function

Private Sub UserForm_Terminate()
    If MW Is Nothing Then
    Else
        MW.UnInstall
        Set MW = Nothing
    End If
End Sub
