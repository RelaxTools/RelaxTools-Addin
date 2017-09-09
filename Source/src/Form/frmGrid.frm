VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrid 
   Caption         =   "かんたん表"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3870
   OleObjectBlob   =   "frmGrid.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmGrid"
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
Private mResult As VBA.VbMsgBoxResult

Public Function Start(ByRef lngHead As Long, ByRef lngCol As Long, ByRef lngHeadColor As Long, ByRef lngEvenColor As Long, ByRef blnHoganMode As Boolean) As VbMsgBoxResult

    mResult = vbCancel
    
    lblHead.BackColor = CLng(GetSetting(C_TITLE, "Grid", "HeadColor", "&H008080FF"))
    lblEven.BackColor = CLng(GetSetting(C_TITLE, "Grid", "EvenColor", "&H00C0C0FF"))
    
    txtHead.Text = lngHead
    txtCol.Text = lngCol
    chkHoganMode.Value = blnHoganMode

    Me.Show vbModal
    
    lngHead = Val(txtHead.Text)
    lngCol = Val(txtCol.Text)
    blnHoganMode = chkHoganMode.Value
    
    Select Case True
        Case OptionButton4
            lngHeadColor = 16764057
            lngEvenColor = -1
            
        Case OptionButton2
            lngHeadColor = 10079487
            lngEvenColor = 10092543
            
        Case OptionButton3
            lngHeadColor = lblHead.BackColor
            lngEvenColor = lblEven.BackColor
            
            Call SaveSetting(C_TITLE, "Grid", "HeadColor", "&H" & Right("00000000" & Hex(lngHeadColor), 8))
            Call SaveSetting(C_TITLE, "Grid", "EvenColor", "&H" & Right("00000000" & Hex(lngEvenColor), 8))
            
        Case Else
            lngHeadColor = 16764057
            lngEvenColor = 16777164
            
    End Select
    
    Start = mResult

End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    mResult = vbOK
    Unload Me
End Sub


Private Sub lblHead_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult
    
    
    lngColor = lblHead.BackColor
    
    Result = frmColor.Start(lngColor)
    
    If Result = vbOK Then
        lblHead.BackColor = lngColor
    End If
    

End Sub

Private Sub lblEven_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult
    
    
    lngColor = lblEven.BackColor
    
    Result = frmColor.Start(lngColor)
    
    If Result = vbOK Then
        lblEven.BackColor = lngColor
    End If
    
End Sub

Private Sub spnCol_SpinDown()
    txtCol.Text = spinDown(txtCol.Text)
End Sub

Private Sub spnCol_SpinUp()
    txtCol.Text = spinUp(txtCol.Text)
End Sub

Private Sub spnHead_SpinDown()
    txtHead.Text = spinDown(txtHead.Text)
End Sub

Private Sub spnHead_SpinUp()
    txtHead.Text = spinUp(txtHead.Text)
End Sub
Private Function spinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 1
    spinUp = lngValue

End Function

Private Function spinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 1
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDown = lngValue

End Function

