VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSectionEx 
   Caption         =   "段落番号の設定"
   ClientHeight    =   9825.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   OleObjectBlob   =   "frmSectionEx.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSectionEx"
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
Private Const C_NO As Long = 0
Private Const C_SECTION_NAME As Long = 1
Private Const C_CLASS_NAME As Long = 2

Private Const C_LEVEL As Long = 0
Private Const C_SECTION As Long = 1
Private Const C_ENABLE As Long = 2
Private Const C_FONT_NAME As Long = 3
Private Const C_FONT_SIZE As Long = 4
Private Const C_FONT_BOLD As Long = 5
Private Const C_FONT_ITALIC As Long = 6
Private Const C_FONT_UNDER_LINE As Long = 7
Private Const C_CLASS As Long = 8

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2

Private Const C_TRUE As String = "○"
Private Const C_FALSE As String = "－"

Private mRet As VbMsgBoxResult
Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1

Private Sub cmdAdd_Click()

    Dim j As Long
    Dim i As Long
    Dim blnFind As Boolean
    Dim strKey As String
     
    blnFind = False
    
    If lstSelect.ListCount = 0 Then
        Exit Sub
    End If
    

    lstSetting.AddItem ""
    j = lstSetting.ListCount - 1
    
    lstSetting.List(j, C_LEVEL) = j + 1
    lstSetting.List(j, C_SECTION) = lstSelect.List(lstSelect.ListIndex, C_SECTION_NAME)
    lstSetting.List(j, C_ENABLE) = C_FALSE
    lstSetting.List(j, C_FONT_NAME) = C_FONT_DEFAULT
    lstSetting.List(j, C_FONT_SIZE) = Application.StandardFontSize
    lstSetting.List(j, C_FONT_BOLD) = C_FALSE
    lstSetting.List(j, C_FONT_ITALIC) = C_FALSE
    lstSetting.List(j, C_FONT_UNDER_LINE) = C_FALSE
    lstSetting.List(j, C_CLASS) = lstSelect.List(lstSelect.ListIndex, C_CLASS_NAME)
    
    lstSetting.ListIndex = j
    lstSetting.TopIndex = j
    
End Sub

Private Sub cmdDel_Click()

    Dim i As Long
    Dim lngLast As Long
    
    If lstSetting.ListCount <= 1 Then
        MsgBox "すべての段落番号を削除することはできません。", vbExclamation + vbOKOnly, C_TITLE
        Exit Sub
    End If
    
    lngLast = lstSetting.ListIndex
    
    If lngLast > -1 Then
        lstSetting.RemoveItem lngLast
    End If
    
    '番号を振りなおす
    For i = 0 To lstSetting.ListCount - 1
        lstSetting.List(i, C_LEVEL) = i + 1
    Next
    
End Sub





Private Sub cboFont_Change()
    If lstSetting.ListIndex > -1 Then
        lstSetting.List(lstSetting.ListIndex, C_FONT_NAME) = cboFont.Text
    End If
    Call previewLabel
End Sub

Private Sub chkFontBold01_Click()

End Sub

Private Sub CheckBox4_Click()

End Sub

Private Sub chkFontBold_Click()
    If lstSetting.ListIndex > -1 Then
        lstSetting.List(lstSetting.ListIndex, C_FONT_BOLD) = chgBoolToStr(chkFontBold.Value)
    End If
    Call previewLabel
End Sub

Private Sub chkFontItalic_Click()
    If lstSetting.ListIndex > -1 Then
        lstSetting.List(lstSetting.ListIndex, C_FONT_ITALIC) = chgBoolToStr(chkFontItalic.Value)
    End If
    Call previewLabel
End Sub

Private Sub chkFontUnderLine_Click()
    If lstSetting.ListIndex > -1 Then
        lstSetting.List(lstSetting.ListIndex, C_FONT_UNDER_LINE) = chgBoolToStr(chkFontUnderLine.Value)
    End If
    Call previewLabel
End Sub

Private Sub chkUseFormat_Click()

    Dim blnValue As Boolean
    
    blnValue = chkUseFormat.Value
    
    cboFont.enabled = blnValue
    txtFontSize.enabled = blnValue
    chkFontBold.enabled = blnValue
    chkFontItalic.enabled = blnValue
    chkFontUnderLine.enabled = blnValue
    If lstSetting.ListIndex > -1 Then
        lstSetting.List(lstSetting.ListIndex, C_ENABLE) = chgBoolToStr(blnValue)
    End If
    Call previewLabel
End Sub

Private Sub chkUseFormat2_Click()

    Dim blnValue As Boolean
    
    blnValue = chkUseFormat2.Value
    
    cboFont2.enabled = blnValue
    txtFontSize2.enabled = blnValue
    chkFontBold2.enabled = blnValue
    chkFontItalic2.enabled = blnValue
    chkFontUnderLine2.enabled = blnValue

End Sub

Private Sub cmdCancel_Click()
    mRet = vbCancel
    Unload Me
End Sub



Private Sub cmdOk_Click()
    mRet = vbOK
    Unload Me
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 選択行を上に移動
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdUp_Click()
     Call moveList(C_UP)
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 選択行を下に移動
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdDown_Click()
     Call moveList(C_DOWN)
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 移動処理
'------------------------------------------------------------------------------------------------------------------------
Private Sub moveList(ByVal lngMode As Long)

    Dim lngCnt As Long
    Dim lngCmp As Long
    
    Dim varTmp As Variant

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngInc As Long

    '１つなら不要
    If lstSetting.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstSetting.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstSetting.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select

    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstSetting.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If
            
            lngCmp = lngCnt + lngInc * -1
            
            Dim i As Long
            For i = C_SECTION To C_CLASS
                varTmp = lstSetting.List(lngCnt, i)
                lstSetting.List(lngCnt, i) = lstSetting.List(lngCmp, i)
                lstSetting.List(lngCmp, i) = varTmp
            Next
            
            lstSetting.Selected(lngCnt) = False
            lstSetting.Selected(lngCnt + lngInc * -1) = True
        End If
    
    Next

End Sub



Private Sub CommandButton1_Click()

End Sub

Sub previewLabel()

    Dim obj As Object
    Dim strBuf As String
    
    If lstSetting.ListIndex = -1 Then
        Exit Sub
    End If
    
    Set obj = rlxCreateSectionObject(lstSetting.List(lstSetting.ListIndex, C_CLASS))

    strBuf = obj.SectionLevelName(1) & vbCrLf
    'strbuf = strbuf & Space(2) & obj.SectionLevelName(2) & vbCrLf

    lblPreview.Caption = strBuf
    If chkUseFormat.Value Then
        lblPreview.Font = cboFont.Text
        lblPreview.fontSize = txtFontSize.Text
        lblPreview.Font.Bold = chkFontBold.Value
        lblPreview.Font.Italic = chkFontItalic.Value
        lblPreview.Font.Underline = chkFontUnderLine.Value
    Else
        lblPreview.Font = C_FONT_DEFAULT
        lblPreview.fontSize = Application.StandardFontSize
        lblPreview.Font.Bold = False
        lblPreview.Font.Italic = False
        lblPreview.Font.Underline = False
    End If
    
    Set obj = Nothing

End Sub

Private Sub lstSelect_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstSelect
End Sub

Private Sub lstSetting_Change()
    Call previewLabel
End Sub

Private Sub lstSetting_Click()


    If lstSetting.List(lstSetting.ListIndex, C_ENABLE) = C_TRUE Then
        chkUseFormat.Value = True
    Else
        chkUseFormat.Value = False
    End If
    
    cboFont.Text = lstSetting.List(lstSetting.ListIndex, C_FONT_NAME)
    txtFontSize.Text = lstSetting.List(lstSetting.ListIndex, C_FONT_SIZE)
    
    If lstSetting.List(lstSetting.ListIndex, C_FONT_BOLD) = C_TRUE Then
        chkFontBold.Value = True
    Else
        chkFontBold.Value = False
    End If
    
    If lstSetting.List(lstSetting.ListIndex, C_FONT_ITALIC) = C_TRUE Then
        chkFontItalic.Value = True
    Else
        chkFontItalic.Value = False
    End If
    
    If lstSetting.List(lstSetting.ListIndex, C_FONT_UNDER_LINE) = C_TRUE Then
        chkFontUnderLine.Value = True
    Else
        chkFontUnderLine.Value = False
    End If
            
End Sub

Private Sub lstSetting_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstSetting
End Sub

Private Sub txtFontSize_Change()
    Call previewLabel
    If lstSetting.ListIndex > -1 Then
        lstSetting.List(lstSetting.ListIndex, C_FONT_SIZE) = txtFontSize.Text
    End If
End Sub

Private Sub UserForm_Initialize()

    Dim i As Long
    Dim objAll As Object

    If mColAllSection Is Nothing Then
        Call createAllSectionObject
    End If
    
    i = 1
    For Each objAll In mColAllSection
        lstSelect.AddItem ""
        lstSelect.List(lstSelect.ListCount - 1, C_NO) = i
        lstSelect.List(lstSelect.ListCount - 1, C_SECTION_NAME) = objAll.SectionName
        lstSelect.List(lstSelect.ListCount - 1, C_CLASS_NAME) = objAll.Class
        i = i + 1
    Next
    lstSelect.ListIndex = 0
 
    With Application.CommandBars("Formatting").Controls(1)
        For i = 1 To .ListCount
            cboFont.AddItem .List(i)
            cboFont2.AddItem .List(i)
        Next i
    End With
    
    For i = 0 To cboFont.ListCount - 1
        If cboFont.List(i) = C_FONT_DEFAULT Then
            cboFont.ListIndex = i
        End If
    Next
    For i = 0 To cboFont2.ListCount - 1
        If cboFont2.List(i) = C_FONT_DEFAULT Then
            cboFont2.ListIndex = i
        End If
    Next
    
    txtFontSize.Text = Application.StandardFontSize
    txtFontSize2.Text = Application.StandardFontSize
    
    chkUseFormat_Click
    chkUseFormat2_Click
    
    Set MW = New MouseWheel
    MW.Install Me

End Sub
Private Function chgBoolToStr(ByVal blnFlg As Boolean) As String
    If blnFlg Then
        chgBoolToStr = C_TRUE
    Else
        chgBoolToStr = C_FALSE
    End If
End Function
'Private Sub UserForm_Activate()
'    MW.Activate
'End Sub
Public Function Start(ByRef col As Collection) As Collection
    
    Dim i As Long
    Dim j As Long
    On Error GoTo 0
    
    For i = 1 To col.Count
            
        lstSetting.AddItem ""
        lstSetting.List(lstSetting.ListCount - 1, C_LEVEL) = i
        lstSetting.List(lstSetting.ListCount - 1, C_SECTION) = col(i).classObj.SectionName
        lstSetting.List(lstSetting.ListCount - 1, C_ENABLE) = chgBoolToStr(col(i).useFormat)
        lstSetting.List(lstSetting.ListCount - 1, C_FONT_NAME) = col(i).fontName
        lstSetting.List(lstSetting.ListCount - 1, C_FONT_SIZE) = col(i).fontSize
        lstSetting.List(lstSetting.ListCount - 1, C_FONT_BOLD) = chgBoolToStr(col(i).fontBold)
        lstSetting.List(lstSetting.ListCount - 1, C_FONT_ITALIC) = chgBoolToStr(col(i).fontItalic)
        lstSetting.List(lstSetting.ListCount - 1, C_FONT_UNDER_LINE) = chgBoolToStr(col(i).fontUnderLine)
        lstSetting.List(lstSetting.ListCount - 1, C_CLASS) = col(i).classObj.Class
    
    Next
    If lstSetting.ListCount > 0 Then
        lstSetting.Selected(0) = True
        
        chkUseFormat2.Value = col(1).useFormat2
            
        cboFont2.Text = col(1).fontName2
        txtFontSize2.Text = col(1).fontSize2
        
        chkFontBold2.Value = col(1).fontBold2
        chkFontItalic2.Value = col(1).fontItalic2
        chkFontUnderLine2.Value = col(1).fontUnderLine2
        
    End If

    Me.Show
    
    If mRet <> vbOK Then
        Set Start = Nothing
        Exit Function
    End If

    Dim strClass As String
    Dim ret As Collection
    
    Set ret = New Collection
    Dim ss As SectionStructDTO
    Dim lngPos As Long
    Dim s As String

    Set ret = New Collection

    For i = 0 To lstSetting.ListCount - 1
    
        strClass = lstSetting.List(i, C_CLASS)
        If strClass <> "" Then
            Set ss = New SectionStructDTO
            Set ss.classObj = rlxCreateSectionObject(strClass)
            
            If lstSetting.List(i, C_ENABLE) = C_TRUE Then
                ss.useFormat = True
            Else
                ss.useFormat = False
            End If
            
            ss.fontName = lstSetting.List(i, C_FONT_NAME)
            ss.fontSize = lstSetting.List(i, C_FONT_SIZE)
            
            If lstSetting.List(i, C_FONT_BOLD) = C_TRUE Then
                ss.fontBold = True
            Else
                ss.fontBold = False
            End If
            
            If lstSetting.List(i, C_FONT_ITALIC) = C_TRUE Then
                ss.fontItalic = True
            Else
                ss.fontItalic = False
            End If
            
            If lstSetting.List(i, C_FONT_UNDER_LINE) = C_TRUE Then
                ss.fontUnderLine = True
            Else
                ss.fontUnderLine = False
            End If
            
            ss.useFormat2 = chkUseFormat2.Value
            
            ss.fontName2 = cboFont2.Text
            ss.fontSize2 = txtFontSize2.Text
            
            ss.fontBold2 = chkFontBold2.Value
            ss.fontItalic2 = chkFontItalic2.Value
            ss.fontUnderLine2 = chkFontUnderLine2.Value
            
            ret.Add ss, Format$(i + 1, "00")
            Set ss = Nothing
        End If
        
    Next
  
    Set Start = ret
    Set ret = Nothing
    
End Function


Private Sub MW_WheelDown(obj As Object)

    On Error GoTo e

    If obj.ListCount = 0 Then Exit Sub
    obj.TopIndex = obj.TopIndex + 3
e:
End Sub

Private Sub MW_WheelUp(obj As Object)

    Dim lngPos As Long

    On Error GoTo e

    If obj.ListCount = 0 Then Exit Sub
    lngPos = obj.TopIndex - 3

    If lngPos < 0 Then
        lngPos = 0
    End If

    obj.TopIndex = lngPos
e:
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = Nothing
End Sub

Private Sub UserForm_Terminate()
    
    MW.Uninstall
    Set MW = Nothing
End Sub
