VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCombo 
   Caption         =   "まとめ実行"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090
   OleObjectBlob   =   "frmCombo.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCombo"
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

Const C_COM_DATA As Long = 25 '13

Const C_COM_NO As Long = 1
Const C_COM_CATEGORY As Long = 2
Const C_COM_MACRO As Long = 3
Const C_COM_DISP_NAME As Long = 4
Const C_COM_USE As Long = 5

Const C_SET_DATA As Long = 3
Const C_SET_NO As Long = 0
Const C_SET_CATEGORY As Long = 1
Const C_SET_DISP_NAME As Long = 2
Const C_SET_MACRO As Long = 3

Const C_LST_NAME As Long = 0
Const C_LST_DATA As Long = 1

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2

Private mblnSainyu As Boolean

Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1

Private Sub cboCategory_Click()
    Call dispCommand
End Sub



Private Sub cmdAdd_Click()

    Dim j As Long
    Dim i As Long
    Dim blnFind As Boolean
    Dim strKey As String

    blnFind = False

    If lstCommand.ListCount = 0 Then
        Exit Sub
    End If

    strKey = lstCommand.List(lstCommand.ListIndex, C_SET_MACRO)

    j = lstCombo.ListCount
    lstCombo.AddItem ""

    lstCombo.List(j, C_SET_NO) = j + 1
    lstCombo.List(j, C_SET_CATEGORY) = lstCommand.List(lstCommand.ListIndex, C_SET_CATEGORY)
    lstCombo.List(j, C_SET_DISP_NAME) = lstCommand.List(lstCommand.ListIndex, C_SET_DISP_NAME)
    lstCombo.List(j, C_SET_MACRO) = lstCommand.List(lstCommand.ListIndex, C_SET_MACRO)

    For i = 0 To lstCombo.ListCount - 1
        lstCombo.Selected(i) = False
    Next

    lstCombo.Selected(j) = True

    If lstCombo.ListCount > 0 Then
        cmdDel.enabled = True
    Else
        cmdDel.enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdDel_Click()

    Dim i As Long
    Dim j As Long

    If lstCombo.ListCount > 0 Then
        i = lstCombo.ListIndex
        lstCombo.RemoveItem i
        If i > lstCombo.ListCount - 1 Then
            i = i - 1
            If i < 0 Then
                i = 0
            End If
        Else
            For j = i To lstCombo.ListCount - 1
                lstCombo.List(j, C_SET_NO) = j + 1
            Next
        End If
        
        Call lstCombo_Change
        
        If lstCombo.ListCount > 0 Then
            lstCombo.ListIndex = i
            cmdDel.enabled = True
        Else
            cmdDel.enabled = False
        End If
    End If
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
    If lstCombo.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstCombo.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstCombo.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select

    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstCombo.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If
            
            lngCmp = lngCnt + lngInc * -1
            
            Dim i As Long
            For i = C_SET_CATEGORY To C_SET_MACRO
                varTmp = lstCombo.List(lngCnt, i)
                lstCombo.List(lngCnt, i) = lstCombo.List(lngCmp, i)
                lstCombo.List(lngCmp, i) = varTmp
            Next
            
            lstCombo.Selected(lngCnt) = False
            lstCombo.Selected(lngCnt + lngInc * -1) = True
        End If
    
    Next

End Sub

Private Sub lstCombo_Change()

    Dim i As Long
    Dim strBuf As String
    
    If mblnSainyu Then
        Exit Sub
    End If
    
    For i = 0 To lstCombo.ListCount - 1
    
        strBuf = strBuf & lstCombo.List(i, C_SET_NO) & vbTab
        strBuf = strBuf & lstCombo.List(i, C_SET_CATEGORY) & vbTab
        strBuf = strBuf & lstCombo.List(i, C_SET_DISP_NAME) & vbTab
        strBuf = strBuf & lstCombo.List(i, C_SET_MACRO)
        
        If i = lstCombo.ListCount - 1 Then
        Else
            strBuf = strBuf & vbVerticalTab
        End If
    Next

    mblnSainyu = True
    lstSetting.List(lstSetting.ListIndex, C_LST_DATA) = strBuf
    mblnSainyu = False

End Sub

Private Sub lstCombo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim i As Long
    Dim strBuf As String

    strBuf = lstCombo.List(lstCombo.ListIndex, C_SET_DISP_NAME)
    
    For i = 0 To lstCommand.ListCount - 1
        If lstCommand.List(i, C_SET_DISP_NAME) = strBuf Then
            lstCommand.ListIndex = i
            lstCommand.TopIndex = i
            Exit For
        End If
    Next
    
End Sub

Private Sub lstCombo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Set MW.obj = lstCombo

End Sub

Private Sub lstCommand_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Set MW.obj = lstCommand
End Sub

Private Sub lstSetting_Change()

    Dim strBuf As String
    Dim i As Long
    Dim varLine As Variant
    Dim varCol As Variant
    
    If mblnSainyu Then
        Exit Sub
    End If
    
    If lstSetting.ListIndex < 0 Then
        Exit Sub
    End If
    
    mblnSainyu = True
    lstCombo.Clear
    mblnSainyu = False
    
    strBuf = lstSetting.List(lstSetting.ListIndex, C_LST_DATA)
    If strBuf = "" Then
        cmdDel.enabled = False
        Exit Sub
    End If
    
    varLine = Split(strBuf, vbVerticalTab)
    
    For i = LBound(varLine) To UBound(varLine)
    
        varCol = Split(varLine(i), vbTab)
        
        lstCombo.AddItem ""
        lstCombo.List(i, C_SET_NO) = i + 1
        lstCombo.List(i, C_SET_CATEGORY) = varCol(C_SET_CATEGORY)
        lstCombo.List(i, C_SET_DISP_NAME) = varCol(C_SET_DISP_NAME)
        lstCombo.List(i, C_SET_MACRO) = varCol(C_SET_MACRO)
        
    Next

    If lstCombo.ListCount < 0 Then
        cmdDel.enabled = False
    Else
        lstCombo.ListIndex = 0
        cmdDel.enabled = True
    End If

End Sub


Private Sub cmdSave_Click()

    Dim i As Long
    Dim strBuf As String
    Dim strLine As String


    strBuf = ""

    For i = 0 To lstSetting.ListCount - 1
        SaveSetting C_TITLE, "Combo", "ComboList" & i + 1, lstSetting.List(i, C_LST_DATA)
    Next


    'ショートカットキーの登録
    Call setShortCutKey

    Unload Me

End Sub
Private Sub txtKinou_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call dispCommand
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Dim WS As Worksheet
    Dim i As Long
    Dim j As Long

    Set WS = ThisWorkbook.Worksheets("HELP")
    i = C_COM_DATA
    j = 0

    Dim strBefore As String
    strBefore = ""
    
    cboCategory.AddItem "すべて"
    'マクロシートのロード
    Do Until WS.Cells(i, C_COM_NO).value = ""

        If WS.Cells(i, C_COM_USE).value <> "－" And WS.Cells(i, C_COM_CATEGORY).value <> "まとめ実行" Then
            If WS.Cells(i, C_COM_CATEGORY).value <> strBefore Then
                cboCategory.AddItem WS.Cells(i, C_COM_CATEGORY).value
                strBefore = WS.Cells(i, C_COM_CATEGORY).value
            End If
        End If
        i = i + 1

    Loop
    cboCategory.ListIndex = 0
    
    Dim strList() As String
    Dim strResult As String
    Dim lngMax As Long
    
    mblnSainyu = True
    For i = 0 To 4
    
        strResult = GetSetting(C_TITLE, "Combo", "ComboList" & (i + 1), "")
        
        lstSetting.AddItem
        lstSetting.List(i, C_LST_NAME) = "まとめ実行" & StrConv(i + 1, vbWide)
        lstSetting.List(i, C_LST_DATA) = strResult
        
    Next
    mblnSainyu = False
    
    
    If lstSetting.ListCount > 0 Then
        lstSetting.ListIndex = 0
        cmdDel.enabled = True
    Else
        cmdDel.enabled = False
    End If
    
    Set MW = basMouseWheel.GetInstance
    MW.Install
End Sub


Sub dispCommand()

    Dim WS As Worksheet
    Dim i As Long
    Dim j As Long

    Set WS = ThisWorkbook.Worksheets("HELP")
    i = C_COM_DATA
    j = 0
    
    lstCommand.Clear

    'マクロシートのロード
    Do Until WS.Cells(i, C_COM_NO).value = ""
        
        If WS.Cells(i, C_COM_USE).value <> "－" And WS.Cells(i, C_COM_CATEGORY).value <> "まとめ実行" Then
            If (cboCategory.ListIndex = 0 Or cboCategory.Text = WS.Cells(i, C_COM_CATEGORY).value) And (txtKinou.Text = "" Or InStr(WS.Cells(i, C_COM_DISP_NAME).value, txtKinou.Text) > 0) Then
                lstCommand.AddItem ""
                lstCommand.List(j, C_SET_NO) = j + 1
                lstCommand.List(j, C_SET_CATEGORY) = WS.Cells(i, C_COM_CATEGORY).value
                lstCommand.List(j, C_SET_DISP_NAME) = WS.Cells(i, C_COM_DISP_NAME).value
                lstCommand.List(j, C_SET_MACRO) = WS.Cells(i, C_COM_MACRO).value
                j = j + 1
            End If
        End If
        i = i + 1

    Loop
    If lstCommand.ListCount > 0 Then
        lstCommand.ListIndex = 0
        cmdAdd.enabled = True
    Else
        cmdAdd.enabled = False
    End If
End Sub

Private Sub MW_WheelDown(obj As Object)

    If obj.ListCount = 0 Then Exit Sub
    obj.TopIndex = obj.TopIndex + 3
    
End Sub

Private Sub MW_WheelUp(obj As Object)

    Dim lngPos As Long

    If obj.ListCount = 0 Then Exit Sub
    lngPos = obj.TopIndex - 3

    If lngPos < 0 Then
        lngPos = 0
    End If

    obj.TopIndex = lngPos

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Set MW.obj = Nothing
End Sub

Private Sub UserForm_Terminate()
    MW.UnInstall
    Set MW = Nothing
End Sub
