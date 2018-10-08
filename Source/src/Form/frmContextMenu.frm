VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContextMenu 
   Caption         =   "右クリックメニュー割り当て"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10185
   OleObjectBlob   =   "frmContextMenu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmContextMenu"
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
Const C_SET_NO As Long = 1
Const C_SET_KEY As Long = 2
Const C_SET_DISP_NAME As Long = 3

Const C_MENU_DISP As Long = 0
Const C_MENU_SETTING As Long = 1
Const C_MENU_KEY As Long = 2

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2
Private mBarFav As Object

Private mblnFlg As Boolean

Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1

Private Sub cboCategory_Click()
    Call dispCommand
End Sub

Private Sub cboCategory_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = cboCategory
End Sub

'Private Function existMacro(ByVal strMacro As String) As Boolean
'
'    Dim i As Long
'    Dim WS As Worksheet
'
'    existMacro = False
'
'    Set WS = ThisWorkbook.Worksheets("HELP")
'    i = C_COM_DATA
'
'    'マクロシートのロード
'    Do Until WS.Cells(i, C_COM_NO).Value = ""
'
'        If WS.Cells(i, C_COM_USE).Value <> "－" Then
'
'            If WS.Cells(i, C_COM_MACRO).Value = strMacro Then
'                existMacro = True
'                Exit Do
'            End If
'
'        End If
'        i = i + 1
'    Loop
'
'End Function

Private Sub cmdAdd_Click()

    Dim j As Long

    lstMenu2.AddItem ""
    j = lstMenu2.ListCount - 1
    lstMenu2.List(j, 0) = lstCommand.List(lstCommand.ListIndex, 2)
    lstMenu2.List(j, 1) = lstCommand.List(lstCommand.ListIndex, 3)
    
    lstMenu2.ListIndex = j
    
    Call SetList

End Sub

Private Sub cmdAddMenu_Click()
    
    frmContextMenuAdd.Start True
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()

    Dim j As Long

    If lstMenu2.ListIndex < 0 Then
        Exit Sub
    End If

    j = lstMenu2.ListIndex
    lstMenu2.RemoveItem lstMenu2.ListIndex
    
    Call SetList
    
    j = j - 1
    If j < 0 Then
    Else
        lstMenu2.ListIndex = j
    End If
    
End Sub

Private Sub cmdDown_Click()
    Call moveList(C_DOWN)
End Sub

Private Sub cmdSave_Click()

    Dim i As Long
    
    For i = 0 To lstMenu1.ListCount - 1
        Call SaveSetting(C_TITLE, "ContextMenu", lstMenu1.List(i, C_MENU_KEY), lstMenu1.List(i, C_MENU_SETTING))
    Next
    Unload Me

End Sub

Private Sub cmdSep_Click()

    Dim i As Long
    Dim j As Long
    
    j = lstMenu2.ListIndex
    lstMenu2.AddItem ""
    For i = lstMenu2.ListCount - 1 To j + 1 Step -1
        lstMenu2.List(i, 0) = lstMenu2.List(i - 1, 0)
        lstMenu2.List(i, 1) = lstMenu2.List(i - 1, 1)
    Next
    
    lstMenu2.List(j, 0) = "----------------------------------------"
    lstMenu2.List(j, 1) = "-"
    
    lstMenu2.ListIndex = j
    
    Call SetList

End Sub

Private Sub cmdUp_Click()
    Call moveList(C_UP)
End Sub



Private Sub lstCommand_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Set MW.obj = lstCommand
    
End Sub

Private Sub lstMenu1_Click()

    Dim strBuf As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim i As Long
    Dim j As Long
    
    If mblnFlg Then
        Exit Sub
    End If
    
    If lstMenu1.ListCount < 0 Then
        Exit Sub
    End If
    
    lstMenu2.Clear
    
    strBuf = lstMenu1.List(lstMenu1.ListIndex, C_MENU_SETTING)

    varRow = Split(strBuf, vbCrLf)
    
    For i = LBound(varRow) To UBound(varRow) - 1
    
        varCol = Split(varRow(i), vbTab)
        
        lstMenu2.AddItem
        lstMenu2.List(lstMenu2.ListCount - 1, 0) = varCol(1)
        lstMenu2.List(lstMenu2.ListCount - 1, 1) = varCol(2)
        
        
    Next
    
    If lstMenu2.ListCount > 0 Then
        lstMenu2.ListIndex = 0
    End If
        
End Sub

Private Sub SetList()

    Dim j As Long
    Dim i As Long
    Dim strBuf As String
    
    mblnFlg = True
    j = lstMenu1.ListIndex
    
    strBuf = ""
    For i = 0 To lstMenu2.ListCount - 1
        strBuf = strBuf & lstMenu1.List(j, C_MENU_DISP) & vbTab & lstMenu2.List(i, 0) & vbTab & lstMenu2.List(i, 1) & vbCrLf
    Next

    lstMenu1.List(j, C_MENU_SETTING) = strBuf
    
    mblnFlg = False

End Sub



Private Sub lstMenu1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 2 Then
    
        Set mBarFav = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
        With mBarFav
            Dim myCBCtrl2 As Variant
            Set myCBCtrl2 = .Controls.Add(Type:=msoControlPopup)
            With myCBCtrl2
                .Caption = "メニューのコピー先を選択"
            End With

            Dim a As Variant
            Dim i As Long
            For i = 0 To lstMenu1.ListCount - 1
                If i <> lstMenu1.ListIndex Then
                    With myCBCtrl2.Controls.Add
                        .Caption = lstMenu1.List(i)
                        .OnAction = "'basContextMenu.copyMenu(""" & i & """)'"
                       '.FaceId = 526
                        .FaceId = 525
                    End With
                End If
            Next

        
        End With
        mBarFav.ShowPopup
    
    End If
End Sub
Public Sub copyMenu(ByVal Index As Long)

    Dim ret As VbMsgBoxResult
    
    ret = vbOK

    If lstMenu1.List(Index, 1) = "" Then
    Else
        ret = MsgBox("上書きされますがよろしいですか？", vbOKCancel + vbQuestion, C_TITLE)
    End If
    
    If ret = vbOK Then
        lstMenu1.List(Index, 1) = lstMenu1.List(lstMenu1.ListIndex, 1)
        lstMenu1.ListIndex = Index
    End If

End Sub

Private Sub lstMenu1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstMenu1
End Sub

Private Sub lstMenu2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstMenu2
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
    Dim lngCount As Long

    mblnFlg = False

    Set WS = ThisWorkbook.Worksheets("HELP")
    i = C_COM_DATA
    j = 0

    Dim strBefore As String
    strBefore = ""
    
    cboCategory.AddItem "すべて"
    'マクロシートのロード
    Do Until WS.Cells(i, C_COM_NO).Value = ""
    
        Select Case True
            Case WS.Cells(i, C_COM_USE).Value = "－"
            Case WS.Cells(i, C_COM_CATEGORY).Value = "ショートカットキー用"
            Case WS.Cells(i, C_COM_CATEGORY).Value = "Excel機能"
            Case Else
                If WS.Cells(i, C_COM_CATEGORY).Value <> strBefore Then
                    cboCategory.AddItem WS.Cells(i, C_COM_CATEGORY).Value
                    strBefore = WS.Cells(i, C_COM_CATEGORY).Value
                End If
        End Select
        i = i + 1

    Loop
    cboCategory.ListIndex = 0

    Call dispCommand

    lngCount = GetSetting(C_TITLE, "ContextMenu", "Count", 0)
        
    lstMenu1.AddItem ""
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_DISP) = "セルの右クリックメニュー"
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_SETTING) = GetSetting(C_TITLE, "ContextMenu", "ContextMenuCell", "")
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_KEY) = "ContextMenuCell"
    lstMenu1.AddItem ""
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_DISP) = "行の右クリックメニュー"
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_SETTING) = GetSetting(C_TITLE, "ContextMenu", "ContextMenuRow", "")
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_KEY) = "ContextMenuRow"
    lstMenu1.AddItem ""
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_DISP) = "列の右クリックメニュー"
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_SETTING) = GetSetting(C_TITLE, "ContextMenu", "ContextMenuCol", "")
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_KEY) = "ContextMenuCol"
    lstMenu1.AddItem ""
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_DISP) = "シェイプの右クリックメニュー"
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_SETTING) = GetSetting(C_TITLE, "ContextMenu", "ContextMenuShape", "")
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_KEY) = "ContextMenuShape"
    lstMenu1.AddItem ""
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_DISP) = "ピクチャの右クリックメニュー"
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_SETTING) = GetSetting(C_TITLE, "ContextMenu", "ContextMenuPicture", "")
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_KEY) = "ContextMenuPicture"
    lstMenu1.AddItem ""
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_DISP) = "シートタブの右クリックメニュー"
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_SETTING) = GetSetting(C_TITLE, "ContextMenu", "ContextMenuWorkbookPly", "")
    lstMenu1.List(lstMenu1.ListCount - 1, C_MENU_KEY) = "ContextMenuWorkbookPly"
    
    lstMenu1.ListIndex = 0
    
    Set MW = New MouseWheel
    MW.Install Me
    
End Sub
'Function getEnable(ByVal strBuf As String) As String
'    If strBuf = "1" Then
'        getEnable = C_ENABLE
'    Else
'        getEnable = C_DISABLE
'    End If
'End Function
'Function getFromEnable(ByVal strBuf As String) As String
'    If strBuf = C_ENABLE Then
'        getFromEnable = "1"
'    Else
'        getFromEnable = "0"
'    End If
'End Function
Sub dispCommand()

    Dim WS As Worksheet
    Dim i As Long
    Dim j As Long

    Set WS = ThisWorkbook.Worksheets("HELP")
    i = C_COM_DATA
    j = 0
    
    lstCommand.Clear

    'マクロシートのロード
    Do Until WS.Cells(i, C_COM_NO).Value = ""
        
        Select Case True
            Case WS.Cells(i, C_COM_USE).Value = "－"
            Case WS.Cells(i, C_COM_CATEGORY).Value = "ショートカットキー用"
            Case WS.Cells(i, C_COM_CATEGORY).Value = "Excel機能"
            Case Else
        
                If (cboCategory.ListIndex = 0 Or cboCategory.Text = WS.Cells(i, C_COM_CATEGORY).Value) And (txtKinou.Text = "" Or InStr(WS.Cells(i, C_COM_DISP_NAME).Value, txtKinou.Text) > 0) Then
                    lstCommand.AddItem ""
                    lstCommand.List(j, 0) = j + 1
                    lstCommand.List(j, 1) = WS.Cells(i, C_COM_CATEGORY).Value
                    lstCommand.List(j, 2) = WS.Cells(i, C_COM_DISP_NAME).Value
                    lstCommand.List(j, 3) = WS.Cells(i, C_COM_MACRO).Value
                    j = j + 1
                End If
        End Select
        i = i + 1

    Loop
    If lstCommand.ListCount > 0 Then
        lstCommand.ListIndex = 0
        cmdAdd.enabled = True
    Else
        cmdAdd.enabled = False
    End If
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
    If lstMenu2.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstMenu2.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstMenu2.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select

    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstMenu2.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If
            
            lngCmp = lngCnt + lngInc * -1
            
            Dim i As Long
            For i = 0 To 1
                varTmp = lstMenu2.List(lngCnt, i)
                lstMenu2.List(lngCnt, i) = lstMenu2.List(lngCmp, i)
                lstMenu2.List(lngCmp, i) = varTmp
            Next
            
            lstMenu2.Selected(lngCnt) = False
            lstMenu2.Selected(lngCnt + lngInc * -1) = True
        End If
    
    Next
    Call SetList
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = Nothing
End Sub

Private Sub UserForm_Terminate()
    MW.Uninstall
    Set MW = Nothing
End Sub

Private Sub MW_WheelDown(obj As Object)

    On Error GoTo e

    If obj.ListCount = 0 Then Exit Sub
    obj.TopIndex = obj.TopIndex + 3
e:
End Sub

Private Sub MW_WheelUp(obj As Object)

    On Error GoTo e

    Dim lngPos As Long

    If obj.ListCount = 0 Then Exit Sub
    lngPos = obj.TopIndex - 3

    If lngPos < 0 Then
        lngPos = 0
    End If

    obj.TopIndex = lngPos
e:
End Sub
