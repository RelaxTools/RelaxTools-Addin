VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetManager 
   Caption         =   "シート管理"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385.001
   OleObjectBlob   =   "frmSheetManager.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSheetManager"
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
Private Const C_SHEET_NO As Long = 0
Private Const C_SHEET_STATUS As Long = 1
Private Const C_SHEET_DSP_NAME As Long = 2
Private Const C_SHEET_NEW_NAME As Long = 3
Private Const C_SHEET_OLD_NAME As Long = 4
Private Const C_SHEET_OLD_STATUS As Long = 5
Private Const C_SHEET_OLD_POS As Long = 6

Private Const C_SORT_ASC As Long = 0
Private Const C_SORT_DESC As Long = 1

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2
'Private Const C_DEL As Long = 3

Private Const C_HIDE As String = " －"
Private Const C_SHOW As String = " ○"
Private Const C_DEL As String = "削除"

Private mBook As Workbook

Private mSainyu As Boolean

Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1



Private Sub cmdPrint_Click()

    Dim lngCnt As Long
    Dim strSheets As String
    Dim r As Range
    Dim s As Worksheet

    strSheets = ""
    For lngCnt = 0 To lstSheet.ListCount - 1

        If lstSheet.Selected(lngCnt) And lstSheet.List(lngCnt, C_SHEET_OLD_STATUS) = C_SHOW Then

            Err.Clear
            On Error Resume Next
            Set s = mBook.Sheets(lstSheet.List(lngCnt, C_SHEET_OLD_NAME))
            If Err.Number = 0 And s.Type = xlWorksheet Then
                If s.PageSetup.Pages.Count > 0 Then
    
                    If strSheets = "" Then
                        strSheets = lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
                    Else
                        strSheets = strSheets & vbTab & lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
                    End If
    
                End If
            End If
        End If
    Next

    If strSheets = "" Then
        MsgBox "プレビューできるシートがありません。", vbOKOnly + vbExclamation, C_TITLE
    Else
        Me.Hide
        mBook.Sheets(Split(strSheets, vbTab)).PrintPreview
        Me.Show
    End If

    
End Sub
Private Sub cmdSaveBook_Click()

    Dim lngCnt As Long
    Dim strSheets As String
    Dim r As Range
    Dim b As Workbook

    strSheets = ""
    For lngCnt = 0 To lstSheet.ListCount - 1

        If lstSheet.Selected(lngCnt) Then
            If lstSheet.List(lngCnt, C_SHEET_OLD_STATUS) = C_SHOW Then
    
                If strSheets = "" Then
                    strSheets = lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
                Else
                    strSheets = strSheets & vbTab & lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
                End If
    
            Else
                MsgBox "表示シート以外は保存できません。", vbOKOnly + vbExclamation, C_TITLE
                Exit Sub
            End If
        End If
    Next

    If strSheets = "" Then
    Else
        Dim o As Object
        Dim vntFileName As Variant
        
        vntFileName = Application.GetSaveAsFilename(InitialFileName:="", fileFilter:="Excel ブック(*.xlsx),*.xlsx,Excel マクロ有効ブック(*.xlsm),*.xlsm,Excel 97-2003ブック(*.xls),*.xls", Title:="ブックの保存")
        
        If vntFileName <> False Then
        
            For Each b In Workbooks
                If UCase(b.Name) = UCase(rlxGetFullpathFromFileName(vntFileName)) Then
                    MsgBox "現在開いているブックと同じ名前は指定できません。", vbOKOnly + vbExclamation, C_TITLE
                    Exit Sub
                End If
            Next
        
            Application.DisplayAlerts = False
            
            ActiveWorkbook.Sheets(Split(strSheets, vbTab)).Copy
            
            'ActiveWorkbook.Windows(1).SelectedSheets.Copy
            Set b = Application.Workbooks(Application.Workbooks.Count)
            
            Select Case LCase(Mid$(vntFileName, InStr(vntFileName, ".") + 1))
                Case "xls"
                    b.SaveAs filename:=vntFileName, FileFormat:=xlExcel8, local:=True
                Case "xlsm"
                    b.SaveAs filename:=vntFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled, local:=True
                Case Else
                    b.SaveAs filename:=vntFileName, local:=True
            End Select
            b.Close
            Set b = Nothing
            Application.DisplayAlerts = True
            MsgBox "保存しました。", vbOKOnly + vbInformation, C_TITLE
        End If
    End If

End Sub

Private Sub lstSheet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set MW.obj = lstSheet
End Sub

Private Sub txtSheetName_Change()

    If mSainyu Then
        Exit Sub
    End If
    
'    Dim strBuf As String
'    Dim lngLen As Long
'    Dim i As Long
'
'    Select Case Len(txtSheetName.Text)
'        Case 1 To 31
'        Case Else
'            Call errorMsg
'            Exit Sub
'    End Select
'
'    strBuf = ":\/?*[]：￥／？＊［］"
'    lngLen = Len(strBuf)
'
'    For i = 1 To lngLen
'
'        If InStr(txtSheetName.Text, Mid$(strBuf, i, 1)) > 0 Then
'            Call errorMsg
'            Exit Sub
'        End If
'
'    Next
'
'    For i = 0 To lstSheet.ListCount - 1
'
'        If lstSheet.Selected(i) Then
'            If lstSheet.List(i, C_SHEET_STATUS) = C_DEL Then
'                MsgBox "削除予定のシート名の修正はできません。", vbOKOnly + vbExclamation, C_TITLE
'                Exit Sub
'            End If
'        End If
'
'    Next

    Dim lngCnt As Long
    
    For lngCnt = 0 To lstSheet.ListCount - 1
    
        If lstSheet.Selected(lngCnt) Then
        
            lstSheet.List(lngCnt, C_SHEET_DSP_NAME) = txtSheetName.Text
            lstSheet.List(lngCnt, C_SHEET_NEW_NAME) = txtSheetName.Text
        
        End If
    Next
    
End Sub

'Private Sub UserForm_Activate()
'    MW.Activate
'End Sub

'------------------------------------------------------------------------------------------------------------------------
' リスト初期表示イベント
'------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()


'    Dim i As Long
'    Dim sh As Object
'    Set mBook = ActiveWorkbook
'    Dim blnSw As Boolean
'
'    blnSw = False
    Call refreshList
    
'    For i = 0 To lstSheet.ListCount - 1
'
'        For Each sh In mBook.Windows(1).SelectedSheets
'            If lstSheet.List(i, C_SHEET_DSP_NAME) = sh.Name Then
'                If blnSw = False Then
'                    lstSheet.TopIndex = i
'                    lstSheet.ListIndex = i
'                    blnSw = True
'                End If
'                lstSheet.Selected(i) = True
'            End If
'        Next
'    Next
    
    If mBook.MultiUserEditing Then
        cmdDel.enabled = False
        cmdUndo.enabled = False
    End If

    Set MW = New MouseWheel
    MW.Install Me
    
End Sub
'------------------------------------------------------------------------------------------------------------------------
' リフレッシュ処理
'------------------------------------------------------------------------------------------------------------------------
Private Sub refreshList()

    Dim i As Long
    Dim sh As Object
    Dim blnSw As Boolean
    Dim WS As Object
    Dim lngCount As Long

    Set mBook = ActiveWorkbook
    
    blnSw = False

    lngCount = 0
    lstSheet.Clear
    
    For Each WS In mBook.Sheets
        
        Dim strStatus As String
        If WS.visible = xlSheetVisible Then
            strStatus = C_SHOW
        Else
            strStatus = C_HIDE
        End If
        
        lstSheet.AddItem ""
        lstSheet.List(lngCount, C_SHEET_NO) = Right("   " & lngCount + 1, 3)
        lstSheet.List(lngCount, C_SHEET_STATUS) = strStatus
        lstSheet.List(lngCount, C_SHEET_DSP_NAME) = WS.Name
        lstSheet.List(lngCount, C_SHEET_NEW_NAME) = WS.Name
        lstSheet.List(lngCount, C_SHEET_OLD_NAME) = WS.Name
        lstSheet.List(lngCount, C_SHEET_OLD_STATUS) = strStatus
        lstSheet.List(lngCount, C_SHEET_OLD_POS) = lngCount
        lngCount = lngCount + 1
    
    Next
    
    For i = 0 To lstSheet.ListCount - 1

        For Each sh In mBook.Windows(1).SelectedSheets
            If lstSheet.List(i, C_SHEET_DSP_NAME) = sh.Name Then
                If blnSw = False Then
                    lstSheet.TopIndex = i
                    lstSheet.ListIndex = i
                    blnSw = True
                End If
                lstSheet.Selected(i) = True
            End If
        Next
    Next
    
End Sub
'------------------------------------------------------------------------------------------------------------------------
' リスト変更イベント
'------------------------------------------------------------------------------------------------------------------------
Private Sub lstSheet_Change()

    Dim lngCnt As Long
    Dim strSheets As String
    
    For lngCnt = 0 To lstSheet.ListCount - 1
    
        If lstSheet.Selected(lngCnt) Then
            
            mSainyu = True
            txtSheetName.Text = lstSheet.List(lngCnt, C_SHEET_NEW_NAME)
            mSainyu = False
            Exit For
        
        End If
    Next

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 閉じるボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
    Unload Me
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 選択ボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdSelect_Click()

    Dim lngCnt As Long
    Dim strSheets As String
    Dim r As Range

    strSheets = ""
    For lngCnt = 0 To lstSheet.ListCount - 1

        If lstSheet.Selected(lngCnt) And lstSheet.List(lngCnt, C_SHEET_OLD_STATUS) = C_SHOW Then

            If strSheets = "" Then
                strSheets = lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
            Else
                strSheets = strSheets & vbTab & lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
            End If

        End If
    Next

    If strSheets = "" Then
        MsgBox "選択できるシートがありません。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If

    mBook.Sheets(Split(strSheets, vbTab)).Select
    
    Dim lngPos As Long
    Dim i As Long
    
    For i = 1 To ActiveWindow.SelectedSheets(1).Index
            
        If mBook.Sheets(i).visible = xlSheetVisible Then
            lngPos = lngPos + 1
        End If
    
    Next
    
    '最初に移動してから表示されているシート分移動する。
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    ActiveWindow.ScrollWorkbookTabs Sheets:=lngPos - 1
    Unload Me

End Sub
Private Sub lstSheet_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call cmdSelect_Click
End Sub
'------------------------------------------------------------------------------------------------------------------------
' シート名変更ボタン
'------------------------------------------------------------------------------------------------------------------------
'Private Sub btnChange_Click()
'
'    Dim strBuf As String
'    Dim lngLen As Long
'    Dim i As Long
'
'    Select Case Len(txtSheetName.Text)
'        Case 1 To 31
'        Case Else
'            Call errorMsg
'            Exit Sub
'    End Select
'
'    strBuf = ":\/?*[]：￥／？＊［］"
'    lngLen = Len(strBuf)
'
'    For i = 1 To lngLen
'
'        If InStr(txtSheetName.Text, Mid$(strBuf, i, 1)) > 0 Then
'            Call errorMsg
'            Exit Sub
'        End If
'
'    Next
'
'    For i = 0 To lstSheet.ListCount - 1
'
'        If lstSheet.Selected(i) Then
'            If lstSheet.List(i, C_SHEET_STATUS) = C_DEL Then
'                MsgBox "削除予定のシート名の修正はできません。", vbOKOnly + vbExclamation, C_TITLE
'                Exit Sub
'            End If
'        End If
'
''        If lstSheet.List(i, C_SHEET_NEW_NAME) = txtSheetName.Text Then
''            MsgBox "シートの名前をほかのシート、Visual Basic で参照されるオブジェクト ライブラリまたはワークシートと同じ名前に変更することはできません。", vbOKOnly + vbExclamation, C_TITLE
''            Exit Sub
''        End If
'
'    Next
'
'    Dim lngCnt As Long
'
'    For lngCnt = 0 To lstSheet.ListCount - 1
'
'        If lstSheet.Selected(lngCnt) Then
'
'            lstSheet.List(lngCnt, C_SHEET_DSP_NAME) = txtSheetName.Text
'            lstSheet.List(lngCnt, C_SHEET_NEW_NAME) = txtSheetName.Text
''            Exit For
'
'        End If
'    Next
'
'End Sub
'------------------------------------------------------------------------------------------------------------------------
' メッセージ表示
'------------------------------------------------------------------------------------------------------------------------
Private Sub errorMsg()
    MsgBox "入力されたシートまたはグラフの名前が正しくありません。次の点を確認して修正してください。" & vbCrLf & vbCrLf & _
    "・入力文字が31文字以内であること" & vbCrLf & _
    "・次の使用できない文字が含まれていないこと(全角も含む):コロン(:)、円記号(\)、スラッシュ(/)、バックスラッシュ(＼)、疑問符(?)、アスタリスク(*)、シングルコーテーション(')、左角かっこ([)、右角かっこ(])、アポストロフィー(＇)" & vbCrLf & _
    "・「履歴」という名前は予約語なのでシート名には使えません。" & vbCrLf & _
    "・名前が空白でないこと", vbOKOnly + vbExclamation, C_TITLE

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 昇順ソートボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdAsc_Click()

    '昇順ソート
    sortList C_SORT_ASC

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 降順ソートボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdDesc_Click()
    
    '降順ソート
    sortList C_SORT_DESC

End Sub
'------------------------------------------------------------------------------------------------------------------------
' ソート処理
'------------------------------------------------------------------------------------------------------------------------
Private Sub sortList(ByVal lngSort As Long)
    
    Dim lngCnt As Long
    Dim lngCmp As Long
    Dim strCmp1 As String
    Dim strCmp2 As String
    
    Dim varTmp As Variant

    Dim idx  As New Collection
    Dim cnt As Long
    

    '１つならソート不要
    If lstSheet.ListCount <= 1 Then
        Exit Sub
    End If

    cnt = 0
    For lngCnt = 0 To lstSheet.ListCount - 1
        If lstSheet.Selected(lngCnt) Then
            cnt = cnt + 1
        End If
    Next

    If cnt > 1 Then
        For lngCnt = 0 To lstSheet.ListCount - 1
            If lstSheet.Selected(lngCnt) Then
                idx.Add lngCnt
            End If
        Next
    Else
        For lngCnt = 0 To lstSheet.ListCount - 1
            idx.Add lngCnt
        Next
    End If
    
'    For lngCnt = 0 To lstSheet.ListCount - 1 - 1
'
'        For lngCmp = lngCnt + 1 To lstSheet.ListCount - 1
'
'            If lngSort = C_SORT_ASC Then
'                strCmp1 = lstSheet.List(lngCnt, C_SHEET_NEW_NAME)
'                strCmp2 = lstSheet.List(lngCmp, C_SHEET_NEW_NAME)
'            Else
'                strCmp2 = lstSheet.List(lngCnt, C_SHEET_NEW_NAME)
'                strCmp1 = lstSheet.List(lngCmp, C_SHEET_NEW_NAME)
'            End If
'
'            If strCmp1 > strCmp2 Then
'                Dim i As Long
'                For i = C_SHEET_STATUS To C_SHEET_OLD_POS
'                    varTmp = lstSheet.List(lngCnt, i)
'                    lstSheet.List(lngCnt, i) = lstSheet.List(lngCmp, i)
'                    lstSheet.List(lngCmp, i) = varTmp
'                Next
'            End If
'        Next
'    Next

    For lngCnt = 1 To idx.Count - 1
    
        For lngCmp = lngCnt + 1 To idx.Count
                
            If lngSort = C_SORT_ASC Then
                strCmp1 = lstSheet.List(idx(lngCnt), C_SHEET_NEW_NAME)
                strCmp2 = lstSheet.List(idx(lngCmp), C_SHEET_NEW_NAME)
            Else
                strCmp2 = lstSheet.List(idx(lngCnt), C_SHEET_NEW_NAME)
                strCmp1 = lstSheet.List(idx(lngCmp), C_SHEET_NEW_NAME)
            End If
            
            If strCmp1 > strCmp2 Then
                Dim i As Long
                For i = C_SHEET_STATUS To C_SHEET_OLD_POS
                    varTmp = lstSheet.List(idx(lngCnt), i)
                    lstSheet.List(idx(lngCnt), i) = lstSheet.List(idx(lngCmp), i)
                    lstSheet.List(idx(lngCmp), i) = varTmp
                Next
            End If
        Next
    Next


End Sub
'------------------------------------------------------------------------------------------------------------------------
' 表示ボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdShow_Click()

    Dim lngCnt As Long
    
    For lngCnt = 0 To lstSheet.ListCount - 1
    
        If lstSheet.Selected(lngCnt) Then
        
            lstSheet.List(lngCnt, C_SHEET_STATUS) = C_SHOW
        
        End If
    Next
    
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 非表示ボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdHide_Click()

    Dim lngCnt As Long
    
    For lngCnt = 0 To lstSheet.ListCount - 1
    
        If lstSheet.Selected(lngCnt) Then
        
            lstSheet.List(lngCnt, C_SHEET_STATUS) = C_HIDE
        
        End If
    Next

End Sub

'------------------------------------------------------------------------------------------------------------------------
' 選択行を上に移動
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdUp_Click()
     Call moveList(C_UP)
End Sub

Private Sub cmdDown_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call cmdDown_Click
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
    If lstSheet.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstSheet.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstSheet.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select

    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstSheet.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If
            
            lngCmp = lngCnt + lngInc * -1
            
            Dim i As Long
            For i = C_SHEET_STATUS To C_SHEET_OLD_POS
                varTmp = lstSheet.List(lngCnt, i)
                lstSheet.List(lngCnt, i) = lstSheet.List(lngCmp, i)
                lstSheet.List(lngCmp, i) = varTmp
            Next
            
            lstSheet.Selected(lngCnt) = False
            lstSheet.Selected(lngCnt + lngInc * -1) = True
        End If
    
    Next

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 削除ボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdDel_Click()

    Dim lngCnt As Long
    Dim strBuf As String
    
    For lngCnt = 0 To lstSheet.ListCount - 1
    
        If lstSheet.Selected(lngCnt) Then
        
'            strBuf = "≪削除≫" & lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
'            lstSheet.List(lngCnt, C_SHEET_DSP_NAME) = strBuf
'            lstSheet.List(lngCnt, C_SHEET_NEW_NAME) = lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
'            lstSheet.List(lngCnt, C_SHEET_DEL) = C_DEL
            lstSheet.List(lngCnt, C_SHEET_STATUS) = C_DEL
        
        End If
    Next
    
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 削除解除ボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdUndo_Click()

    Dim lngCnt As Long
    Dim strBuf As String
    
    For lngCnt = 0 To lstSheet.ListCount - 1
    
        If lstSheet.List(lngCnt, C_SHEET_STATUS) = C_DEL Then
        
            strBuf = lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
            lstSheet.List(lngCnt, C_SHEET_DSP_NAME) = strBuf
            lstSheet.List(lngCnt, C_SHEET_STATUS) = lstSheet.List(lngCnt, C_SHEET_OLD_STATUS)
        
        End If
    Next

End Sub
'------------------------------------------------------------------------------------------------------------------------
'初期状態に戻すボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdInitialize_Click()
    Call refreshList

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 変更内容を反映ボタン
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdSubmit_Click()


    On Error GoTo ErrHandle
    
    Const C_TEMP_NAME As String = "~~temp"
    Dim strSheetName As String


    Dim WS As Object
    Dim lngCnt As Long
    Dim lngCnt2 As Long
    Dim lngVisibleCount As Long
    Dim lngDeleteCount As Long

    Dim i As Long
    
    strSheetName = C_TEMP_NAME & "_" & Format(Now, "yyyymmddhhmmss") & "_"

    Dim lngLast As Long

    For lngCnt = 0 To lstSheet.ListCount - 1
    
        If lstSheet.List(lngCnt, C_SHEET_NEW_NAME) <> lstSheet.List(lngCnt, C_SHEET_OLD_NAME) Then
            If IsErrSheetNameChar(lstSheet.List(lngCnt, C_SHEET_NEW_NAME)) Or Len(Trim(lstSheet.List(lngCnt, C_SHEET_NEW_NAME))) = 0 Or Len(Trim(lstSheet.List(lngCnt, C_SHEET_NEW_NAME))) > 31 Then
                Call errorMsg
                Exit Sub
            End If
        End If
    
    Next
    
    lngVisibleCount = 0
    For lngCnt = 0 To lstSheet.ListCount - 1
        If lstSheet.List(lngCnt, C_SHEET_STATUS) = C_HIDE Then
            lngVisibleCount = lngVisibleCount + 1
        End If
    Next
    
    lngDeleteCount = 0
    For lngCnt = 0 To lstSheet.ListCount - 1
    
        '異なる場合、リストを正とし、現在のシートの前に移動。
        If lstSheet.List(lngCnt, C_SHEET_STATUS) = C_DEL Then
            lngDeleteCount = lngDeleteCount + 1
        End If
        
    Next
    
    If (lngVisibleCount = lstSheet.ListCount) Or (lngDeleteCount = lstSheet.ListCount) Or (lngDeleteCount + lngVisibleCount = lstSheet.ListCount) Then
        MsgBox "すべてのシートを非表示・削除はできません。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    For lngCnt = 0 To lstSheet.ListCount - 2
    
        For lngCnt2 = lngCnt + 1 To lstSheet.ListCount - 1
            If lstSheet.List(lngCnt, C_SHEET_STATUS) <> C_DEL And lstSheet.List(lngCnt2, C_SHEET_STATUS) <> C_DEL Then
                If StrConv(UCase(lstSheet.List(lngCnt, C_SHEET_NEW_NAME)), vbNarrow) = StrConv(UCase(lstSheet.List(lngCnt2, C_SHEET_NEW_NAME)), vbNarrow) Then
                    MsgBox "シートの名前をほかのシート、Visual Basic で参照されるオブジェクト ライブラリまたはワークシートと同じ名前に変更することはできません。", vbOKOnly + vbExclamation, C_TITLE
                    Exit Sub
                End If
            End If
        Next
    Next
    
    
    If MsgBox("編集内容を反映します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
'    lngLast = lstSheet.ListIndex

    Dim strDel As String
    For lngCnt = lstSheet.ListCount - 1 To 0 Step -1
        
        'シートの削除
        If lstSheet.List(lngCnt, C_SHEET_STATUS) = C_DEL Then
            strDel = lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
            mBook.Sheets(strDel).Delete
            lstSheet.RemoveItem lngCnt
        End If
        
    Next
    
    Set WS = mBook.ActiveSheet
    'シートの並び順反映
    For lngCnt = 0 To lstSheet.ListCount - 1
    
        '名称が同じなら何もしない
        If mBook.Sheets(lngCnt + 1).visible = xlSheetVeryHidden Then
        Else
            If mBook.Sheets(lngCnt + 1).Name = lstSheet.List(lngCnt, C_SHEET_OLD_NAME) Then
            Else
                '異なる場合、リストを正とし、現在のシートの前に移動。
                mBook.Sheets(lstSheet.List(lngCnt, C_SHEET_OLD_NAME)).Move Before:=mBook.Sheets(lngCnt + 1)
            End If
        End If
    Next
    'もともとアクティブだったシートを選択
    WS.Select
    
    '表示→非表示の順番に行う。（途中で全非表示になる可能性があるため）
    For lngCnt = 0 To lstSheet.ListCount - 1
        
        '表示
        Select Case lstSheet.List(lngCnt, C_SHEET_STATUS)
            Case C_SHOW
                mBook.Sheets(lstSheet.List(lngCnt, C_SHEET_OLD_NAME)).visible = xlSheetVisible
        End Select
        
    Next
    For lngCnt = 0 To lstSheet.ListCount - 1
        
        '非表示
        Select Case lstSheet.List(lngCnt, C_SHEET_STATUS)
            Case C_HIDE
                mBook.Sheets(lstSheet.List(lngCnt, C_SHEET_OLD_NAME)).visible = xlSheetHidden
        End Select
        
    Next
    
    For lngCnt = 0 To lstSheet.ListCount - 1
        Dim strOld As String
        Dim strNew As String

        
        'シート名変更
        If lstSheet.List(lngCnt, C_SHEET_NEW_NAME) <> lstSheet.List(lngCnt, C_SHEET_OLD_NAME) Then
            strNew = strSheetName & lngCnt
            strOld = lstSheet.List(lngCnt, C_SHEET_OLD_NAME)
            If strNew <> strOld Then
                mBook.Sheets(strOld).Name = strNew
            End If
        End If
        
    Next
    
    For lngCnt = 0 To lstSheet.ListCount - 1
        
        'シート名変更
        If lstSheet.List(lngCnt, C_SHEET_NEW_NAME) <> lstSheet.List(lngCnt, C_SHEET_OLD_NAME) Then
            strNew = lstSheet.List(lngCnt, C_SHEET_NEW_NAME)
            strOld = strSheetName & lngCnt
            If strNew <> strOld Then
                mBook.Sheets(strOld).Name = strNew
            End If
        End If
    Next
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    '再表示
    Call refreshList
    
'    If lngLast > lstSheet.ListCount - 1 Then
'        lngLast = lstSheet.ListCount - 1
'    End If
'
'    If lstSheet.ListCount > 0 Then
'        lstSheet.Selected(lngLast) = True
'    End If

    Exit Sub
ErrHandle:
    MsgBox "エラーが発生しました。", vbOKOnly, C_TITLE

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
