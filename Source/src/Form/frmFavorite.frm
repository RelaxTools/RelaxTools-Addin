VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFavorite 
   Caption         =   "お気に入り"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   OleObjectBlob   =   "frmFavorite.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFavorite"
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


Private WithEvents MW As MouseWheel
Attribute MW.VB_VarHelpID = -1

Option Explicit
Private Const C_FILE_NO As Long = 0
Private Const C_FILE_NAME As Long = 1
Private Const C_PATH_NAME As Long = 2
Private Const C_ORIGINAL As Long = 3
Private Const C_CATEGORY As Long = 4

Private mlngWidth As Long
Private mlngLeft As Long

Private RW As ResizeWindow
Private mlngButtonLeft As Long
Private mlngListWidth As Long
Private mlngListHeight As Long
Private mlngLblBookWidth As Long
Private mlngLblWidth As Long
Private mlngLblTop As Long

Private mlngDetailTop As Long
Private mlngDetailWidth As Long

Private mlngMsgTop As String
Private mlngMsgWidth As String

Private mBarFav As Object
Private mBarFavDrop As Object
Private mBarCat As Object

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2

'Private Const C_ADD As Long = 1
'Private Const C_MOD As Long = 2
Private Const C_HEAD As Long = 3
Private Const C_TAIL As Long = 4
Private Const C_FAV_ALL As String = "規定のカテゴリ"
Private mblnSainyu As Boolean
Private mstrFirstBook As String
Private Const C_FILE_INFO As String = "ファイル情報："

Public mobjCategory As Object

'Private mlngPos As Long


    
Private Sub lstCategory_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Set MW.obj = lstCategory

End Sub

'Private Sub lstFavorite_Enter()
'    mlngPos = -1
'End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Set MW.obj = Nothing

End Sub




Public Sub execActiveAdd()

    Dim strBook As String
    Dim i As Long
    Dim lngCnt As Long
    
    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    strBook = ActiveWorkbook.FullName
    
    If Not rlxIsFileExists(strBook) Then
        MsgBox "ブックが存在しません。保存してから処理を行ってください。", vbOKOnly + vbExclamation, C_TITLE
        Exit Sub
    End If
    
    With lstFavorite
    
        For i = 0 To .ListCount - 1
        
            If .List(i, C_ORIGINAL) = strBook Then
                MsgBox "すでに同名のブック登録されています。", vbOKOnly + vbExclamation, C_TITLE
                Exit Sub
            End If
        
        Next
        
        
        lngCnt = .ListCount
        
        .AddItem ""
        .List(lngCnt, C_FILE_NO) = i + 1
        .List(lngCnt, C_FILE_NAME) = setFile(strBook)
        .List(lngCnt, C_PATH_NAME) = rlxGetFullpathFromPathName(strBook)
        .List(lngCnt, C_ORIGINAL) = strBook
        .List(lngCnt, C_CATEGORY) = lstCategory.List(lstCategory.ListIndex)

        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next
        
        .Selected(lngCnt) = True
        .ListIndex = lngCnt
        
        
    End With
    
    Call favCurrentUpdate
End Sub

Public Sub execDel()
    Dim i As Long
    Dim lngLast As Long
    
    If MsgBox("選択行を削除します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    With lstFavorite
    
        lngLast = .ListIndex
       
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                .RemoveItem i
            End If
        Next
        For i = 0 To .ListCount - 1
            .List(i, C_FILE_NO) = i + 1
        Next
    
'        setEnabled
        
        If lngLast > .ListCount - 1 Then
            lngLast = .ListCount - 1
        End If
    
        If .ListCount > 0 Then
            .Selected(lngLast) = True
        End If
        
    End With
    
    Call favCurrentUpdate
    
End Sub

Private Sub cmdAdd_Click()
    Call execActiveAdd
End Sub

Private Sub cmdDel_Click()
    Call execDel
End Sub

Private Sub chkDetail_Change()
    Call SaveSetting(C_TITLE, "Favirite", "Detail", chkDetail.Value)
End Sub

Private Sub lstCategory_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
'Cancel = True
End Sub

Private Sub lstCategory_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
        
'        Set mBarFavDrop = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
'        With mBarFavDrop
'
'            If lstCategory.ListCount <= 1 Then
'                With .Controls.Add
'                    .Caption = "移動できるカテゴリがありません"
'                End With
'            Else
'                Dim a As Variant
'                Dim i As Long
'                For i = 0 To lstCategory.ListCount - 1
'                    If i <> lstCategory.ListIndex Then
'                        With .Controls.Add
'                            .Caption = lstCategory.List(i)
'                            .OnAction = "'basFavorite.moveCategory(""" & lstCategory.List(i) & """)'"
'                            .FaceId = 526
'                        End With
'                    End If
'                Next
'            End If
'
'        End With
'        mBarFavDrop.ShowPopup
End Sub

Private Sub lstCategory_Change()
    Dim i As Long
    Dim c As Variant
    Dim fav As favoriteDTO
    Dim Key As String
    
    Dim blnFind As Boolean
    
    With frmFavorite
    
        .lstFavorite.Clear
        
        If .lstCategory.ListIndex < 0 Then
            Exit Sub
        End If
        
        .lstFavorite.Clear
        
        Key = .lstCategory.List(.lstCategory.ListIndex)
        
        If Not mobjCategory.Exists(Key) Then
            Exit Sub
        End If
        
        Dim cat As Variant
        Set cat = mobjCategory.Item(Key)
        
        i = 0
        For Each c In cat
        
            Set fav = cat.Item(c)
        
            .lstFavorite.AddItem ""
            .lstFavorite.List(i, C_FILE_NO) = i + 1
            .lstFavorite.List(i, C_FILE_NAME) = setFile(fav.filename)
            .lstFavorite.List(i, C_PATH_NAME) = rlxGetFullpathFromPathName(fav.filename)
            .lstFavorite.List(i, C_ORIGINAL) = fav.filename
            .lstFavorite.List(i, C_CATEGORY) = fav.Category
            i = i + 1
        
        Next
        
'        setEnabled
        
        blnFind = False
    
        If mstrFirstBook <> "" Then
            For i = 0 To .lstFavorite.ListCount - 1
                If LCase(mstrFirstBook) = LCase(.lstFavorite.List(i, C_ORIGINAL)) Then
                    .lstFavorite.Selected(i) = True
                    .lstFavorite.ListIndex = i
                    blnFind = True
                    Exit For
                End If
            Next
            mstrFirstBook = ""
        End If
        If blnFind = False Then
            If .lstFavorite.ListCount > 0 Then
                .lstFavorite.ListIndex = 0
                .lstFavorite.Selected(0) = True
            End If
        End If
    End With
    End Sub





Private Sub lstCategory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyV
            If (Shift And 2) Then
                Call favPaste
            End If
        Case vbKeyEscape
            Unload Me
        Case vbKeyRight
            lstFavorite.SetFocus
    End Select



End Sub



Private Sub lstFavorite_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''マウス左ボタンのドラッグ時に対応
'If Button <> 1 Then Exit Sub
''データオブジェクトに現在の選択地を格納
'    Dim D As DataObject
'    Set D = New DataObject
''D.SetText ListBox1.Value
'D.StartDrag 'ドラッグ開始

    Set MW.obj = lstFavorite


End Sub

Private Sub txtDetail_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyV
            If (Shift And 2) Then
                Call favPaste
            End If
        Case vbKeyEscape
            Unload Me
'        Case vbKeyReturn
'            execOpen False
        
    End Select
End Sub

Private Sub lstFavorite_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyV
            If (Shift And 2) Then
                Call favPaste
                Exit Sub
            End If
        Case vbKeyC
            If (Shift And 2) Then
                Call favCopy
                Exit Sub
            End If
        Case vbKeyA
            If (Shift And 2) Then
                Call favAllSelect
                Exit Sub
            End If
        Case vbKeyEscape
            Unload Me
            Exit Sub
        Case vbKeyReturn
            execOpen False
            Exit Sub
        Case vbKeyLeft
            lstCategory.SetFocus
            Exit Sub
        Case vbKeyDelete
            execDel
            Exit Sub
    End Select


'    If lstFavorite.ListIndex >= 0 Then
'        Dim i As Long
'        For i = 0 To lstFavorite.ListCount - 1
'            If i = lstFavorite.ListIndex Then
'                lstFavorite.Selected(i) = True
'            Else
'                lstFavorite.Selected(i) = False
'            End If
'        Next
'    End If

End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyV
            If (Shift And 2) Then
                Call favPaste
            End If
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            execOpen False
        
    End Select
End Sub
'Private Sub UserForm_Activate()
'    MW.Activate
'End Sub
Private Sub UserForm_Initialize()

    Dim strList() As String
    Dim strResult As String
    Dim i As Long
    Dim lngMax As Long
    
    Set RW = New ResizeWindow
    
    Set mBarCat = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
    With mBarCat
        With .Controls.Add
            .Caption = "先頭に移動"
            .OnAction = "'basFavorite.moveListCategoryFirst(""" & C_HEAD & """)'"
            .FaceId = 594
        End With
        With .Controls.Add
            .Caption = "1つ上に移動"
            .OnAction = "'basFavorite.moveListCategory(""" & C_UP & """)'"
            .FaceId = 595
        End With
        With .Controls.Add
            .BeginGroup = True
            .Caption = "1つ下に移動"
            .OnAction = "'basFavorite.moveListCategory(""" & C_DOWN & """)'"
            .FaceId = 596
        End With
        With .Controls.Add
            .Caption = "最後に移動"
            .OnAction = "'basFavorite.moveListCategoryFirst(""" & C_TAIL & """)'"
            .FaceId = 597
        End With
        With .Controls.Add
            .BeginGroup = True
            .Caption = "カテゴリの追加"
            .OnAction = "basFavorite.addCategory"
            .FaceId = 535
        End With
        With .Controls.Add
            .Caption = "カテゴリの変更"
            .OnAction = "basFavorite.modCategory"
            .FaceId = 534
        End With
        With .Controls.Add
            .Caption = "カテゴリの削除"
            .OnAction = "basFavorite.delCategory"
            .FaceId = 536
        End With
    End With
    
    Set mobjCategory = CreateObject("Scripting.Dictionary")
    
    
    mstrFirstBook = GetSetting(C_TITLE, "Favirite", "CurrentBook", "")
    
    strResult = GetSetting(C_TITLE, "Favirite", "FileList", "")
    strList = Split(strResult, vbVerticalTab)

    lngMax = UBound(strList)
    Dim fav As favoriteDTO
    Dim strDat() As String
    
    
    Dim strCategory As String
    
    Dim objfav As Variant
    
    strCategory = ""
    For i = 0 To lngMax
       
       Set fav = New favoriteDTO
    
       strDat = Split(strList(i), vbTab)
       
        Select Case True
            Case UBound(strDat) = 0
                fav.filename = strDat(0)
                fav.Category = C_FAV_ALL
'                fav.Text = rlxGetFullpathFromFileName(strDat(0))
                
            Case Else
                fav.filename = strDat(0)
                fav.Category = strDat(1)
'                fav.Text = rlxGetFullpathFromFileName(strDat(0))
                
'            Case UBound(strDat) = 2
'                fav.Filename = strDat(0)
'                fav.Category = strDat(1)
'                fav.Text = strDat(2)
        End Select
         
        If Not mobjCategory.Exists(fav.Category) Then
            Set objfav = CreateObject("Scripting.Dictionary")
            mobjCategory.Add fav.Category, objfav
       End If
       
       If objfav.Exists(fav.filename) Then
       Else
           objfav.Add fav.filename, fav
        End If
    Next

    If Not mobjCategory.Exists("Fast Pin") Then
        mobjCategory.Add "Fast Pin", CreateObject("Scripting.Dictionary")
    End If
    
    Dim cat As Variant
    For Each cat In mobjCategory.Keys
        lstCategory.AddItem cat
    Next
    
    If lstCategory.ListCount = 0 Then
        lstCategory.AddItem C_FAV_ALL
    End If
    
    Dim blnFind As Boolean
    blnFind = False
    strCategory = GetSetting(C_TITLE, "Favirite", "CurrentCategory", "")
    For i = 0 To lstCategory.ListCount - 1
        If lstCategory.List(i) = strCategory Then
            blnFind = True
            Exit For
        End If
    Next
    If blnFind Then
        lstCategory.ListIndex = i
    Else
        lstCategory.ListIndex = 0
    End If

    
'    cmdDel.Caption = "一覧から" & vbCrLf & "削除"
'    cmdAdd.Caption = "現在のブックを" & vbCrLf & "追加"
    
    RW.FormWidth = Me.width
    RW.FormHeight = Me.Height
    
'    mlngButtonLeft = Me.fraButton.Left
    mlngListWidth = Me.lstFavorite.width
    mlngListHeight = Me.lstFavorite.Height
    mlngLblBookWidth = Me.lblBook.width
    
    mlngDetailTop = Me.txtDetail.Top
    mlngDetailWidth = Me.txtDetail.width
    
    mlngLblTop = Me.lblMsg.Top
    mlngLblWidth = Me.lblMsg.width
    
'    Me.Top = GetSetting(C_TITLE, "Favirite", "Top", Application.Top + 20)
'    Me.Left = GetSetting(C_TITLE, "Favirite", "Left", Application.Left + 20)
'    Me.Width = GetSetting(C_TITLE, "Favirite", "Width", Me.Width)
'    Me.Height = GetSetting(C_TITLE, "Favirite", "Height", Me.Height)
    
    lblMsg.Caption = " 操作はリストを右クリック。一覧への追加はエクスプローラからのコピペ(CTRL+V)で可能です。Excelファイル以外のファイル、フォルダも追加可能です。"
    
    chkDetail.Value = GetSetting(C_TITLE, "Favirite", "Detail", False)
    
    Set MW = basMouseWheel.GetInstance
    MW.Install Me

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 移動処理
'------------------------------------------------------------------------------------------------------------------------
Public Sub moveList(ByVal lngMode As Long)

    Dim lngCnt As Long
    Dim lngCmp As Long
    
    Dim varTmp As Variant

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngInc As Long
    Dim lngSel As Long

    '１つなら不要
    If lstFavorite.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstFavorite.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstFavorite.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select
    lngSel = lngStart
    lngCmp = 0
    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstFavorite.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If

            lngCmp = lngCnt + lngInc * -1
            
            Dim i As Long
            For i = C_FILE_NAME To C_CATEGORY
                varTmp = lstFavorite.List(lngCnt, i)
                lstFavorite.List(lngCnt, i) = lstFavorite.List(lngCmp, i)
                lstFavorite.List(lngCmp, i) = varTmp
            Next
            
            lstFavorite.Selected(lngCnt) = False
            lstFavorite.Selected(lngCmp) = True
            If lstFavorite.ListIndex = lngCnt Then
                lstFavorite.ListIndex = lngCmp
            End If
        End If
    
    Next
    
    Call favCurrentUpdate

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 移動処理
'------------------------------------------------------------------------------------------------------------------------
Public Sub moveListFirst(ByVal lngMode As Long)

    Dim i As Long
    Dim j As Long
    Dim lngCmp As Long
    Dim lngSel As Long
    
    Dim varTmp As Variant
    Dim blnTmp As Boolean

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngInc As Long
    Dim lngDec As Long
    Dim lngLast As Long

    '１つなら不要
    If lstFavorite.ListCount <= 1 Then
        Exit Sub
    End If
    
    Select Case lngMode
        Case C_HEAD
            lngStart = 0
            lngEnd = lstFavorite.ListCount - 1
            lngInc = 1
            lngDec = -1
        Case C_TAIL
            lngStart = lstFavorite.ListCount - 1
            lngEnd = 0
            lngInc = -1
            lngDec = 1
    End Select
    
    lngLast = lngStart + lngInc
    For i = lngStart To lngEnd Step lngInc
    
        If lstFavorite.Selected(i) Then
        
            For j = i To lngLast Step lngDec
            
                lngSel = j
                lngCmp = j + lngDec
            
                Dim k As Long
                For k = C_FILE_NAME To C_CATEGORY
                    '値を交換
                    varTmp = lstFavorite.List(lngSel, k)
                    lstFavorite.List(lngSel, k) = lstFavorite.List(lngCmp, k)
                    lstFavorite.List(lngCmp, k) = varTmp
                Next
                
                '選択も交換
                blnTmp = lstFavorite.Selected(lngSel)
                lstFavorite.Selected(lngSel) = lstFavorite.Selected(lngCmp)
                lstFavorite.Selected(lngCmp) = blnTmp
                
                'カレント行も交換
                If lstFavorite.ListIndex = lngSel Then
                    lstFavorite.ListIndex = lngCmp
                End If
            
            Next
            
            lngLast = lngLast + lngInc
        
        End If
    
    Next
    
    Call favCurrentUpdate

End Sub

Sub favCurrentUpdate()

    Dim Key As String
    Dim i As Long
    Dim objfav As Variant
    Dim lngMax As Long
    Dim fav As favoriteDTO
    
    
    Key = lstCategory.List(lstCategory.ListIndex)
    If mobjCategory.Exists(Key) Then
        mobjCategory.remove Key
    End If
    
    Set objfav = CreateObject("Scripting.Dictionary")
    
    lngMax = lstFavorite.ListCount - 1
    
    For i = 0 To lngMax
       
        Set fav = New favoriteDTO
    
        fav.filename = lstFavorite.List(i, C_ORIGINAL)
        fav.Category = lstFavorite.List(i, C_CATEGORY)
'        fav.Text = lstFavorite.List(i, C_FILE_NAME)
       
        objfav.Add fav.filename, fav
              
    Next
    
    mobjCategory.Add Key, objfav
    
End Sub
Private Sub cmdReadOnly_Click()

    execOpen True

End Sub

Private Sub cmdSave_Click()

    Unload Me
    
End Sub

Private Sub cmdSelect_Click()
    execOpen False

End Sub
Public Sub execOpen(ByVal blnReadOnly As Boolean)

    Dim strBook As String
    Dim lngCnt As Long
    
    Dim varFile As Variant
    
    If lstFavorite.ListIndex = -1 Then
        Exit Sub
    End If
    
    On Error Resume Next
    Me.Hide
'    Application.ScreenUpdating = False
    For lngCnt = 0 To lstFavorite.ListCount - 1
    
        If lstFavorite.Selected(lngCnt) Then
    
            strBook = lstFavorite.List(lngCnt, C_ORIGINAL)
            
            Select Case True
                Case rlxIsExcelFile(strBook)
                    If Not rlxIsFileExists(strBook) Then
                        MsgBox "ブックが存在しません。", vbOKOnly + vbExclamation, C_TITLE
                    Else
                        On Error Resume Next
                        Err.Clear
                        Workbooks.Open filename:=strBook, ReadOnly:=blnReadOnly, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True
                        If Err.Number <> 0 Then
                            MsgBox "ブックを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
                        End If
                        AppActivate Application.Caption
                    End If
                    
                Case rlxIsPowerPointFile(strBook)
                    On Error Resume Next
                    Err.Clear
                    With CreateObject("PowerPoint.Application")
                        .visible = True
                        Call .Presentations.Open(filename:=strBook, ReadOnly:=blnReadOnly)
                        If Err.Number <> 0 Then
                            MsgBox "ファイルを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
                        End If
                        AppActivate .Caption
                    End With
                    
                Case rlxIsWordFile(strBook)
                    On Error Resume Next
                    Err.Clear
                    With CreateObject("Word.Application")
                        .visible = True
                        .Documents.Open filename:=strBook, ReadOnly:=blnReadOnly
                        If Err.Number <> 0 Then
                            MsgBox "ファイルを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
                        End If
                        AppActivate .Caption
                    End With
                    
                Case Else
                    On Error Resume Next
                    Dim WSH As Object
                    Set WSH = CreateObject("WScript.Shell")
                    
                    WSH.Run ("""" & strBook & """")
                     If Err.Number <> 0 Then
                        MsgBox "ファイルを開けませんでした。", vbOKOnly + vbExclamation, C_TITLE
                    End If
                    Set WSH = Nothing
                End Select
        End If
     
    Next
     
    Unload Me
'    Application.ScreenUpdating = True
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



Private Sub lstFavorite_Change()
    
    Application.OnTime Now, "lstFavoriteDispDetail"

End Sub
'Private Sub lstFavorite_Click()
'    Application.OnTime Now, "lstFavoriteDispDetail"
'End Sub
Public Sub lstFavoriteDispDetail()

    Dim strMsg As String
    Dim strBook As String
    Dim i As Long
    Dim strCat As String
    
    If lstFavorite.ListIndex = -1 Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    strMsg = C_FILE_INFO
        
    Dim lngCount As Long
    
    lngCount = 0
    For i = 0 To lstFavorite.ListCount - 1
        If lstFavorite.Selected(i) Then
            lngCount = lngCount + 1
        End If
    Next
    
    If lngCount = 1 Then
    
        Dim shell As Object, folder As Object
        
        strBook = lstFavorite.List(lstFavorite.ListIndex, C_ORIGINAL)
        
'        If rlxIsFileExists(strBook) Or rlxIsFolderExists(strBook) Then
        
                
            strMsg = strMsg & vbCrLf
            strMsg = strMsg & "　フォルダ名：" & rlxGetFullpathFromPathName(strBook) & vbCrLf           ''ファイル名
            strMsg = strMsg & "　ファイル名：" & rlxGetFullpathFromFileName(strBook) & vbCrLf           ''ファイル名
            
            If GetSetting(C_TITLE, "Favirite", "Detail", False) Then
                Set shell = CreateObject("Shell.Application")
                Set folder = shell.Namespace(rlxGetFullpathFromPathName(strBook))
                strMsg = strMsg & "　作成者：" & folder.GetDetailsOf(folder.ParseName(rlxGetFullpathFromFileName(strBook)), 20) & vbCrLf  ''作成者
                strMsg = strMsg & "　タイトル：" & folder.GetDetailsOf(folder.ParseName(rlxGetFullpathFromFileName(strBook)), 21) & vbCrLf   ''タイトル
                strMsg = strMsg & "　サブタイトル：" & folder.GetDetailsOf(folder.ParseName(rlxGetFullpathFromFileName(strBook)), 22) & vbCrLf   ''サブタイトル
                Set folder = Nothing
                Set shell = Nothing
            End If
            
            Select Case True
                Case rlxIsExcelFile(strBook)
                    strCat = "Excelファイル"
                Case rlxIsPowerPointFile(strBook)
                    strCat = "PowerPointファイル"
                Case rlxIsWordFile(strBook)
                    strCat = "Wordファイル"
                Case Else
                    strCat = "その他"
            End Select
            
            strMsg = strMsg & "　種類：" & strCat
            
'        Else
'            strMsg = strMsg & vbCrLf
'            strMsg = strMsg & "　フォルダまたはファイルがありません。" & vbCrLf
'        End If
    Else
    
        For i = 0 To lstFavorite.ListCount - 1
        
            If lstFavorite.Selected(i) Then
            
                strBook = lstFavorite.List(i, C_ORIGINAL)
                strMsg = strMsg & vbCrLf & "  " & strBook
                
                Select Case True
                    Case rlxIsExcelFile(strBook)
                        strCat = "Excelファイル"
                    Case rlxIsPowerPointFile(strBook)
                        strCat = "PowerPointファイル"
                    Case rlxIsWordFile(strBook)
                        strCat = "Wordファイル"
                    Case Else
                        strCat = "その他"
                End Select
                
                strMsg = strMsg & ", " & strCat
                
            End If
        Next
    End If
    
    txtDetail.Text = strMsg
End Sub

Private Sub lstFavorite_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call cmdSelect_Click
End Sub

Private Sub lstFavorite_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
    
        Set mBarFav = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
        With mBarFav
            With .Controls.Add
                .Caption = "開く"
                .OnAction = "'basFavorite.execOpen(""" & False & """)'"
                .FaceId = 23
            End With
            With .Controls.Add
                .Caption = "読み取り専用で開く"
                .OnAction = "'basFavorite.execOpen(""" & True & """)'"
                .FaceId = 456
            End With
            With .Controls.Add
                .Caption = "ファイルのあるフォルダを開く"
                .OnAction = "basFavorite.execOpenFolder"
                .FaceId = 23
            End With
            With .Controls.Add
                .BeginGroup = True
                .Caption = "先頭に移動"
                .OnAction = "'basFavorite.moveListFirst(""" & C_HEAD & """)'"
                .FaceId = 594
            End With
            With .Controls.Add
                .Caption = "1つ上に移動"
                .OnAction = "'basFavorite.moveList(""" & C_UP & """)'"
                .FaceId = 595
            End With
            With .Controls.Add
                .BeginGroup = True
                .Caption = "1つ下に移動"
                .OnAction = "'basFavorite.moveList(""" & C_DOWN & """)'"
                .FaceId = 596
            End With
            With .Controls.Add
                .Caption = "最後に移動"
                .OnAction = "'basFavorite.moveListFirst(""" & C_TAIL & """)'"
                .FaceId = 597
            End With
            
            With .Controls.Add
                .Caption = "アクティブブックを追加"
                .BeginGroup = True
                .OnAction = "basFavorite.execActiveAdd"
                .FaceId = 535
            End With
            
            With .Controls.Add
                .Caption = "追加"
'                .BeginGroup = True
                .OnAction = "basFavorite.execAdd"
                .FaceId = 535
            End With
            
            With .Controls.Add
                .Caption = "編集"
                .OnAction = "basFavorite.execEdit"
                .FaceId = 534
            End With
            

            With .Controls.Add
                .BeginGroup = True
                .Caption = "コピー"
                .OnAction = "basFavorite.favCopy"
                .FaceId = 19
            End With
            
            With .Controls.Add
                .Caption = "貼り付け"
                .OnAction = "basFavorite.favPaste"
                .FaceId = 1436
            End With
        
            With .Controls.Add
                .Caption = "削除"
                .OnAction = "basFavorite.execDel"
                .FaceId = 536
            End With
            
        
            Dim myCBCtrl2 As Variant
            Set myCBCtrl2 = .Controls.Add(Type:=msoControlPopup)
            With myCBCtrl2
                .Caption = "カテゴリー移動"
                .BeginGroup = True
            End With
        
            If lstCategory.ListCount <= 1 Then
                With myCBCtrl2.Controls.Add
                    .Caption = "移動できるカテゴリがありません"
                End With
            Else
                Dim a As Variant
                Dim i As Long
                For i = 0 To lstCategory.ListCount - 1
                    If i <> lstCategory.ListIndex Then
                        With myCBCtrl2.Controls.Add
                            .Caption = lstCategory.List(i)
                            .OnAction = "'basFavorite.moveCategory(""" & lstCategory.List(i) & """)'"
                            .FaceId = 526
                        End With
                    End If
                Next
            End If
        
        
        End With
        mBarFav.ShowPopup
    
    End If
End Sub
Private Sub lstCategory_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then mBarCat.ShowPopup
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveSetting C_TITLE, "Favirite", "Top", Me.Top
    SaveSetting C_TITLE, "Favirite", "Left", Me.Left
    SaveSetting C_TITLE, "Favirite", "Width", Me.width
    SaveSetting C_TITLE, "Favirite", "Height", Me.Height
    
End Sub

Private Sub UserForm_Resize()

'    On Error Resume Next
'
'    If RW.FormWidth > Me.Width Then
'        Me.Width = RW.FormWidth
'    End If
'    If RW.FormHeight > Me.Height Then
'        Me.Height = RW.FormHeight
'    End If
'
'    fraButton.Left = mlngButtonLeft + (Me.Width - RW.FormWidth)
'    lstFavorite.Width = mlngListWidth + (Me.Width - RW.FormWidth)
'    lstCategory.Height = mlngListHeight + (Me.Height - RW.FormHeight)
'    lstFavorite.Height = mlngListHeight + (Me.Height - RW.FormHeight)
'    lblBook.Width = mlngLblBookWidth + (Me.Width - RW.FormWidth)
''    lblMsg.Top = mlngLblTop + (Me.Height - RW.FormHeight)
'
'    Me.txtDetail.Top = mlngDetailTop + (Me.Height - RW.FormHeight)
'    Me.txtDetail.Width = mlngDetailWidth + (Me.Width - RW.FormWidth)
'
'    Me.lblMsg.Top = mlngLblTop + (Me.Height - RW.FormHeight)
'    Me.lblMsg.Width = mlngLblWidth + (Me.Width - RW.FormWidth)
'
'    DoEvents
    
End Sub

Private Sub UserForm_Terminate()

    MW.UnInstall
    Set MW = Nothing
        
    Dim key1 As Variant
    Dim key2 As Variant
    Dim cat As Variant
    Dim fav As favoriteDTO
    Dim strBuf As String
    Dim i As Long
    Dim blnFind As Boolean
    Dim j As Long
    
    blnFind = False
        
    If lstFavorite.ListIndex <> -1 Then
        SaveSetting C_TITLE, "Favirite", "CurrentBook", lstFavorite.List(lstFavorite.ListIndex, C_ORIGINAL)
    End If
    On Error Resume Next
    DeleteSetting C_TITLE, "FastPin"
    
    j = 0
    strBuf = ""
    For i = 0 To lstCategory.ListCount - 1
    
        key1 = lstCategory.List(i)
        
        If mobjCategory.Exists(key1) Then
        
            Set cat = mobjCategory.Item(key1)
            
            If cat.count = 0 And Not key1 = "Fast Pin" Then
                blnFind = True
            End If
            
            For Each key2 In cat
            
                Set fav = cat.Item(key2)
            
                If Len(strBuf) = 0 Then
                    strBuf = fav.filename & vbTab & key1
                Else
                    strBuf = strBuf & vbVerticalTab & fav.filename & vbTab & key1
                End If
                If key1 = "Fast Pin" Then
                    j = j + 1
                    SaveSetting C_TITLE, "FastPin", "runFastPin" & Format(j, "00"), fav.filename
                End If
                
            Next
        
        End If
    Next
    SaveSetting C_TITLE, "Favirite", "FileList", strBuf
    If lstCategory.ListIndex <> -1 Then
        SaveSetting C_TITLE, "Favirite", "CurrentCategory", lstCategory.List(lstCategory.ListIndex)
    End If
    
    Call RefreshRibbon
    
    If blnFind Then
        MsgBox "中身のないカテゴリは削除されます。", vbOKOnly + vbExclamation, C_TITLE
    End If
    
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 移動処理
'------------------------------------------------------------------------------------------------------------------------
Public Sub moveListCategory(ByVal lngMode As Long)

    Dim lngCnt As Long
    Dim lngCmp As Long
    
    Dim varTmp As Variant

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngInc As Long
    Dim lngSel As Long

    '１つなら不要
    If lstCategory.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstCategory.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstCategory.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select
    lngSel = lngStart
    lngCmp = 0
    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstCategory.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If
            
            lngCmp = lngCnt + lngInc * -1
            
            varTmp = lstCategory.List(lngCnt)
            lstCategory.List(lngCnt) = lstCategory.List(lngCmp)
            lstCategory.List(lngCmp) = varTmp
            
            lstCategory.Selected(lngCnt) = False
            lstCategory.Selected(lngCmp) = True
            If lstCategory.ListIndex = lngCnt Then
                lstCategory.ListIndex = lngCmp
            End If
        End If
    
    Next

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 移動処理
'------------------------------------------------------------------------------------------------------------------------
Public Sub moveListCategoryFirst(ByVal lngMode As Long)

    Dim i As Long
    Dim j As Long
    Dim lngCmp As Long
    Dim lngSel As Long
    
    Dim varTmp As Variant
    Dim blnTmp As Boolean

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngInc As Long
    Dim lngDec As Long
    Dim lngLast As Long

    '１つなら不要
    If lstCategory.ListCount <= 1 Then
        Exit Sub
    End If
    
    Select Case lngMode
        Case C_HEAD
            lngStart = 0
            lngEnd = lstCategory.ListCount - 1
            lngInc = 1
            lngDec = -1
        Case C_TAIL
            lngStart = lstCategory.ListCount - 1
            lngEnd = 0
            lngInc = -1
            lngDec = 1
    End Select
    
    lngLast = lngStart + lngInc
    For i = lngStart To lngEnd Step lngInc
    
        If lstCategory.Selected(i) Then
        
            For j = i To lngLast Step lngDec
            
                lngSel = j
                lngCmp = j + lngDec
            
                '値を交換
                varTmp = lstCategory.List(lngSel)
                lstCategory.List(lngSel) = lstCategory.List(lngCmp)
                lstCategory.List(lngCmp) = varTmp
                
                '選択も交換
                blnTmp = lstCategory.Selected(lngSel)
                lstCategory.Selected(lngSel) = lstCategory.Selected(lngCmp)
                lstCategory.Selected(lngCmp) = blnTmp
                
                'カレント行も交換
                If lstCategory.ListIndex = lngSel Then
                    lstCategory.ListIndex = lngCmp
                End If
            
            Next
            
            lngLast = lngLast + lngInc
        
        End If
    
    Next

End Sub
Sub moveCategory(ByVal strCategory As String)

    Dim i As Long
    Dim lngCmp As Long
    
    For i = 0 To lstFavorite.ListCount - 1
    
        If lstFavorite.Selected(i) Then
        
            Dim cat2 As Variant
            If mobjCategory.Exists(strCategory) Then
                Set cat2 = mobjCategory.Item(strCategory)
            Else
                Set cat2 = CreateObject("Scripting.Dictionary")
            End If
            Dim d As favoriteDTO
            
            Set d = New favoriteDTO
            d.filename = lstFavorite.List(i, C_ORIGINAL)
            d.Category = strCategory
'            d.Text = lstFavorite.List(i, C_FILE_NAME)

            If cat2.Exists(d.filename) Then
                Exit Sub
            End If
            
            cat2.Add d.filename, d
            If mobjCategory.Exists(strCategory) Then
                mobjCategory.remove strCategory
            End If
            mobjCategory.Add strCategory, cat2
        
        
            Dim cat As Variant
            Set cat = mobjCategory.Item(lstFavorite.List(i, C_CATEGORY))
            
            cat.remove lstFavorite.List(i, C_ORIGINAL)
        
        
        End If
    Next
    
    Call lstCategory_Change


End Sub
Sub favCopy()

    Dim strBuf() As String
    Dim i As Long
    Dim lngCnt As Long
    Dim strBook As String
    
    If lstFavorite.ListCount = 0 Then
        Exit Sub
    End If
    
    On Error Resume Next

    i = 1
    
    For lngCnt = 0 To lstFavorite.ListCount - 1
    
        If lstFavorite.Selected(lngCnt) Then
    
            strBook = lstFavorite.List(lngCnt, C_ORIGINAL)
            ReDim Preserve strBuf(1 To i)
            strBuf(i) = strBook
            
            i = i + 1
        End If
     
    Next

    Call SetCopyClipText(strBuf)
    

End Sub
Sub favPaste()

    Dim files As Variant
    Dim strBuf As String
    Dim lngPos As Long
    Dim i As Long
    Dim cb As DataObject
    
    strBuf = rlxGetFileNameFromCli()
    
    If strBuf = "" Then
    
        Set cb = New DataObject
    
        cb.GetFromClipboard
        strBuf = cb.getText()
        If strBuf = "" Then
            Exit Sub
        End If
        strBuf = Replace(strBuf, """", "")
        files = Split(strBuf, vbCrLf)
    Else
        files = Split(strBuf, vbCrLf)
    End If
    

    
    lngPos = lstFavorite.ListCount
    
    If lngPos > 0 Then
        For i = 0 To lngPos - 1
            lstFavorite.Selected(i) = False
        Next
    End If
    
    For i = LBound(files) To UBound(files) Step 1
    
    
        With lstFavorite
    
            Dim j As Long
            For j = 0 To .ListCount - 1
            
                'すでにあったら無視
                If .List(j, C_ORIGINAL) = files(i) Then
                    GoTo pass
                End If
            
            Next
        
            j = .ListCount
            
            .AddItem ""
            .List(j, C_FILE_NO) = j + 1
            .List(j, C_FILE_NAME) = setFile(files(i))
            .List(j, C_PATH_NAME) = rlxGetFullpathFromPathName(files(i))
            .List(j, C_ORIGINAL) = files(i)
            .List(j, C_CATEGORY) = lstCategory.List(lstCategory.ListIndex)
            .Selected(j) = True

        End With
pass:
    Next
    
    If lngPos < lstFavorite.ListCount Then
        lstFavorite.ListIndex = lngPos
    End If
    
    Call favCurrentUpdate
    
End Sub
Sub addCategory()
    Call frmFavCategory.Start(C_FAVORITE_ADD)
End Sub
Sub modCategory()
    Call frmFavCategory.Start(C_FAVORITE_MOD)
End Sub
Sub delCategory()
    Dim i As Long
    
    With lstCategory
        
        If .ListCount = 0 Then
            Exit Sub
        End If
        
        If .ListIndex < 0 Then
            Exit Sub
        End If
        
        If .ListCount = 1 Then
            MsgBox "カテゴリは少なくとも１つ必要です。削除できません。", vbOKOnly + vbExclamation, C_TITLE
            Exit Sub
        End If
        
        Dim Key As String
        
        If MsgBox("カテゴリとそれ以下のお気に入りを削除しますがよろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
        
        Key = lstCategory.List(lstCategory.ListIndex)
        
        If mobjCategory.Exists(Key) Then
            mobjCategory.remove Key
        End If
        
        'お気に入りをクリア
        lstFavorite.Clear
        
        .RemoveItem lstCategory.ListIndex
    
    End With
    
End Sub
Sub execOpenFolder()

    Dim strBook As String
    Dim lngCnt As Long
    
    For lngCnt = 0 To lstFavorite.ListCount - 1
    
        If lstFavorite.Selected(lngCnt) Then
    
            strBook = lstFavorite.List(lngCnt, C_PATH_NAME)
            
            Dim WSH As Object
            Set WSH = CreateObject("WScript.Shell")
            
            WSH.Run ("""" & strBook & """")
            
            Set WSH = Nothing

        End If
     
    Next
    Unload Me
     
End Sub
Sub execEdit()

    Dim strText As String
    Dim strFile As String
    Dim i As Long
    Dim lngCnt As Long
    
    For i = 0 To lstFavorite.ListCount - 1
        If lstFavorite.Selected(i) Then
            lngCnt = lngCnt + 1
        End If
    Next
    If lngCnt <> 1 Then
        MsgBox "リストを１つ選択してください。", vbExclamation + vbOKOnly, C_TITLE
        Exit Sub
    End If
    
    strFile = lstFavorite.List(lstFavorite.ListIndex, C_ORIGINAL)
    
    If frmFavEdit.Start(C_FAVORITE_MOD, strFile) = vbOK Then
    
        lstFavorite.List(lstFavorite.ListIndex, C_FILE_NAME) = setFile(strFile)
        lstFavorite.List(lstFavorite.ListIndex, C_PATH_NAME) = rlxGetFullpathFromPathName(strFile)
        lstFavorite.List(lstFavorite.ListIndex, C_ORIGINAL) = strFile

        Call favCurrentUpdate
    End If

End Sub
Sub execAdd()

    Dim strFile As String
    Dim lngCnt As Long
    Dim strCategory As String
    
    strFile = ""
    
    If frmFavEdit.Start(C_FAVORITE_ADD, strFile) = vbOK Then
    
        lstFavorite.AddItem ""
        lngCnt = lstFavorite.ListCount - 1
        
        If lstCategory.ListIndex = -1 Then
            strCategory = C_FAV_ALL
        Else
            strCategory = lstCategory.List(lstCategory.ListIndex)
        End If
        
        lstFavorite.List(lngCnt, C_FILE_NO) = lstFavorite.ListCount
        lstFavorite.List(lngCnt, C_FILE_NAME) = setFile(strFile)
        lstFavorite.List(lngCnt, C_PATH_NAME) = rlxGetFullpathFromPathName(strFile)
        lstFavorite.List(lngCnt, C_ORIGINAL) = strFile
        lstFavorite.List(lngCnt, C_CATEGORY) = strCategory
        
        Dim i As Long
        For i = 0 To lstFavorite.ListCount - 1
            lstFavorite.Selected(i) = False
        Next
        lstFavorite.Selected(lngCnt) = True
        lstFavorite.ListIndex = lngCnt
        
        Call favCurrentUpdate
        
    End If

End Sub
Function setFile(ByVal strBuf As String) As String

    Dim strLine As String
    
    strLine = rlxGetFullpathFromFileName(strBuf)

    If InStr(strLine, ".") = 0 Then
        setFile = "<" & strLine & ">"
    Else
        setFile = strLine
    End If
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
Sub favAllSelect()
    Dim i As Long

    For i = 0 To lstFavorite.ListCount - 1
        lstFavorite.Selected(i) = True
    Next
    
End Sub
