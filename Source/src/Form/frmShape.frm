VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShape 
   Caption         =   "カスタムシェイプ"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15765
   OleObjectBlob   =   "frmShape.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmShape"
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
'
'Option Explicit
'
'Private Const C_Text As String = 0
'Private Const C_Font As String = 1
'Private Const C_Color As String = 2
'Private Const C_Size As String = 3
'Private Const C_Line As String = 4
'Private Const C_StampType As String = 5
'Private Const C_File As String = 6
'Private Const C_LineSize As String = 7
'Private Const C_Round As String = 8
'Private Const C_Rotate As String = 9
'
'Private Const C_DATA As Long = 1
'
'Private Const C_UP As Long = 1
'Private Const C_DOWN As Long = 2
'
'Private mResult As VbMsgBoxResult
'Private mblnRefresh As Boolean
'
'Sub dispPreview()
'
'    Dim FS As Object
'    Dim strPath As String
'    Dim lngWidth As Long
'    Dim lngHeight As Long
'
'    If mblnRefresh Then
'        Exit Sub
'    End If
'
'    If Not IsNumeric(txtSize.Text) Then
'        Exit Sub
'    End If
'
'    If Not IsNumeric(txtLineSize.Text) Then
'        Exit Sub
'    End If
'
'    If Not IsNumeric(txtRound.Text) Then
'        Exit Sub
'    End If
'
'    strPath = rlxGetTempFolder() & C_STAMP_FILE_NAME & ".jpg"
'
'    Dim s As StampMitome
'
'    Set s = New StampMitome
'
'    Select Case True
'        Case optNormal.Value
'            s.StampType = C_STAMP_MITOME_NORMAL
'            s.Text = txtName.Text
'            Dim c As control
'            For Each c In Controls
'                Select Case c.Tag
'                    Case "N"
'                        c.enabled = True
'                    Case "F"
'                        c.enabled = False
'                End Select
'            Next
'        Case Else
'             s.StampType = C_STAMP_MITOME_FILE
'            s.Text = rlxGetFullpathFromFileName(txtFile.Text)
'            For Each c In Controls
'                Select Case c.Tag
'                    Case "N"
'                        c.enabled = False
'                    Case "F"
'                        c.enabled = True
'                End Select
'            Next
'    End Select
'
'    s.Font = cmbFont.Text
'
'    Select Case True
'        Case optLineSingle.Value
'            s.Line = C_STAMP_LINE_SINGLE
'        Case optLineDouble.Value
'            s.Line = C_STAMP_LINE_DOUBLE
'        Case Else
'            s.Line = C_STAMP_LINE_BOLD
'    End Select
'
'    Dim lngSize As Double
'    Select Case True
'        Case optLineSingle.Value, optLineBold.Value
'            lngSize = ThisWorkbook.Worksheets("stampEx").Shapes("shpMitome").Width
'
'        Case optLineDouble.Value
'            lngSize = ThisWorkbook.Worksheets("stampEx").Shapes("shpMitome").Width * 0.8
'
'    End Select
'    ThisWorkbook.Worksheets("stampEx").Shapes("shpMitome").Height = lngSize
'
'    s.Color = getHexColor(lblColor.BackColor)
'    s.Size = txtSize.Text
'    s.FilePath = txtFile.Text
'    s.LineSize = txtLineSize.Text
'    s.Round = txtRound.Text
'
'    Select Case True
'        Case optVertical.Value
'            s.Rotate = C_STAMP_ROTATE_VERTICAL
'        Case optHolizontal.Value
'            s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
'    End Select
'
'    '編集結果をリストに設定
'    Dim i As Long
'
'    i = lstStamp.ListIndex
'    If i = -1 Then
'        Exit Sub
'    End If
'
'    mblnRefresh = True
'
'    Dim strBuf As String
'    Dim varBuf() As Variant
'
'    ReDim varBuf(C_Text To C_Rotate)
'
'    varBuf(C_StampType) = s.StampType
'    varBuf(C_Text) = s.Text
'    varBuf(C_File) = s.FilePath
'    varBuf(C_Color) = s.Color
'    varBuf(C_Size) = s.Size
'    varBuf(C_Line) = s.Line
'    varBuf(C_Font) = s.Font
'    varBuf(C_LineSize) = s.LineSize
'    varBuf(C_Round) = s.Round
'    varBuf(C_Rotate) = s.Rotate
'
'    strBuf = Join(varBuf, vbTab)
'
'    lstStamp.List(i, C_Text) = s.Text
'    lstStamp.List(i, C_DATA) = strBuf
'
'    mblnRefresh = False
'
'    If s.StampType = C_STAMP_MITOME_NORMAL Then
'
'        Call editStampMitome(s, xlBitmap, lngHeight, lngWidth)
'
'    Else
'        imgPreview.Picture = LoadPicture("")
'
'        'ファイルが存在しない場合、
'        If Not rlxIsFileExists(txtFile.Text) Then
'           Exit Sub
'        End If
'
'        With ActiveSheet.Pictures.Insert(s.FilePath)
'            lngWidth = .Width
'            lngHeight = .Height
'            .CopyPicture
'            .Delete
'        End With
'    End If
'
'    With ThisWorkbook.Worksheets("stampEx").ChartObjects.Add(0, 0, lngWidth, lngHeight).Chart
'
'        .Paste
'        .ChartArea.Border.LineStyle = 0
'        .Export strPath, "JPG"
'
'        .Parent.Delete
'
'    End With
'
'    imgPreview.Picture = LoadPicture(strPath)
'
'End Sub
'Private Sub cmbFont_Click()
'    If cmbFont.ListIndex = -1 Then
'    Else
'        dispPreview
'    End If
'End Sub
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
''------------------------------------------------------------------------------------------------------------------------
'' 選択行を上に移動
''------------------------------------------------------------------------------------------------------------------------
'Private Sub cmdUp_Click()
'    mblnRefresh = True
'     Call moveList(C_UP)
'    mblnRefresh = False
'End Sub
''------------------------------------------------------------------------------------------------------------------------
'' 選択行を下に移動
''------------------------------------------------------------------------------------------------------------------------
'Private Sub cmdDown_Click()
'    mblnRefresh = True
'     Call moveList(C_DOWN)
'    mblnRefresh = False
'End Sub
''------------------------------------------------------------------------------------------------------------------------
'' 移動処理
''------------------------------------------------------------------------------------------------------------------------
'Private Sub moveList(ByVal lngMode As Long)
'
'    Dim lngCnt As Long
'    Dim lngCmp As Long
'
'    Dim varTmp As Variant
'
'    Dim lngStart As Long
'    Dim lngEnd As Long
'    Dim lngInc As Long
'
'    '１つなら不要
'    If lstStamp.ListCount <= 1 Then
'        Exit Sub
'    End If
'
'    Select Case lngMode
'        Case C_UP
'            lngStart = 0
'            lngEnd = lstStamp.ListCount - 1
'            lngInc = 1
'        Case C_DOWN
'            lngStart = lstStamp.ListCount - 1
'            lngEnd = 0
'            lngInc = -1
'    End Select
'
'    For lngCnt = lngStart To lngEnd Step lngInc
'
'        If lstStamp.Selected(lngCnt) Then
'            '選択された行がすでに開始行の場合移動不可
'            If lngCnt = lngStart Then
'                Exit For
'            End If
'
'            lngCmp = lngCnt + lngInc * -1
'
'            Dim i As Long
'            For i = C_Text To C_DATA
'                varTmp = lstStamp.List(lngCnt, i)
'                lstStamp.List(lngCnt, i) = lstStamp.List(lngCmp, i)
'                lstStamp.List(lngCmp, i) = varTmp
'            Next
'
'            lstStamp.Selected(lngCnt) = False
'            lstStamp.Selected(lngCnt + lngInc * -1) = True
'        End If
'
'    Next
'
'End Sub
'
'
'Private Sub cmdFile_Click()
'   Dim strFile As String
'
'
'    strFile = Application.GetOpenFilename("ファイル(*.*),(*.*)", , "画像ファイル", , False)
'    If strFile = "False" Then
'        'ファイル名が指定されなかった場合
'        Exit Sub
'    End If
'
'    txtFile.Text = strFile
'
'End Sub
'
'Private Sub cmdOK_Click()
'
'    Dim s As StampMitome
'    Dim col As Collection
'    Dim i As Long
'
'    Set col = New Collection
'    '設定情報取得
'
'    For i = 0 To lstStamp.ListCount - 1
'
'        Set s = New StampMitome
'
'        Dim varBuf As Variant
'        varBuf = Split(lstStamp.List(i, C_DATA), vbTab)
'
'        s.StampType = varBuf(C_StampType)
'        s.Text = varBuf(C_Text)
'        s.Font = varBuf(C_Font)
'        s.Color = varBuf(C_Color)
'        s.Size = varBuf(C_Size)
'        s.Line = varBuf(C_Line)
'        s.FilePath = varBuf(C_File)
'        s.LineSize = varBuf(C_LineSize)
'        s.Round = varBuf(C_Round)
'        s.Rotate = varBuf(C_Rotate)
'
'        If IsNumeric(s.Size) Then
'        Else
'            MsgBox "幅には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
'            lstStamp.Selected(i) = True
'            txtSize.SetFocus
'            Exit Sub
'        End If
'
'        If CDbl(s.Size) < 0 Then
'            MsgBox "幅は０以上を入力してください。", vbExclamation + vbOKOnly, C_TITLE
'            lstStamp.Selected(i) = True
'            txtSize.SetFocus
'            Exit Sub
'        End If
'
'        If IsNumeric(s.LineSize) Then
'        Else
'            MsgBox "外枠には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
'            lstStamp.Selected(i) = True
'            txtLineSize.SetFocus
'            Exit Sub
'        End If
'
'        If CDbl(s.LineSize) < 0 Then
'            MsgBox "外枠は０以上を入力してください。", vbExclamation + vbOKOnly, C_TITLE
'            lstStamp.Selected(i) = True
'            txtLineSize.SetFocus
'            Exit Sub
'        End If
'
'        If IsNumeric(s.Round) Then
'        Else
'            MsgBox "角丸には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
'            lstStamp.Selected(i) = True
'            txtRound.SetFocus
'            Exit Sub
'        End If
'
'        If CDbl(s.Round) < 0 Then
'            MsgBox "角丸は0.00～0.50を入力してください。", vbExclamation + vbOKOnly, C_TITLE
'            lstStamp.Selected(i) = True
'            txtRound.SetFocus
'            Exit Sub
'        End If
'
'        'ファイルの存在チェック
'        If s.StampType = C_STAMP_MITOME_FILE Then
'            If Not rlxIsFileExists(s.FilePath) Then
'                MsgBox "画像ファイルが存在しません。", vbExclamation + vbOKOnly, C_TITLE
'                lstStamp.Selected(i) = True
'                txtFile.SetFocus
'                Exit Sub
'            End If
'        End If
'
'        col.Add s
'
'        Set s = Nothing
'
'    Next
'
'    'プロパティ保存
'    setPropertyMitome col
'
'    Set col = Nothing
'
'    'リボンのリフレッシュ
'    Call RefreshRibbon
'
'    On Error GoTo 0
'
'    mResult = vbOK
'    Unload Me
'
'End Sub
'
'
'
'
'Private Sub optColorBlack_Click()
'    dispPreview
'End Sub
'
'Private Sub optColorRed_Click()
'    dispPreview
'
'End Sub
'
'
'
'Private Sub lblColor_Click()
'
'    Dim lngColor As Long
'    Dim result As VbMsgBoxResult
'
'    lngColor = lblColor.BackColor
'
'    result = frmColor.Start(lngColor)
'
'    If result = vbOK Then
'        lblColor.BackColor = lngColor
'        dispPreview
'    End If
'
'End Sub
'
'Private Sub lstStamp_Change()
'
'    Dim i As Long
'
'    If mblnRefresh Then
'        Exit Sub
'    End If
'
'    mblnRefresh = True
'
'    i = lstStamp.ListIndex
'    If i = -1 Then
'        Exit Sub
'    End If
'
'    Dim varBuf As Variant
'
'    varBuf = Split(lstStamp.List(i, C_DATA), vbTab)
'
'    Select Case varBuf(C_StampType)
'        Case C_STAMP_MITOME_NORMAL
'            optNormal.Value = True
'            txtName.Text = varBuf(C_Text)
'        Case Else
'            optFile.Value = True
'            txtName.Text = ""
'    End Select
'
'    Select Case varBuf(C_Line)
'        Case C_STAMP_LINE_SINGLE
'            optLineSingle.Value = True
'        Case C_STAMP_LINE_DOUBLE
'            optLineDouble.Value = True
'        Case Else
'            optLineBold.Value = True
'    End Select
'
'
'    Dim lngColor As Long
'    lngColor = getLongColor(varBuf(C_Color))
'    lblColor.BackColor = lngColor
'
'    txtSize.Text = varBuf(C_Size)
'    txtFile.Text = varBuf(C_File)
'    txtLineSize.Text = varBuf(C_LineSize)
'    txtRound.Text = varBuf(C_Round)
'
'    Select Case varBuf(C_Rotate)
'        Case C_STAMP_ROTATE_HOLIZONTAL
'            optHolizontal.Value = True
'        Case C_STAMP_ROTATE_VERTICAL
'            optVertical.Value = True
'    End Select
'
'    Dim strFont As String
'    Dim pos As Long
'
'    strFont = varBuf(C_Font)
'
'    For i = 0 To cmbFont.ListCount - 1
'        If strFont = cmbFont.List(i) Then
'            pos = i
'        End If
'    Next i
'    cmbFont.ListIndex = pos
'
'    mblnRefresh = False
'
'    dispPreview
'
'End Sub
'
'Private Sub optFile_Change()
'    dispPreview
'End Sub
'
'Private Sub optHolizontal_Click()
'    dispPreview
'End Sub
'
'Private Sub optLineBold_Change()
'    dispPreview
'End Sub
'
'
'Private Sub optLineDouble_Click()
'    dispPreview
'End Sub
'
'Private Sub optLineSingle_Click()
'    dispPreview
'End Sub
'Private Sub optNormal_Change()
'    dispPreview
'End Sub
'
'Private Sub optVertical_Click()
'    dispPreview
'End Sub
'
'Private Sub spnLine_SpinUp()
'    txtLineSize.Text = spinUpSize(txtLineSize.Text)
'End Sub
'
'Private Sub spnLine_Spindown()
'    txtLineSize.Text = spinDownSize(txtLineSize.Text)
'End Sub
'
'Private Sub spnRound_SpinDown()
'    txtRound.Text = spinDownRound(txtRound.Text)
'End Sub
'
'Private Sub spnRound_SpinUp()
'    txtRound.Text = spinUpRound(txtRound.Text)
'End Sub
'
'Private Sub spnSize_SpinDown()
'    txtSize.Text = spinDown(txtSize.Text)
'End Sub
'
'Private Sub spnSize_SpinUp()
'    txtSize.Text = spinUp(txtSize.Text)
'End Sub
'
'Private Sub txtFile_Change()
'    dispPreview
'End Sub
'
'Private Sub txtLineSize_Change()
'    dispPreview
'End Sub
'
'Private Sub txtName_Change()
'
'    dispPreview
'
'End Sub
'
'Private Sub txtName_Enter()
'
'    dispPreview
'
'End Sub
'
'Private Sub txtRound_Change()
'    dispPreview
'End Sub
'
'Private Sub txtSize_Change()
'    dispPreview
'End Sub
'
'Private Sub UserForm_Initialize()
'
'    Dim s As StampMitome
'    Dim col As Collection
'    Dim i As Long
'
'    Dim strBuf As String
'    Dim varBuf() As Variant
'
'    ReDim varBuf(C_Text To C_Rotate)
'
'    '設定情報取得
'    Set col = getPropertyMitome()
'
'    For i = 1 To col.count
'
'        Set s = col(i)
'
'        varBuf(C_StampType) = s.StampType
'        varBuf(C_Text) = s.Text
'        varBuf(C_Font) = s.Font
'        varBuf(C_Color) = s.Color
'        varBuf(C_Size) = s.Size
'        varBuf(C_Line) = s.Line
'        varBuf(C_File) = s.FilePath
'        varBuf(C_LineSize) = s.LineSize
'        varBuf(C_Round) = s.Round
'        varBuf(C_Rotate) = s.Rotate
'
'        strBuf = Join(varBuf, vbTab)
'
'        lstStamp.AddItem ""
'        lstStamp.List(i - 1, C_Text) = s.Text
'        lstStamp.List(i - 1, C_DATA) = strBuf
'
'    Next
'
'    ActiveCell.Select
'    With Application.CommandBars("Formatting").Controls(1)
'        For i = 1 To .ListCount
'            cmbFont.AddItem .List(i)
'        Next i
'    End With
'
'
'    If col.count > 0 Then
'        lstStamp.Selected(0) = True
'    Else
'
'        mblnRefresh = True
'
'        txtName.Text = ""
'        txtFile.Text = ""
'        cmbFont.Text = "ＭＳ ゴシック"
'        optLineSingle.Value = True
'        optNormal.Value = True
'        txtSize.Text = "10.5"
'        lblColor.BackColor = vbRed
'        txtLineSize.Text = "10"
'        txtRound.Text = "0.50"
'        optVertical.Value = True
'
'        mblnRefresh = False
'
'    End If
'
'
'End Sub
'Private Function spinUp(ByVal vntValue As Variant) As Variant
'
'    Dim lngValue As Variant
'
'    lngValue = Val(vntValue)
'    lngValue = lngValue + 0.5
'
'    spinUp = lngValue
'
'End Function
'
'Private Function spinDown(ByVal vntValue As Variant) As Variant
'
'    Dim lngValue As Variant
'
'    lngValue = Val(vntValue)
'    lngValue = lngValue - 0.5
'    If lngValue < 10.5 Then
'        lngValue = 10.5
'    End If
'    spinDown = lngValue
'
'End Function
'
'Private Sub cmdAdd_Click()
'
'    Dim i As Long
'    Dim strBuf As String
'    Dim varBuf() As Variant
'
'    ReDim varBuf(C_Text To C_Rotate)
'
'    i = lstStamp.ListCount
'
'
'    Select Case True
'        Case optNormal.Value
'           varBuf(C_StampType) = C_STAMP_MITOME_NORMAL
'        Case Else
'            varBuf(C_StampType) = C_STAMP_MITOME_FILE
'    End Select
'
'    varBuf(C_Text) = txtName.Text
'
'    varBuf(C_Color) = getHexColor(lblColor.BackColor)
'
'    varBuf(C_Size) = txtSize.Text
'
'    Select Case True
'        Case optLineSingle.Value
'            varBuf(C_Line) = C_STAMP_LINE_SINGLE
'        Case optLineDouble.Value
'            varBuf(C_Line) = C_STAMP_LINE_DOUBLE
'        Case Else
'            varBuf(C_Line) = C_STAMP_LINE_BOLD
'    End Select
'
'    varBuf(C_Font) = cmbFont.Text
'    varBuf(C_File) = txtFile.Text
'
'    varBuf(C_LineSize) = txtLineSize.Text
'    varBuf(C_Round) = txtRound.Text
'
'
'    strBuf = Join(varBuf, vbTab)
'
'    lstStamp.AddItem ""
'    lstStamp.List(i, C_Text) = txtName.Text
'    lstStamp.List(i, C_DATA) = strBuf
'
'    lstStamp.Selected(i) = True
'
'End Sub
'
'
'Private Sub cmdDel_Click()
'    Dim i As Long
'
'    For i = 0 To lstStamp.ListCount
'
'        If lstStamp.Selected(i) Then
'            lstStamp.RemoveItem i
'            Exit Sub
'        End If
'
'    Next
'End Sub
'
'Private Function spinUpRound(ByVal vntValue As Variant) As Variant
'
'    Dim lngValue As Variant
'
'    lngValue = Val(vntValue)
'    lngValue = lngValue + 0.01
'    If lngValue > 0.5 Then
'        lngValue = 0.5
'    End If
'    spinUpRound = Format(lngValue, "0.00")
'
'End Function
'
'Private Function spinDownRound(ByVal vntValue As Variant) As Variant
'
'    Dim lngValue As Variant
'
'    lngValue = Val(vntValue)
'    lngValue = lngValue - 0.01
'    If lngValue < 0 Then
'        lngValue = 0
'    End If
'    spinDownRound = Format(lngValue, "0.00")
'
'End Function
'Private Function spinUpSize(ByVal vntValue As Variant) As Variant
'
'    Dim lngValue As Variant
'
'    lngValue = Val(vntValue)
'    lngValue = lngValue + 1
'
'    spinUpSize = Format(lngValue, "0")
'
'End Function
'
'Private Function spinDownSize(ByVal vntValue As Variant) As Variant
'
'    Dim lngValue As Variant
'
'    lngValue = Val(vntValue)
'    lngValue = lngValue - 1
'    If lngValue < 0 Then
'        lngValue = 0
'    End If
'    spinDownSize = Format(lngValue, "0")
'
'End Function
Private Sub cmdHelp_Click()

    If Val(Application.Version) >= C_EXCEL_VERSION_2013 Then
    
        If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
            Exit Sub
        End If
        
        Dim WSH As Object
        
        Set WSH = CreateObject("WScript.Shell")
        
        Call WSH.Run(C_STAMP_URL)
        
        Set WSH = Nothing
    
    Else
        frmHelp.Start "format"
    End If
    
End Sub

