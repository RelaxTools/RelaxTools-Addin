VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStampBz 
   Caption         =   "ビジネス印"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   OleObjectBlob   =   "frmStampBz.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmStampBz"
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


Private Const C_Text As Long = 0
Private Const C_StampType As Long = 1
Private Const C_DateType As Long = 2
Private Const C_DateFormat As Long = 3
Private Const C_UserDate As Long = 4
Private Const C_Font As Long = 5
Private Const C_Color As Long = 6
Private Const C_SIZE As Long = 7
Private Const C_Round As Long = 8
Private Const C_Rotate As Long = 9
Private Const C_LineSize As Long = 10

Private Const C_DATA As Long = 1

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2

Private mResult As VbMsgBoxResult

Private mblnRefresh As Boolean

Private Sub cmbFont_Click()
    If cmbFont.ListIndex = -1 Then
    Else
        dispPreview
    End If
End Sub

Private Sub cmdAdd_Click()

    Dim i As Long
    
    i = lstStampBz.ListCount

    lstStampBz.AddItem ""

'    Select Case True
'        Case optRectangle.Value
'            lstStampBz.List(i, C_StampType) = "1"
'        Case optSquare.Value
'            lstStampBz.List(i, C_StampType) = "2"
'        Case optCircle.Value
'            lstStampBz.List(i, C_StampType) = "3"
'    End Select
'
'    lstStampBz.List(i, C_Text) = txtText.Text
'
'    Select Case True
'        Case optSystemDate.Value
'            lstStampBz.List(i, C_DateType) = C_STAMP_DATE_SYSTEM
'        Case optUserDate.Value
'            lstStampBz.List(i, C_DateType) = C_STAMP_DATE_USER
'    End Select
'
'    lstStampBz.List(i, C_DateFormat) = txtFormat.Text
'
'    lstStampBz.List(i, C_UserDate) = txtUserDate.Text
'
'    lstStampBz.List(i, C_Color) = "&H" & Right$("00000000" & Hex(lblColor.BackColor), 8)
'
'    lstStampBz.List(i, C_Size) = txtWidth.Text
'
'    lstStampBz.List(i, C_Round) = txtRound.Text
'
'    lstStampBz.List(i, C_Font) = cmbFont.Text
'
'    lstStampBz.List(i, C_LineSize) = txtLineSize.Text


    Dim varBuf() As Variant
    Dim strBuf As String
    ReDim varBuf(C_Text To C_LineSize)
    
    Select Case True
        Case optRectangle.value
            varBuf(C_StampType) = "1"
        Case optSquare.value
            varBuf(C_StampType) = "2"
        Case optCircle.value
            varBuf(C_StampType) = "3"
    End Select
    
    varBuf(C_Text) = txtText.Text
    
    Select Case True
        Case optSystemDate.value
            varBuf(C_DateType) = C_STAMP_DATE_SYSTEM
        Case optUserDate.value
            varBuf(C_DateType) = C_STAMP_DATE_USER
    End Select
    
    varBuf(C_DateFormat) = txtFormat.Text
    
    varBuf(C_UserDate) = txtUserDate.Text
    
    varBuf(C_Color) = "&H" & Right$("00000000" & Hex(lblColor.BackColor), 8)
    
    varBuf(C_SIZE) = txtWidth.Text
    
    varBuf(C_Round) = txtRound.Text
    
    varBuf(C_Font) = cmbFont.Text
    
    varBuf(C_LineSize) = txtLineSize.Text


    strBuf = Join(varBuf, vbTab)
    
    lstStampBz.List(i, C_Text) = txtText.Text
    lstStampBz.List(i, C_DATA) = strBuf
        
    
    lstStampBz.Selected(i) = True
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()

    Dim i As Long
    
    For i = 0 To lstStampBz.ListCount
    
        If lstStampBz.Selected(i) Then
            lstStampBz.RemoveItem i
            Exit Sub
        End If
    
    Next
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 選択行を上に移動
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdUp_Click()
    mblnRefresh = True
     Call moveList(C_UP)
    mblnRefresh = False
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 選択行を下に移動
'------------------------------------------------------------------------------------------------------------------------
Private Sub cmdDown_Click()
    mblnRefresh = True
     Call moveList(C_DOWN)
    mblnRefresh = False
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
    If lstStampBz.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstStampBz.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstStampBz.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select

    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstStampBz.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If
            
            lngCmp = lngCnt + lngInc * -1
            
            Dim i As Long
            For i = C_Text To C_DATA
                varTmp = lstStampBz.List(lngCnt, i)
                lstStampBz.List(lngCnt, i) = lstStampBz.List(lngCmp, i)
                lstStampBz.List(lngCmp, i) = varTmp
            Next
            
            lstStampBz.Selected(lngCnt) = False
            lstStampBz.Selected(lngCnt + lngInc * -1) = True
        End If
    
    Next

End Sub

Private Sub cmdHelp_Click()
        
    If MsgBox("インターネットに接続します。よろしいですか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
        Exit Sub
    End If
    
    Dim WSH As Object
    
    Set WSH = CreateObject("WScript.Shell")
    
    Call WSH.Run(C_STAMP_URL)
    
    Set WSH = Nothing

End Sub
Sub dispPreview()

'    Dim FS As Object
'    Dim strPath As String
'    Dim lngWidth As Long
'    Dim lngHeight As Long
'    Dim lngColor As Long
'    Dim lngStyle As Long
'    Dim sngWeight As Single
'    Dim strText As String
'
'    Dim strSheet As String
'
'    Select Case True
'        Case optLong.Value
'            strSheet = "stampBz1"
'        Case optSquare.Value
'            strSheet = "stampBz2"
'        Case optCircle.Value
'            strSheet = "stampBz3"
'    End Select
'
'    If mblnRefresh Then
'        Exit Sub
'    End If
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'
'    strPath = rlxAddFileSeparator(FS.GetSpecialFolder(2)) & "\bzStamp.jpg"
'
'    Set FS = Nothing
'
'    Dim strFormat As String
'    Dim strType As String
'    Dim strUserDate As String
'
'    strFormat = txtFormat.Text
'    Select Case True
'        Case optSystemDate.Value
'            strType = C_STAMP_DATE_SYSTEM
'        Case optUserDate.Value
'            strType = C_STAMP_DATE_USER
'    End Select
'    strUserDate = txtUserDate.Text
'
'    lngColor = lblColor.BackColor
'
'    strText = txtText.Text
'
'    Dim strDate As String
'    strDate = getFormatDate(strFormat, strType, strUserDate)
'
'    strText = Replace(strText, "$d", strDate)
'
'    If InStr(strText, vbCrLf) = 0 Then
'
'        With ThisWorkbook.Worksheets(strSheet).Shapes("shpSquMid").TextFrame2.TextRange
'            .Font.NameComplexScript = cmbFont.Text
'            .Font.NameFarEast = cmbFont.Text
'            .Font.Name = cmbFont.Text
'            .Font.Fill.ForeColor.RGB = lngColor
'            .Text = strText
'        End With
'
'        With ThisWorkbook.Worksheets(strSheet).Shapes("shpSquUp").TextFrame2.TextRange
'            .Font.NameComplexScript = cmbFont.Text
'            .Font.NameFarEast = cmbFont.Text
'            .Font.Name = cmbFont.Text
'            .Font.Fill.ForeColor.RGB = lngColor
'            .Text = ""
'        End With
'
'        With ThisWorkbook.Worksheets(strSheet).Shapes("shpSquDown").TextFrame2.TextRange
'            .Font.NameComplexScript = cmbFont.Text
'            .Font.NameFarEast = cmbFont.Text
'            .Font.Name = cmbFont.Text
'            .Font.Fill.ForeColor.RGB = lngColor
'            .Text = ""
'        End With
'    Else
'
'        Dim strHigh As String
'        Dim strLow As String
'        Dim lngPos As Long
'
'        lngPos = InStr(strText, vbCrLf)
'
'        strHigh = Mid$(strText, 1, lngPos - 1)
'        strLow = Mid$(strText, lngPos + 2)
'
'        With ThisWorkbook.Worksheets(strSheet).Shapes("shpSquMid").TextFrame2.TextRange
'            .Font.NameComplexScript = cmbFont.Text
'            .Font.NameFarEast = cmbFont.Text
'            .Font.Name = cmbFont.Text
'            .Font.Fill.ForeColor.RGB = lngColor
'            .Text = ""
'        End With
'
'        With ThisWorkbook.Worksheets(strSheet).Shapes("shpSquUp").TextFrame2.TextRange
'            .Font.NameComplexScript = cmbFont.Text
'            .Font.NameFarEast = cmbFont.Text
'            .Font.Name = cmbFont.Text
'            .Font.Fill.ForeColor.RGB = lngColor
'            .Text = strHigh
'        End With
'
'        With ThisWorkbook.Worksheets(strSheet).Shapes("shpSquDown").TextFrame2.TextRange
'            .Font.NameComplexScript = cmbFont.Text
'            .Font.NameFarEast = cmbFont.Text
'            .Font.Name = cmbFont.Text
'            .Font.Fill.ForeColor.RGB = lngColor
'            .Text = strLow
'        End With
'    End If
'
'
'
'    Select Case True
'        Case optLong.Value, optSquare.Value
'            ThisWorkbook.Worksheets(strSheet).Shapes("shpSquare").Adjustments.Item(1) = CDbl(txtRound.Text)
'    End Select
'
'
'    ThisWorkbook.Worksheets(strSheet).Shapes("grpSquare").Line.ForeColor.RGB = lngColor
''    ThisWorkbook.Worksheets("stampEx").Shapes("shpSquare").Line.Weight = sngWeight
'
'    ThisWorkbook.Worksheets(strSheet).Shapes("grpSquare").CopyPicture xlScreen, xlBitmap
'
'    lngWidth = ThisWorkbook.Worksheets(strSheet).Shapes("grpSquare").Width
'    lngHeight = ThisWorkbook.Worksheets(strSheet).Shapes("grpSquare").Height
'
'
'
'    With ThisWorkbook.Worksheets(strSheet).ChartObjects.Add(0, 0, lngWidth, lngHeight).Chart
'
'        .Paste
'        .ChartArea.Border.LineStyle = 0
'        .Export strPath, "JPG"
'
'        .Parent.Delete
'
'    End With
        
        
    '設定情報取得
'    Dim lngWidth As Long
'    Dim lngHeight As Long
    Dim strPath As String
    
    Dim col As Collection
'    Dim r As Worksheet
    Dim bz As StampBzDTO
    
    If mblnRefresh Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtWidth.Text) Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtLineSize.Text) Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtRound.Text) Then
        Exit Sub
    End If
    
    Set bz = New StampBzDTO
    
    Select Case True
        Case optRectangle.value
            bz.StampType = C_STAMP_BZ_RECTANGLE
        Case optSquare.value
            bz.StampType = C_STAMP_BZ_SQUARE
        Case optCircle.value
            bz.StampType = C_STAMP_BZ_CIRCLE
    End Select
    
    Select Case True
        Case optHolizontal.value
            bz.Rotate = C_STAMP_ROTATE_HOLIZONTAL
        Case optVertical.value
            bz.Rotate = C_STAMP_ROTATE_VERTICAL
    End Select
    
    bz.Text = txtText.Text
    
    bz.DateFormat = txtFormat.Text
    Select Case True
        Case optSystemDate.value
            bz.DateType = C_STAMP_DATE_SYSTEM
        Case optUserDate.value
            bz.DateType = C_STAMP_DATE_USER
    End Select
    bz.UserDate = txtUserDate.Text
    
    bz.Color = "&H" & Right$("00000000" & Hex(lblColor.BackColor), 8)
    bz.Font = cmbFont.Text
    bz.Round = txtRound.Text
    bz.Size = txtWidth.Text
    
    bz.LineSize = txtLineSize.Text
    
    
'    Select Case True
'        Case optHolizontal.Value
'            Call editStampBz(bz, xlBitmap)
'        Case optVertical.Value
'            Call editStampBz(bz, xlBitmap)
'    End Select
    imgPreview.Picture = editStampBz(bz, xlBitmap)
'    imgPreview.Picture = CreatePictureFromClipboard()

    '編集結果をリストに設定
    Dim i As Long

    i = lstStampBz.ListIndex
    If i = -1 Then
        Exit Sub
    End If
    
    mblnRefresh = True
    
    Dim varBuf() As Variant
    Dim strBuf As String
    ReDim varBuf(C_Text To C_LineSize)
    
    Select Case True
        Case optRectangle.value
            varBuf(C_StampType) = C_STAMP_BZ_RECTANGLE
        Case optSquare.value
            varBuf(C_StampType) = C_STAMP_BZ_SQUARE
        Case optCircle.value
            varBuf(C_StampType) = C_STAMP_BZ_CIRCLE
    End Select
    
    Select Case True
        Case optHolizontal.value
            varBuf(C_Rotate) = C_STAMP_ROTATE_HOLIZONTAL
        Case optVertical.value
            varBuf(C_Rotate) = C_STAMP_ROTATE_VERTICAL
    End Select
    
   varBuf(C_Text) = txtText.Text
    
    Select Case True
        Case optSystemDate.value
           varBuf(C_DateType) = C_STAMP_DATE_SYSTEM
        Case optUserDate.value
            varBuf(C_DateType) = C_STAMP_DATE_USER
    End Select
    
    varBuf(C_DateFormat) = txtFormat.Text
    
    varBuf(C_UserDate) = txtUserDate.Text
    
    varBuf(C_Color) = "&H" & Right$("00000000" & Hex(lblColor.BackColor), 8)
    
    varBuf(C_SIZE) = txtWidth.Text
    
    varBuf(C_Round) = txtRound.Text
    
    varBuf(C_Font) = cmbFont.Text

    varBuf(C_LineSize) = txtLineSize.Text
    
    strBuf = Join(varBuf, vbTab)
    lstStampBz.List(i, C_Text) = txtText.Text
    lstStampBz.List(i, C_DATA) = strBuf
    
    mblnRefresh = False

End Sub



Private Sub cmdOK_Click()
        
    Dim datStampBz As StampBzDTO
    Dim col As Collection
    Dim i As Long
    
    Dim varBuf As Variant


    Set col = New Collection
    '設定情報取得

    For i = 0 To lstStampBz.ListCount - 1
        
        Set datStampBz = New StampBzDTO
        
        varBuf = Split(lstStampBz.List(i, C_DATA), vbTab)
        
        datStampBz.StampType = varBuf(C_StampType)
        datStampBz.Text = varBuf(C_Text)
        datStampBz.DateType = varBuf(C_DateType)
        datStampBz.DateFormat = varBuf(C_DateFormat)
        datStampBz.UserDate = varBuf(C_UserDate)
        datStampBz.Font = varBuf(C_Font)
        datStampBz.Color = varBuf(C_Color)
        datStampBz.Size = varBuf(C_SIZE)
        datStampBz.Round = varBuf(C_Round)
        datStampBz.Rotate = varBuf(C_Rotate)
        datStampBz.LineSize = varBuf(C_LineSize)
        
        
        If datStampBz.DateType = C_STAMP_DATE_USER Then
            If IsDate(datStampBz.UserDate) Then
            Else
                MsgBox "指定日付には有効な日付をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
                lstStampBz.Selected(i) = True
                txtUserDate.SetFocus
                Exit Sub
            End If
        End If
        
        If IsNumeric(datStampBz.Size) Then
        Else
            MsgBox "幅には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStampBz.Selected(i) = True
            txtWidth.SetFocus
            Exit Sub
        End If
        
        If CDbl(datStampBz.Size) < 0 Then
            MsgBox "幅は０以上を入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStampBz.Selected(i) = True
            txtWidth.SetFocus
            Exit Sub
        End If
        
        If IsNumeric(datStampBz.LineSize) Then
        Else
            MsgBox "外枠には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStampBz.Selected(i) = True
            txtLineSize.SetFocus
            Exit Sub
        End If
        
        If CDbl(datStampBz.LineSize) < 0 Then
            MsgBox "外枠は０以上を入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStampBz.Selected(i) = True
            txtLineSize.SetFocus
            Exit Sub
        End If
        
        If IsNumeric(datStampBz.Round) Then
        Else
            MsgBox "角丸には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStampBz.Selected(i) = True
            txtRound.SetFocus
            Exit Sub
        End If
        
        Select Case CDbl(datStampBz.Round)
            Case 0 To 0.5
            Case Else
                MsgBox "角丸は０～０．５の間で入力してください。", vbExclamation + vbOKOnly, C_TITLE
                lstStampBz.Selected(i) = True
                txtRound.SetFocus
                Exit Sub
        End Select
        
        col.Add datStampBz
        
        Set datStampBz = Nothing
        
    Next

    'プロパティ保存
    setPropertyBz col
    
    Set col = Nothing
    
    'リボンのリフレッシュ
    Call RefreshRibbon
    
    Unload Me
    
End Sub
Private Sub optColorBlack_Click()
    dispPreview
End Sub

Private Sub optColorBlue_Click()
    dispPreview
End Sub

Private Sub optColorRed_Click()
    dispPreview

End Sub



Private Sub lblColor_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult
    
    lngColor = lblColor.BackColor
    
    Result = frmColor.Start(lngColor)
    
    If Result = vbOK Then
        lblColor.BackColor = lngColor
        dispPreview
    End If
    
    
End Sub

Private Sub optLineBold_Click()
    dispPreview

End Sub

Private Sub optLineDouble_Click()
    dispPreview

End Sub

Private Sub optLineSingle_Click()
    dispPreview

End Sub



Private Sub lstStampBz_Click()

    Dim i As Long
    
    If mblnRefresh Then
        Exit Sub
    End If

    mblnRefresh = True

    i = lstStampBz.ListIndex
    If i = -1 Then
        Exit Sub
    End If
    
    Dim varBuf As Variant
    
    varBuf = Split(lstStampBz.List(i, C_DATA), vbTab)

    Select Case varBuf(C_StampType)
        Case C_STAMP_BZ_RECTANGLE
            optRectangle.value = True
        Case C_STAMP_BZ_SQUARE
            optSquare.value = True
        Case C_STAMP_BZ_CIRCLE
            optCircle.value = True
'        Case C_STAMP_BZ_SAKURA
'            optSakura.Value = True
    End Select
    
    Select Case varBuf(C_Rotate)
        Case C_STAMP_ROTATE_HOLIZONTAL
            optHolizontal.value = True
        Case C_STAMP_ROTATE_VERTICAL
            optVertical.value = True
    End Select
    
    txtText.Text = varBuf(C_Text)
    
    Dim strType As String
    strType = varBuf(C_DateType)
    Select Case strType
        Case C_STAMP_DATE_SYSTEM
            optSystemDate.value = True
        Case C_STAMP_DATE_USER
            optUserDate.value = True
    End Select
    
    Dim strFormat As String
    strFormat = varBuf(C_DateFormat)
    txtFormat.Text = strFormat
    
    Dim strUserDate As String
    strUserDate = varBuf(C_UserDate)
    txtUserDate.Text = strUserDate
    
    Dim lngColor As Long
    lngColor = CLng(varBuf(C_Color))
    lblColor.BackColor = lngColor
    
    txtWidth.Text = varBuf(C_SIZE)
    txtRound.Text = Format$(CDbl(varBuf(C_Round)), "0.00")

    txtLineSize.Text = varBuf(C_LineSize)
    
    Dim strFont As String
    Dim pos As Long
    
    strFont = varBuf(C_Font)

    For i = 0 To cmbFont.ListCount - 1
        If strFont = cmbFont.List(i) Then
            pos = i
        End If
    Next i
    cmbFont.ListIndex = pos

    mblnRefresh = False

    dispPreview
    
End Sub

Private Sub optCircle_Change()
    dispPreview

End Sub

Private Sub optLong_Change()
    dispPreview
End Sub


Private Sub optHolizontal_Change()
    dispPreview
End Sub

Private Sub optSakura_Change()
    dispPreview
End Sub

Private Sub optSquare_Change()
    dispPreview

End Sub

Private Sub optSystemDate_Change()

    dispPreview

End Sub

Private Sub optUserDate_Change()

    dispPreview

End Sub

Private Sub scrBorder_Change()
    dispPreview
End Sub

Private Sub optVertical_Change()
    dispPreview
End Sub

Private Function spinUpSize(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 1
'    If lngValue > 0.5 Then
'        lngValue = 0.5
'    End If
    spinUpSize = Format(lngValue, "0")

End Function

Private Function spinDownSize(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 1
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDownSize = Format(lngValue, "0")

End Function
Private Sub spnLine_Spindown()
    txtLineSize.Text = spinDownSize(txtLineSize.Text)
End Sub
Private Sub spnLine_SpinUp()
    txtLineSize.Text = spinUpSize(txtLineSize.Text)
End Sub

Private Sub spnWidth_SpinDown()
    txtWidth.Text = spinDownWidth(txtWidth.Text)
End Sub

Private Sub spnWidth_SpinUp()
    txtWidth.Text = spinUpWidth(txtWidth.Text)
End Sub

Private Sub spnRound_SpinDown()
    txtRound.Text = spinDownRound(txtRound.Text)
End Sub

Private Sub spnRound_SpinUp()
    txtRound.Text = spinUpRound(txtRound.Text)
End Sub

Private Sub txtFormat_Change()

    dispPreview

End Sub


Private Sub txtLineSize_Change()
    dispPreview
End Sub

Private Sub txtRound_Change()

    dispPreview

End Sub

Private Sub txtText_Change()

    dispPreview
    
End Sub


Private Sub txtUserDate_Change()
    
    dispPreview

End Sub

Private Sub txtWidth_Change()

    dispPreview
    
End Sub

Private Sub UserForm_Initialize()

    Dim datStampBz As StampBzDTO
    Dim col As Collection
    Dim i As Long
    
    Dim strBuf As String
    Dim varBuf() As Variant
    
    ReDim varBuf(C_Text To C_LineSize)

    '設定情報取得
    Set col = getPropertyBz()

    For i = 1 To col.count
        
        Set datStampBz = col(i)
        
        varBuf(C_StampType) = datStampBz.StampType
        varBuf(C_Text) = datStampBz.Text
        varBuf(C_DateType) = datStampBz.DateType
        varBuf(C_DateFormat) = datStampBz.DateFormat
        varBuf(C_UserDate) = datStampBz.UserDate
        varBuf(C_Font) = datStampBz.Font
        varBuf(C_Color) = datStampBz.Color
        varBuf(C_SIZE) = datStampBz.Size
        varBuf(C_Round) = datStampBz.Round
        varBuf(C_Rotate) = datStampBz.Rotate
        varBuf(C_LineSize) = datStampBz.LineSize
        
        strBuf = Join(varBuf, vbTab)
        
        lstStampBz.AddItem ""
        lstStampBz.List(i - 1, C_Text) = datStampBz.Text
        lstStampBz.List(i - 1, C_DATA) = strBuf
    Next

    ActiveCell.Select
    With Application.CommandBars("Formatting").Controls(1)
        For i = 1 To .ListCount
            cmbFont.AddItem .List(i)
        Next i
    End With
    
    If col.count > 0 Then
        lstStampBz.Selected(0) = True
    Else
        mblnRefresh = True
        
        cmbFont.Text = "ＭＳ ゴシック"
        txtUserDate.Text = Format$(Now, "yyyy/m/d")
        optRectangle.value = True
        optSystemDate = True
        txtFormat.Text = "yyyy.m.d"
        txtRound.Text = "0.15"
        txtWidth.Text = "42"
        
        mblnRefresh = False
    End If

End Sub

Private Function getFormatDate(ByVal strFormat As String, _
                        ByVal strType As String, _
                        ByVal strUserDate As String)
    
    On Error Resume Next

    If Len(Trim(strFormat)) = 0 Then
        getFormatDate = ""
        Exit Function
    End If
    
    Select Case strType
        Case C_STAMP_DATE_SYSTEM
            getFormatDate = Format(Now, strFormat)
            
        Case C_STAMP_DATE_USER
            If IsDate(strUserDate) Then
                getFormatDate = Format(CDate(strUserDate), strFormat)
            Else
                getFormatDate = ""
            End If
    End Select

End Function
Private Function spinUpWidth(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 0.5
'    If lngValue > 25 Then
'        lngValue = 25
'    End If
    spinUpWidth = lngValue

End Function

Private Function spinDownWidth(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 0.5
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDownWidth = lngValue

End Function
Private Function spinUpHeight(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 0.5
    If lngValue > 25 Then
        lngValue = 25
    End If
    spinUpHeight = lngValue

End Function

Private Function spinDownHeight(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 0.5
    If lngValue < 15 Then
        lngValue = 15
    End If
    spinDownHeight = lngValue

End Function

Private Function spinUpRound(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 0.01
    If lngValue > 0.5 Then
        lngValue = 0.5
    End If
    spinUpRound = Format(lngValue, "0.00")

End Function

Private Function spinDownRound(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 0.01
    If lngValue < 0 Then
        lngValue = 0
    End If
    spinDownRound = Format(lngValue, "0.00")

End Function

Private Sub UserForm_Terminate()

'    Set mColSet = Nothing

End Sub
