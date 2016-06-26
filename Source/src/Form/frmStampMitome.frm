VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStampMitome 
   Caption         =   "認め印"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11595
   OleObjectBlob   =   "frmStampMitome.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmStampMitome"
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

Private Const C_Text As String = 0
Private Const C_Font As String = 1
Private Const C_Color As String = 2
Private Const C_SIZE As String = 3
Private Const C_Line As String = 4
Private Const C_StampType As String = 5
Private Const C_File As String = 6
Private Const C_LineSize As String = 7
Private Const C_Round As String = 8
Private Const C_Rotate As String = 9
Private Const C_Fill As String = 10
Private Const C_Rect As String = 11

Private Const C_DATA As Long = 1

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2

Private mResult As VbMsgBoxResult
Private mblnRefresh As Boolean

Sub dispPreview()
    
    Dim FS As Object
    Dim strPath As String
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    If mblnRefresh Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtSize.Text) Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtLineSize.Text) Then
        Exit Sub
    End If

    If Not IsNumeric(txtRound.Text) Then
        Exit Sub
    End If

    If Not IsNumeric(txtRect.Text) Then
        Exit Sub
    End If
    
    strPath = rlxGetTempFolder() & C_STAMP_FILE_NAME & ".jpg"
    
    Dim s As StampMitomeDTO
    
    Set s = New StampMitomeDTO
    
    Select Case True
        Case optNormal.value
            s.StampType = C_STAMP_MITOME_NORMAL
            s.Text = txtName.Text
            Dim c As control
            For Each c In Controls
                Select Case c.Tag
                    Case "N"
                        c.enabled = True
                    Case "F"
                        c.enabled = False
                End Select
            Next
        Case Else
             s.StampType = C_STAMP_MITOME_FILE
            s.Text = rlxGetFullpathFromFileName(txtFile.Text)
'            Dim c As control
            For Each c In Controls
                Select Case c.Tag
                    Case "N"
                        c.enabled = False
                    Case "F"
                        c.enabled = True
                End Select
            Next
    End Select
    
    s.Font = cmbFont.Text
    
    Select Case True
        Case optLineSingle.value
'            fraRotate.enabled = True
'            optVertical.enabled = True
'            optHolizontal.enabled = True
            s.Line = C_STAMP_LINE_SINGLE
        Case optLineDouble.value
'            fraRotate.enabled = False
'            optVertical.enabled = False
'            optHolizontal.enabled = False
            s.Line = C_STAMP_LINE_DOUBLE
        Case Else
'            fraRotate.enabled = True
'            optVertical.enabled = True
'            optHolizontal.enabled = True
            s.Line = C_STAMP_LINE_BOLD
    End Select
    
    
    Dim lngSize As Double
    Select Case True
        Case optLineSingle.value, optLineBold.value
            lngSize = ThisWorkbook.Worksheets("stampEx").Shapes("shpMitome").Width

        Case optLineDouble.value
            lngSize = ThisWorkbook.Worksheets("stampEx").Shapes("shpMitome").Width * 0.8

    End Select
    ThisWorkbook.Worksheets("stampEx").Shapes("shpMitome").Height = lngSize
    
    s.Color = getHexColor(lblColor.BackColor)
    s.Size = txtSize.Text
    s.FilePath = txtFile.Text
    s.LineSize = txtLineSize.Text
    s.Round = txtRound.Text
    s.rect = txtRect.Text
    
    Select Case True
        Case optVertical.value
            s.Rotate = C_STAMP_ROTATE_VERTICAL
        Case optHolizontal.value
            s.Rotate = C_STAMP_ROTATE_HOLIZONTAL
    End Select

    If chkFill.value Then
        s.Fill = C_STAMP_FILL_ON
    Else
        s.Fill = C_STAMP_FILL_OFF
    End If

    '編集結果をリストに設定
    Dim i As Long

    i = lstStamp.ListIndex
    If i = -1 Then
        Exit Sub
    End If
    
    mblnRefresh = True
    
    Dim strBuf As String
    Dim varBuf() As Variant
    
    ReDim varBuf(C_Text To C_Rect)

    varBuf(C_StampType) = s.StampType
    varBuf(C_Text) = s.Text
    varBuf(C_File) = s.FilePath
    varBuf(C_Color) = s.Color
    varBuf(C_SIZE) = s.Size
    varBuf(C_Line) = s.Line
    varBuf(C_Font) = s.Font
    varBuf(C_LineSize) = s.LineSize
    varBuf(C_Round) = s.Round
    varBuf(C_Rotate) = s.Rotate
    varBuf(C_Fill) = s.Fill
    varBuf(C_Rect) = s.rect
    
    strBuf = Join(varBuf, vbTab)

    lstStamp.List(i, C_Text) = s.Text
    lstStamp.List(i, C_DATA) = strBuf

    mblnRefresh = False
    
    If s.StampType = C_STAMP_MITOME_NORMAL Then
        
'        Call editStampMitome(s, xlBitmap, lngHeight, lngWidth)
'        Call editStampMitome(s, xlBitmap)
        imgPreview.Picture = editStampMitome(s, xlBitmap)
        
'        lngHeight = ThisWorkbook.Worksheets("stampEx").Shapes("shpMitome").Width
'        lngWidth = ThisWorkbook.Worksheets("stampEx").Shapes("shpMitome").Height
    
    Else
        imgPreview.Picture = LoadPicture("")

        'ファイルが存在しない場合、
        If Not rlxIsFileExists(txtFile.Text) Then
           Exit Sub
        End If
        
        Dim o As Object
        
        Set o = ActiveSheet.Pictures.Insert(s.FilePath)
            
'        lngWidth = .Width
'        lngHeight = .Height
'
'        o.CopyPicture xlScreen, xlBitmap
        
'
'        lngWidth = o.Width
'        lngHeight = o.Height
        
        imgPreview.Picture = CreatePictureFromClipboard(o)
        
        o.Delete
    End If
    
'    With ThisWorkbook.Worksheets("stampEx").ChartObjects.Add(0, 0, lngWidth, lngHeight).Chart
'
'        .Paste
'        .ChartArea.Border.LineStyle = 0
'        .Export strPath, "JPG"
'
'        .Parent.Delete
'
'    End With
        
'    imgPreview.Picture = CreatePictureFromClipboard()
    
End Sub

Private Sub chkFill_Change()
    dispPreview
End Sub

Private Sub cmbFont_Click()
    If cmbFont.ListIndex = -1 Then
    Else
'        getTextSize lblName.Caption, cmbFont.List(cmbFont.ListIndex)
        dispPreview
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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
    If lstStamp.ListCount <= 1 Then
        Exit Sub
    End If

    Select Case lngMode
        Case C_UP
            lngStart = 0
            lngEnd = lstStamp.ListCount - 1
            lngInc = 1
        Case C_DOWN
            lngStart = lstStamp.ListCount - 1
            lngEnd = 0
            lngInc = -1
    End Select

    For lngCnt = lngStart To lngEnd Step lngInc
    
        If lstStamp.Selected(lngCnt) Then
            '選択された行がすでに開始行の場合移動不可
            If lngCnt = lngStart Then
                Exit For
            End If
            
            lngCmp = lngCnt + lngInc * -1
            
            Dim i As Long
            For i = C_Text To C_DATA
                varTmp = lstStamp.List(lngCnt, i)
                lstStamp.List(lngCnt, i) = lstStamp.List(lngCmp, i)
                lstStamp.List(lngCmp, i) = varTmp
            Next
            
            lstStamp.Selected(lngCnt) = False
            lstStamp.Selected(lngCnt + lngInc * -1) = True
        End If
    
    Next

End Sub


Private Sub cmdFile_Click()
   Dim strFile As String


    strFile = Application.GetOpenFilename("ファイル(*.*),(*.*)", , "画像ファイル", , False)
    If strFile = "False" Then
        'ファイル名が指定されなかった場合
        Exit Sub
    End If
    
    txtFile.Text = strFile
'    Call dispPicture
End Sub

Private Sub cmdOK_Click()
    
'    Dim strSize As String
'    Dim strColor As String
'    Dim strLine As String
'    Dim dblSize As Double
'
'    dblSize = Val(txtSize.Text)
'    Select Case dblSize
'        Case 10.5 To 24
'        Case Else
'            MsgBox "サイズは１０．５ｍｍ～２４ｍｍで入力してください。", vbExclamation + vbOKOnly, C_TITLE
'            Exit Sub
'    End Select
'
'    mResult = vbOK
'    Unload Me
    
'    SaveSetting C_TITLE, "Mitome", "Name", txtName.Text
'    SaveSetting C_TITLE, "Mitome", "Font", cmbFont.Text
'
''    Select Case True
''        Case optColorBlack.Value
''            strColor = C_STAMP_COLOR_BLACK
''        Case optColorRed.Value
''            strColor = C_STAMP_COLOR_RED
''    End Select
'    strColor = "&H" & Right$("00000000" & Hex(lblColor.BackColor), 8)
'    SaveSetting C_TITLE, "Mitome", "Color", strColor
'
'    Select Case True
'        Case optLineSingle.Value
'            strLine = C_STAMP_LINE_SINGLE
'        Case optLineDouble.Value
'            strLine = C_STAMP_LINE_DOUBLE
'    End Select
'    SaveSetting C_TITLE, "Mitome", "Line", strLine
'
''    Select Case True
''        Case optSize1.Value
''            strSize = C_STAMP_SIZE_105
''        Case optSize2.Value
''            strSize = C_STAMP_SIZE_120
''        Case optSize3.Value
''            strSize = C_STAMP_SIZE_135
''        Case optSize4.Value
''            strSize = C_STAMP_SIZE_150
''        Case optSize5.Value
''            strSize = C_STAMP_SIZE_165
''        Case optSize6.Value
''            strSize = C_STAMP_SIZE_180
''        Case optSize7.Value
''            strSize = C_STAMP_SIZE_210
''        Case optSize8.Value
''            strSize = C_STAMP_SIZE_240
''    End Select
'    SaveSetting C_TITLE, "Mitome", "Size", dblSize
    
    Dim s As StampMitomeDTO
    Dim col As Collection
    Dim i As Long

    Set col = New Collection
    '設定情報取得

    For i = 0 To lstStamp.ListCount - 1
        
        Set s = New StampMitomeDTO
        
        Dim varBuf As Variant
        varBuf = Split(lstStamp.List(i, C_DATA), vbTab)
        
        s.StampType = varBuf(C_StampType)
        s.Text = varBuf(C_Text)
        s.Font = varBuf(C_Font)
        s.Color = varBuf(C_Color)
        s.Size = varBuf(C_SIZE)
        s.Line = varBuf(C_Line)
        s.FilePath = varBuf(C_File)
        s.LineSize = varBuf(C_LineSize)
        s.Round = varBuf(C_Round)
        s.Rotate = varBuf(C_Rotate)
        s.Fill = varBuf(C_Fill)
        s.rect = varBuf(C_Rect)
        
        If IsNumeric(s.rect) Then
        Else
            MsgBox "社畜度には数値をで入力してください。（-100%～100%）", vbExclamation + vbOKOnly, C_TITLE
            lstStamp.Selected(i) = True
            txtSize.SetFocus
            Exit Sub
        End If
        
        If IsNumeric(s.Size) Then
        Else
            MsgBox "幅には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStamp.Selected(i) = True
            txtSize.SetFocus
            Exit Sub
        End If
        
        If CDbl(s.Size) < 0 Then
            MsgBox "幅は０以上を入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStamp.Selected(i) = True
            txtSize.SetFocus
            Exit Sub
        End If
        
        If IsNumeric(s.LineSize) Then
        Else
            MsgBox "外枠には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStamp.Selected(i) = True
            txtLineSize.SetFocus
            Exit Sub
        End If
        
        If CDbl(s.LineSize) < 0 Then
            MsgBox "外枠は０以上を入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStamp.Selected(i) = True
            txtLineSize.SetFocus
            Exit Sub
        End If
        
        If IsNumeric(s.Round) Then
        Else
            MsgBox "角丸には数値をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStamp.Selected(i) = True
            txtRound.SetFocus
            Exit Sub
        End If
        
        If CDbl(s.Round) < 0 Then
            MsgBox "角丸は0.00～0.50を入力してください。", vbExclamation + vbOKOnly, C_TITLE
            lstStamp.Selected(i) = True
            txtRound.SetFocus
            Exit Sub
        End If
        
        'ファイルの存在チェック
        If s.StampType = C_STAMP_MITOME_FILE Then
            If Not rlxIsFileExists(s.FilePath) Then
                MsgBox "画像ファイルが存在しません。", vbExclamation + vbOKOnly, C_TITLE
                lstStamp.Selected(i) = True
                txtFile.SetFocus
                Exit Sub
            End If
        End If
        
        col.Add s
        
        Set s = Nothing
        
    Next

    'プロパティ保存
    setPropertyMitome col
    
    Call SaveSetting(C_TITLE, "StampMitome", "Confirm", chkConfirm.value)
    
    Set col = Nothing
    
    'リボンのリフレッシュ
    Call RefreshRibbon
    
    On Error GoTo 0
    
    mResult = vbOK
    Unload Me
    
End Sub




Private Sub optColorBlack_Click()
'    Call changeBackGround
    dispPreview
End Sub

Private Sub optColorRed_Click()
'    Call changeBackGround
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

Private Sub lstStamp_Change()

    Dim i As Long
    
    If mblnRefresh Then
        Exit Sub
    End If

    mblnRefresh = True

    i = lstStamp.ListIndex
    If i = -1 Then
        Exit Sub
    End If
    
    Dim varBuf As Variant
    
    varBuf = Split(lstStamp.List(i, C_DATA), vbTab)
    
    Select Case varBuf(C_StampType)
        Case C_STAMP_MITOME_NORMAL
            optNormal.value = True
            txtName.Text = varBuf(C_Text)
        Case Else
            optFile.value = True
            txtName.Text = ""
    End Select
    
    Select Case varBuf(C_Line)
        Case C_STAMP_LINE_SINGLE
            optLineSingle.value = True
        Case C_STAMP_LINE_DOUBLE
            optLineDouble.value = True
        Case Else
            optLineBold.value = True
    End Select
    
    
    Dim lngColor As Long
    lngColor = getLongColor(varBuf(C_Color))
    lblColor.BackColor = lngColor
    
    txtSize.Text = varBuf(C_SIZE)
    txtFile.Text = varBuf(C_File)
    txtLineSize.Text = varBuf(C_LineSize)
    txtRound.Text = varBuf(C_Round)
    txtRect.Text = varBuf(C_Rect)

    Select Case varBuf(C_Rotate)
        Case C_STAMP_ROTATE_HOLIZONTAL
            optHolizontal.value = True
        Case C_STAMP_ROTATE_VERTICAL
            optVertical.value = True
    End Select
    
    If varBuf(C_Fill) = C_STAMP_FILL_ON Then
        chkFill.value = True
    Else
        chkFill.value = False
    End If
    
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

Private Sub optFile_Change()
    dispPreview
End Sub

Private Sub optHolizontal_Click()
    dispPreview
End Sub

Private Sub optLineBold_Change()
    dispPreview
End Sub


Private Sub optLineDouble_Click()
    dispPreview
End Sub

Private Sub optLineSingle_Click()
    dispPreview
End Sub
Private Sub optNormal_Change()
    dispPreview
End Sub

Private Sub optVertical_Click()
    dispPreview
End Sub

Private Sub spnLine_SpinUp()
    txtLineSize.Text = spinUpSize(txtLineSize.Text)
End Sub

Private Sub spnLine_Spindown()
    txtLineSize.Text = spinDownSize(txtLineSize.Text)
End Sub

Private Sub spnRect_SpinDown()
    txtRect.Text = spinDownRect(txtRect.Text)
End Sub

Private Sub spnRect_SpinUp()
    txtRect.Text = spinUpRect(txtRect.Text)
End Sub

Private Sub spnRound_SpinDown()
    txtRound.Text = spinDownRound(txtRound.Text)
End Sub

Private Sub spnRound_SpinUp()
    txtRound.Text = spinUpRound(txtRound.Text)
End Sub

Private Sub spnSize_SpinDown()
    txtSize.Text = spinDown(txtSize.Text)
End Sub

Private Sub spnSize_SpinUp()
    txtSize.Text = spinUp(txtSize.Text)
End Sub

Private Sub txtFile_Change()
    dispPreview
End Sub

Private Sub txtLineSize_Change()
    dispPreview
End Sub

Private Sub txtName_Change()
    
'    getTextSize txtName.Text, lblName.Font.Name
    dispPreview
    
End Sub



Private Sub txtName_Enter()

'    getTextSize txtName.Text, lblName.Font.Name
    dispPreview

End Sub

Private Sub txtRect_Change()
    dispPreview
End Sub

Private Sub txtRound_Change()
    dispPreview
End Sub

Private Sub txtSize_Change()
    dispPreview
End Sub

Private Sub UserForm_Initialize()

'    Dim strFont As String
'    Dim strName As String
'    Dim lngColor As Long
'    Dim strLine As String
'    Dim strSize As String
'
'    mblnRefresh = True
'
'    getPropertyMitome strName, strFont, lngColor, strLine, strSize
'
'
''    lblName.Font.Name = strFont
'    txtName.Text = strName
'
''    Select Case strColor
''        Case C_STAMP_COLOR_BLACK
''            optColorBlack.Value = True
''        Case C_STAMP_COLOR_RED
''            optColorRed.Value = True
''    End Select
'
'    lblColor.BackColor = lngColor
'
'    Select Case strLine
'        Case C_STAMP_LINE_SINGLE
'            optLineSingle.Value = True
'        Case C_STAMP_LINE_DOUBLE
'            optLineDouble.Value = True
'    End Select
'
''    Select Case strSize
''        Case C_STAMP_SIZE_105
''            optSize1.Value = True
''        Case C_STAMP_SIZE_120
''            optSize2.Value = True
''        Case C_STAMP_SIZE_135
''            optSize3.Value = True
''        Case C_STAMP_SIZE_150
''            optSize4.Value = True
''        Case C_STAMP_SIZE_165
''            optSize5.Value = True
''        Case C_STAMP_SIZE_180
''            optSize6.Value = True
''        Case C_STAMP_SIZE_210
''            optSize7.Value = True
''        Case C_STAMP_SIZE_240
''            optSize8.Value = True
''    End Select
'    txtSize.Text = strSize
'
'
'    Dim i As Long
'    Dim pos As Long
'
'    pos = -1
'    ActiveCell.Select
'    With Application.CommandBars("Formatting").Controls(1)
'        For i = 1 To .ListCount
'            cmbFont.AddItem .List(i)
'            If strFont = .List(i) Then
'                pos = i - 1
'            End If
'        Next i
'    End With
'
'    cmbFont.ListIndex = pos
'
'    mblnRefresh = False
'
'    dispPreview

    Dim s As StampMitomeDTO
    Dim col As Collection
    Dim i As Long
    
    Dim strBuf As String
    Dim varBuf() As Variant
    
    ReDim varBuf(C_Text To C_Rect)
    
    '設定情報取得
    Set col = getPropertyMitome()

    For i = 1 To col.count
        
        Set s = col(i)
        
        varBuf(C_StampType) = s.StampType
        varBuf(C_Text) = s.Text
        varBuf(C_Font) = s.Font
        varBuf(C_Color) = s.Color
        varBuf(C_SIZE) = s.Size
        varBuf(C_Line) = s.Line
        varBuf(C_File) = s.FilePath
        varBuf(C_LineSize) = s.LineSize
        varBuf(C_Round) = s.Round
        varBuf(C_Rotate) = s.Rotate
        varBuf(C_Fill) = s.Fill
        varBuf(C_Rect) = s.rect
        
        strBuf = Join(varBuf, vbTab)
        
        lstStamp.AddItem ""
        lstStamp.List(i - 1, C_Text) = s.Text
        lstStamp.List(i - 1, C_DATA) = strBuf
        
    Next

    ActiveCell.Select
    With Application.CommandBars("Formatting").Controls(1)
        For i = 1 To .ListCount
            cmbFont.AddItem .List(i)
        Next i
    End With
        

    If col.count > 0 Then
        lstStamp.Selected(0) = True
    Else
        
        mblnRefresh = True
        
        txtName.Text = ""
        txtFile.Text = ""
        cmbFont.Text = "ＭＳ ゴシック"
        optLineSingle.value = True
        optNormal.value = True
        txtSize.Text = "10.5"
        lblColor.BackColor = vbRed
        txtLineSize.Text = "10"
        txtRound.Text = "0.50"
        optVertical.value = True
    
        mblnRefresh = False
'        dispPreview
        
    End If

    chkConfirm.value = GetSetting(C_TITLE, "StampMitome", "Confirm", False)
    
End Sub
Private Function spinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 0.5
'    If lngValue > 24 Then
'        lngValue = 24
'    End If
    spinUp = lngValue

End Function

Private Function spinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 0.5
    If lngValue < 10.5 Then
        lngValue = 10.5
    End If
    spinDown = lngValue

End Function
'Sub changeBackGround()
'
'    Select Case True
'        Case optLineSingle.Value
'            Select Case True
'                Case optColorBlack.Value
'                    Image1.Picture = CBlack.Picture
'                    lblName.ForeColor = vbBlack
'                Case optColorRed.Value
'                    Image1.Picture = CRed.Picture
'                    lblName.ForeColor = vbRed
'            End Select
'        Case optLineDouble.Value
'            Select Case True
'                Case optColorBlack.Value
'                    Image1.Picture = DBlack.Picture
'                    lblName.ForeColor = vbBlack
'                Case optColorRed.Value
'                    Image1.Picture = DRed.Picture
'                    lblName.ForeColor = vbRed
'            End Select
'    End Select
'
'End Sub
'Private Sub getTextSize(ByVal strChar As String, ByVal strFont As String)
'
'    Dim sngSize As Single
'
'    sngSize = 128
'
''    lblSize.Caption = strChar
''    lblSize.Font.Name = strFont
''
''    Do
''        sngSize = sngSize - 1
''        With lblSize
''            .Font.Size = sngSize
''            .AutoSize = False
'''            .Width = 40
''            .AutoSize = True
''        End With
''    Loop Until lblSize.Height < 100
'
'    lblName.Caption = strChar
'    lblName.Font.Name = strFont
'
'    Do
'        sngSize = sngSize - 1
'        With lblName
'            .Font.Size = sngSize
'            .AutoSize = False
''            .Width = 40
'            .AutoSize = True
'        End With
'    Loop Until lblName.Height < 100
'
'
''    lblName.Width = lblSize.Width
''    lblName.Caption = strChar
''    lblName.Font.Name = strFont
''    lblName.Font.Size = sngSize
'
'    Dim lngLeft As Long
'    lngLeft = (Image1.Left + (Image1.Width / 2)) - (lblName.Left + (lblName.Width / 2))
'    lblName.Left = lblName.Left + lngLeft
'
'    Dim lngTop As Long
'    lngTop = (Image1.Top + (Image1.Height / 2)) - (lblName.Top + (lblName.Height / 2))
'    lblName.Top = lblName.Top + lngTop
'
'End Sub
'Private Function getFormatDate(ByVal strFormat As String, _
'                        ByVal strType As String, _
'                        ByVal strUserDate As String)
'
'    On Error Resume Next
'
'    If Len(Trim(strFormat)) = 0 Then
'        getFormatDate = ""
'        Exit Function
'    End If
'
'    Select Case strType
'        Case C_STAMP_DATE_SYSTEM
'            getFormatDate = Format(Now, strFormat)
'
'        Case C_STAMP_DATE_USER
'            If IsDate(strUserDate) Then
'                getFormatDate = Format(CDate(strUserDate), strFormat)
'            Else
'                getFormatDate = ""
'            End If
'    End Select
'
'End Function
Private Sub cmdAdd_Click()

    Dim i As Long
    Dim strBuf As String
    Dim varBuf() As Variant
    
    ReDim varBuf(C_Text To C_Rect)
    
    i = lstStamp.ListCount

    
    Select Case True
        Case optNormal.value
           varBuf(C_StampType) = C_STAMP_MITOME_NORMAL
        Case Else
            varBuf(C_StampType) = C_STAMP_MITOME_FILE
    End Select
    
    varBuf(C_Text) = txtName.Text
    
    varBuf(C_Color) = getHexColor(lblColor.BackColor)
    
    varBuf(C_SIZE) = txtSize.Text
    
    Select Case True
        Case optLineSingle.value
            varBuf(C_Line) = C_STAMP_LINE_SINGLE
        Case optLineDouble.value
            varBuf(C_Line) = C_STAMP_LINE_DOUBLE
        Case Else
            varBuf(C_Line) = C_STAMP_LINE_BOLD
    End Select
    
    varBuf(C_Font) = cmbFont.Text
    varBuf(C_File) = txtFile.Text
    
    varBuf(C_LineSize) = txtLineSize.Text
    varBuf(C_Round) = txtRound.Text
    varBuf(C_Rect) = txtRect.Text
    
    If chkFill.value Then
        varBuf(C_Fill) = C_STAMP_FILL_ON
    Else
        varBuf(C_Fill) = C_STAMP_FILL_OFF
    End If
    
    strBuf = Join(varBuf, vbTab)
    
    lstStamp.AddItem ""
    lstStamp.List(i, C_Text) = txtName.Text
    lstStamp.List(i, C_DATA) = strBuf
    
    lstStamp.Selected(i) = True
    
End Sub


Private Sub cmdDel_Click()
    Dim i As Long
    
    For i = 0 To lstStamp.ListCount
    
        If lstStamp.Selected(i) Then
            lstStamp.RemoveItem i
            Exit Sub
        End If
    
    Next
End Sub

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
Private Function spinUpRect(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 5
    If lngValue > 100 Then
        lngValue = 100
    End If
    spinUpRect = Format(lngValue, "0")

End Function

Private Function spinDownRect(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 5
    If lngValue < -100 Then
        lngValue = -100
    End If
    spinDownRect = Format(lngValue, "0")

End Function
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
