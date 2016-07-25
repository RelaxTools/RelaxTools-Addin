VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStamp 
   Caption         =   "データ印"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   OleObjectBlob   =   "frmStamp.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmStamp"
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

Private Const C_TEXT As Long = 0
Private Const C_Upper As Long = 1
Private Const C_DateType As Long = 2
Private Const C_DateFormat As Long = 3
Private Const C_UserDate As Long = 4
Private Const C_Lower As Long = 5
Private Const C_Font As Long = 6
Private Const C_Color As Long = 7
Private Const C_SIZE As Long = 8
Private Const C_Line As Long = 9
Private Const c_WordArt As Long = 10
Private Const C_Fill As Long = 11

Private Const C_DATA As Long = 1

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2

Private mResult As VbMsgBoxResult

Private mblnRefresh As Boolean

Private Sub chkFill_Change()
    dispPreview
End Sub


Private Sub chkWordArt_Change()
    dispPreview
End Sub

Private Sub cmbFont_Click()
    If cmbFont.ListIndex = -1 Then
    Else
        dispPreview
    End If
End Sub

Private Sub cmdAdd_Click()

    Dim i As Long
    
    i = lstStamp.ListCount

    Dim strBuf As String
    Dim varBuf() As Variant
    
    ReDim varBuf(C_TEXT To C_Fill)

    
    
    varBuf(C_Upper) = txtUpper.Text
    varBuf(C_Lower) = txtLower.Text
    
    Select Case True
        Case optSystemDate.value
            varBuf(C_DateType) = C_STAMP_DATE_SYSTEM
        Case optUserDate.value
            varBuf(C_DateType) = C_STAMP_DATE_USER
    End Select
    
    varBuf(C_DateFormat) = txtFormat.Text
    
    varBuf(C_UserDate) = txtUserDate.Text
    
    varBuf(C_Color) = getHexColor(lblColor.BackColor)
    
    varBuf(C_SIZE) = txtSize.Text
    
    Select Case True
        Case optLineSingle.value
            varBuf(C_Line) = C_STAMP_LINE_SINGLE
        Case optLineDouble.value
            varBuf(C_Line) = C_STAMP_LINE_DOUBLE
        Case optLineBold.value
            varBuf(C_Line) = C_STAMP_LINE_BOLD
    End Select
    
    varBuf(C_Font) = cmbFont.Text
    
    If chkWordArt.value = True Then
        varBuf(c_WordArt) = C_STAMP_WORDART_ON
    Else
        varBuf(c_WordArt) = C_STAMP_WORDART_OFF
    End If
    
    If chkFill.value = True Then
        varBuf(C_Fill) = C_STAMP_FILL_ON
    Else
        varBuf(C_Fill) = C_STAMP_FILL_OFF
    End If
    
    strBuf = Join(varBuf, vbTab)
    
    lstStamp.AddItem ""
    
    lstStamp.List(i, C_TEXT) = txtUpper.Text & " + " & txtLower.Text
    lstStamp.List(i, C_DATA) = strBuf
    
    lstStamp.Selected(i) = True
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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
            For i = C_TEXT To C_DATA
                varTmp = lstStamp.List(lngCnt, i)
                lstStamp.List(lngCnt, i) = lstStamp.List(lngCmp, i)
                lstStamp.List(lngCmp, i) = varTmp
            Next
            
            lstStamp.Selected(lngCnt) = False
            lstStamp.Selected(lngCnt + lngInc * -1) = True
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

    Dim FS As Object
'    Dim strPath As String
    
    If mblnRefresh Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtSize.Text) Then
        Exit Sub
    End If

'    strPath = rlxGetTempFolder() & C_STAMP_FILE_NAME & ".jpg"
    
    Dim s As StampDatDTO
    
    Set s = New StampDatDTO
    
    s.Upper = txtUpper.Text
    
    Select Case True
        Case optSystemDate.value
            s.DateType = C_STAMP_DATE_SYSTEM
        Case optUserDate.value
            s.DateType = C_STAMP_DATE_USER
    End Select
    
    s.Lower = txtLower.Text
    s.DateFormat = txtFormat.Text
    s.Font = cmbFont.Text
    
    Select Case True
        Case optLineSingle.value
             s.Line = C_STAMP_LINE_SINGLE

        Case optLineDouble.value
             s.Line = C_STAMP_LINE_DOUBLE

        Case optLineBold.value
             s.Line = C_STAMP_LINE_BOLD

    End Select
    s.Color = getHexColor(lblColor.BackColor)
    s.Size = txtSize.Text
    s.UserDate = txtUserDate.Text
    
    If chkWordArt.value = True Then
        s.WordArt = C_STAMP_WORDART_ON
    Else
        s.WordArt = C_STAMP_WORDART_OFF
    End If
    
    If chkFill.value = True Then
        s.Fill = C_STAMP_FILL_ON
    Else
        s.Fill = C_STAMP_FILL_OFF
    End If

'    Call editStamp(s, xlBitmap)
    imgPreview.Picture = editStamp(s, xlBitmap)
    
'    imgPreview.Picture = CreatePictureFromClipboard()

    '編集結果をリストに設定
    Dim i As Long

    i = lstStamp.ListIndex
    If i = -1 Then
        Exit Sub
    End If
    
    Dim strBuf As String
    Dim varBuf() As Variant
    
    ReDim varBuf(C_TEXT To C_Fill)
    
    mblnRefresh = True
            
    varBuf(C_Upper) = s.Upper
    varBuf(C_Lower) = s.Lower
    varBuf(C_DateType) = s.DateType
    varBuf(C_DateFormat) = s.DateFormat
    varBuf(C_UserDate) = s.UserDate
    varBuf(C_Color) = s.Color
    varBuf(C_SIZE) = s.Size
    varBuf(C_Line) = s.Line
    varBuf(C_Font) = s.Font
    varBuf(c_WordArt) = s.WordArt
    varBuf(C_Fill) = s.Fill
    strBuf = Join(varBuf, vbTab)

    lstStamp.List(i, C_TEXT) = s.Upper & " + " & s.Lower
    lstStamp.List(i, C_DATA) = strBuf
    
    mblnRefresh = False
    
End Sub
Private Sub cmdOK_Click()
        
    Dim s As StampDatDTO
    Dim col As Collection
    Dim i As Long

    Set col = New Collection
    '設定情報取得

    For i = 0 To lstStamp.ListCount - 1
        
        Set s = New StampDatDTO
        
        Dim varBuf As Variant
        
        varBuf = Split(lstStamp.List(i, C_DATA), vbTab)
        
        s.Upper = varBuf(C_Upper)
        s.DateType = varBuf(C_DateType)
        s.DateFormat = varBuf(C_DateFormat)
        s.UserDate = varBuf(C_UserDate)
        s.Font = varBuf(C_Font)
        s.Color = varBuf(C_Color)
        s.Size = varBuf(C_SIZE)
        s.Line = varBuf(C_Line)
        s.Lower = varBuf(C_Lower)
        s.WordArt = varBuf(c_WordArt)
        s.Fill = varBuf(C_Fill)
        
        If s.DateType = C_STAMP_DATE_USER Then
            If IsDate(s.UserDate) Then
            Else
                MsgBox "指定日付には有効な日付をで入力してください。", vbExclamation + vbOKOnly, C_TITLE
                lstStamp.Selected(i) = True
                txtUserDate.SetFocus
                Exit Sub
            End If
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
        
        col.Add s
        
        Set s = Nothing
        
    Next

    'プロパティ保存
    setProperty col
    Call SaveSetting(C_TITLE, "Stamp", "Confirm", chkConfirm.value)

    
    Set col = Nothing
    
    'リボンのリフレッシュ
    Call RefreshRibbon
    
    On Error GoTo 0
    
    mResult = vbOK
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

Private Sub lstStamp_Click()

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
    
    
    
    Select Case varBuf(C_Line)
        Case C_STAMP_LINE_SINGLE
            optLineSingle.value = True
        Case C_STAMP_LINE_DOUBLE
            optLineDouble.value = True
        Case C_STAMP_LINE_BOLD
            optLineBold.value = True
    End Select
    
    txtUpper.Text = varBuf(C_Upper)
    txtLower.Text = varBuf(C_Lower)
    
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
    lngColor = getLongColor(varBuf(C_Color))
    lblColor.BackColor = lngColor
    
    txtSize.Text = varBuf(C_SIZE)

    If varBuf(c_WordArt) = C_STAMP_WORDART_ON Then
        chkWordArt.value = True
    Else
        chkWordArt.value = False
    End If

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
Private Sub optLineBold_Click()
    dispPreview
End Sub
Private Sub optLineDouble_Click()
    dispPreview
End Sub
Private Sub optLineSingle_Click()
    dispPreview
End Sub
Private Sub optSystemDate_Change()
    dispPreview
End Sub
Private Sub optUserDate_Change()
    dispPreview
End Sub
Private Sub spnSize_SpinDown()
    txtSize.Text = spinDown(txtSize.Text)
End Sub
Private Sub spnSize_SpinUp()
    txtSize.Text = spinUp(txtSize.Text)
End Sub
Private Sub txtFormat_Change()
    dispPreview
End Sub
Private Sub txtLower_Change()
    dispPreview
End Sub
Private Sub txtSize_Change()
    dispPreview
End Sub
Private Sub txtUpper_Change()
    dispPreview
End Sub
Private Sub txtUserDate_Change()
    dispPreview
End Sub
Private Sub UserForm_Initialize()

    Dim s As StampDatDTO
    Dim col As Collection
    Dim i As Long
    
    Dim strBuf As String
    Dim varBuf() As Variant
    
    ReDim varBuf(C_TEXT To C_Fill)
    
    '設定情報取得
    Set col = getProperty()

    For i = 1 To col.count
        
        Set s = col(i)
        
        varBuf(C_Upper) = s.Upper
        varBuf(C_DateType) = s.DateType
        varBuf(C_DateFormat) = s.DateFormat
        varBuf(C_UserDate) = s.UserDate
        varBuf(C_Lower) = s.Lower
        varBuf(C_Font) = s.Font
        varBuf(C_Color) = s.Color
        varBuf(C_SIZE) = s.Size
        varBuf(C_Line) = s.Line
        varBuf(c_WordArt) = s.WordArt
        varBuf(C_Fill) = s.Fill
        
        lstStamp.AddItem ""
        
        strBuf = Join(varBuf, vbTab)
        
        lstStamp.List(i - 1, C_TEXT) = s.Upper & " + " & s.Lower
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
        
        txtUpper.Text = ""
        txtLower.Text = ""
        cmbFont.Text = "ＭＳ ゴシック"
        txtUserDate.Text = Format$(Now, "yyyy/m/d")
        optLineSingle.value = True
        optSystemDate = True
        txtFormat.Text = "yyyy.m.d"
        txtSize.Text = "15"
        lblColor.BackColor = vbRed
    
        mblnRefresh = False
    End If
    
    chkConfirm.value = GetSetting(C_TITLE, "Stamp", "Confirm", False)

    mblnRefresh = False
    
    txtFormat.AddItem "yyyy/mm/dd"
    txtFormat.AddItem "yyyy.mm.dd"
    txtFormat.AddItem "'yy.mm.dd"
    txtFormat.AddItem "ge.m.d"
    txtFormat.AddItem "gge.m.d"
    txtFormat.AddItem "ggge年m月d日"

    dispPreview

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
Private Function spinUp(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue + 0.5
    If lngValue > 25 Then
        lngValue = 25
    End If
    spinUp = lngValue

End Function

Private Function spinDown(ByVal vntValue As Variant) As Variant

    Dim lngValue As Variant
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 0.5
    If lngValue < 15 Then
        lngValue = 15
    End If
    spinDown = lngValue

End Function

