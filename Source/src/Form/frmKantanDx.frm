VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKantanDx 
   Caption         =   "かんたん表Dx"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   OleObjectBlob   =   "frmKantanDx.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmKantanDx"
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
' この画面の動き
'
'　テキストボックスの入力をリストに設定→各ChangeイベントからdispPreviewでセット
'　リストを各テキストボックスに設定→リストのクリックイベントで各テキストボックスへ設定後、DispPrewview
'
'-----------------------------------------------------------------------------------------------------
Option Explicit
Private colBorder As New Collection
Private Const C_TEXT As Long = 0
Private Const C_DATA As Long = 1

Private Const C_UP As Long = 1
Private Const C_DOWN As Long = 2
Private mblnRefresh As Boolean

Private Sub txtCol_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtHead_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtHoganJudgeLine_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()

    Dim lngCnt As Long
    Dim i As Long
    Dim strBuf As String

    'スタイル設定
    colBorder.Add lbl01
    colBorder.Add lbl02
    colBorder.Add lbl03
    colBorder.Add lbl04
    colBorder.Add lbl05
    colBorder.Add lbl06
    colBorder.Add lbl07
    colBorder.Add lbl08
    colBorder.Add lbl09
    colBorder.Add lbl10
    colBorder.Add lbl11
    colBorder.Add lbl12
    colBorder.Add lbl13
    colBorder.Add lbl14
    
    Dim lngPos As Long
    
    '実線デフォルト
    lngPos = 7
    setBorder lngPos


    Dim col As Collection
    
    Set col = basKantan.getPropertyKantan()

    Dim k As KantanLine

    For i = 0 To col.count - 1
    
        lstStamp.AddItem ""
        strBuf = GetSetting(C_TITLE, "KantanDx", Format(i + 1, "000"), "")
        
        Set k = col(i + 1)
        lstStamp.List(i, C_TEXT) = k.Text
        
        'デシリアライズ
        strBuf = basKantan.serialize(k)
        lstStamp.List(i, C_DATA) = strBuf
    
    Next

    If col.count > 0 Then
        lstStamp.Selected(0) = True
    End If

End Sub

Private Sub chkAutoHogan_Change()
    dispPreview
End Sub
Private Sub chkHoganMode_Change()
    dispPreview
End Sub
Private Sub chkHRepeat_Change()
    dispPreview
End Sub
Private Sub chkVRepeat_Change()
    dispPreview
End Sub
Private Sub cmdAdd_Click()

    Dim i As Long
    Dim k As KantanLine
    Dim strBuf As String

    Set k = New KantanLine
    
    k.Text = txtName.Text
    k.HHeadLineCount = txtHead.Text
    k.VHeadLineCount = txtCol.Text
    
    k.OutSideLine = Val(imgOutSide.tag)
    k.VHeadBorderLine = Val(imgVHead.tag)
    k.HHeadBorderLine = Val(imgHHead.tag)
    k.HInsideLine = Val(imgHLine.tag)
    k.VInsideLine = Val(imgVLine.tag)
    
    k.HeadColor = lblHead.BackColor
    k.EvenColor = lblEven.BackColor
    
    k.EnableHogan = chkHoganMode.value
    k.EnableEvenColor = chkLineColor.value
    
    k.AuthoHogan = chkAutoHogan.value
    k.HoganJudgeLineCount = Val(txtHoganJudgeLine.Text)
    
    k.EnableHRepeat = chkHRepeat.value
    k.EnableVRepeat = chkVRepeat.value

    i = lstStamp.ListCount
    
    lstStamp.AddItem ""

    lstStamp.List(i, C_TEXT) = k.Text
    strBuf = serialize(k)
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
'
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


Private Sub cmdOK_Click()

    Dim s As KantanLine
    Dim col As Collection
    Dim i As Long
    Dim varBuf As Variant

    Set col = New Collection
    '設定情報取得

    For i = 0 To lstStamp.ListCount - 1

        varBuf = lstStamp.List(i, C_DATA)

        Set s = deserialize(varBuf)

        col.Add s

        Set s = Nothing

    Next

    'プロパティ保存
    setPropertyKantan col

    Set col = Nothing

    'リボンのリフレッシュ
    Call RefreshRibbon

    On Error GoTo 0

    Unload Me

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

    Dim strBuf As String

    strBuf = lstStamp.List(i, C_DATA)

    Dim k As KantanLine

    Set k = basKantan.deserialize(strBuf)

    txtName.Text = k.Text
    txtHead.Text = k.HHeadLineCount
    txtCol.Text = k.VHeadLineCount

    Call setPicture(imgOutSide, k.OutSideLine)
    Call setPicture(imgVHead, k.VHeadBorderLine)
    Call setPicture(imgHHead, k.HHeadBorderLine)
    Call setPicture(imgHLine, k.HInsideLine)
    Call setPicture(imgVLine, k.VInsideLine)

    lblHead.BackColor = k.HeadColor
    lblEven.BackColor = k.EvenColor
    chkLineColor.value = k.EnableEvenColor
    chkHRepeat.value = k.EnableHRepeat
    chkVRepeat.value = k.EnableVRepeat
    txtHoganJudgeLine.Text = k.HoganJudgeLineCount
    chkAutoHogan.value = k.AuthoHogan
    chkHoganMode.value = k.EnableHogan

    mblnRefresh = False

    dispPreview
End Sub
Private Sub chkLineColor_Change()

    Call dispPreview
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub imgHHead_Click()
    Call setPicture(imgHHead, getBorder)
End Sub
Private Sub imgHLine_Click()
    Call setPicture(imgHLine, getBorder)
End Sub
Private Sub imgOutSide_Click()
    Call setPicture(imgOutSide, getBorder)
End Sub
Private Sub imgVHead_Click()
    Call setPicture(imgVHead, getBorder)
End Sub
Private Sub imgVLine_Click()
    Call setPicture(imgVLine, getBorder)
End Sub
Private Sub spnHoganJudgeLine_SpinDown()
    txtHoganJudgeLine.Text = spinDown(txtHoganJudgeLine.Text, 1)
End Sub
Private Sub spnHoganJudgeLine_SpinUp()
    txtHoganJudgeLine.Text = spinUp(txtHoganJudgeLine.Text)
End Sub
Private Sub txtCol_Change()
    dispPreview
End Sub
Private Sub txtHead_Change()
    dispPreview
End Sub
Private Sub lbl01_Click()
    setBorder 1
End Sub
Private Sub lbl02_Click()
    setBorder 2
End Sub
Private Sub lbl03_Click()
    setBorder 3
End Sub
Private Sub lbl04_Click()
    setBorder 4
End Sub
Private Sub lbl05_Click()
    setBorder 5
End Sub
Private Sub lbl06_Click()
    setBorder 6
End Sub
Private Sub lbl07_Click()
    setBorder 7
End Sub
Private Sub lbl08_Click()
    setBorder 8
End Sub
Private Sub lbl09_Click()
    setBorder 9
End Sub
Private Sub lbl10_Click()
    setBorder 10
End Sub
Private Sub lbl11_Click()
    setBorder 11
End Sub
Private Sub lbl12_Click()
    setBorder 12
End Sub
Private Sub lbl13_Click()
    setBorder 13
End Sub
Private Sub lbl14_Click()
    setBorder 14
End Sub
Private Sub txtHoganJudgeLine_Change()
    dispPreview
End Sub
Private Sub txtName_Change()
    dispPreview
End Sub
Private Sub spnCol_SpinDown()
    txtCol.Text = spinDown(txtCol.Text, 0)
End Sub
Private Sub spnCol_SpinUp()
    txtCol.Text = spinUp(txtCol.Text)
End Sub
Private Sub spnHead_SpinDown()
    txtHead.Text = spinDown(txtHead.Text, 0)
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
Private Function spinDown(ByVal vntValue As Variant, ByVal limit As Long) As Variant

    Dim lngValue As Long
    
    lngValue = Val(vntValue)
    lngValue = lngValue - 1
    If lngValue < limit Then
        lngValue = limit
    End If
    spinDown = lngValue

End Function
Private Sub lblHead_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult
    
    lngColor = lblHead.BackColor
    
    Result = frmColor.Start(lngColor)
    
    If Result = vbOK Then
        lblHead.BackColor = lngColor
        dispPreview
    End If

End Sub
Private Sub lblEven_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult
    
    lngColor = lblEven.BackColor
    
    Result = frmColor.Start(lngColor)
    
    If Result = vbOK Then
        lblEven.BackColor = lngColor
        dispPreview
    End If
    
End Sub
'------------------------------------------------------------------------------------------------------------------------
' 再表示処理
'------------------------------------------------------------------------------------------------------------------------
Sub dispPreview()

    Dim i As Long
    Dim strBuf As String
    
    If mblnRefresh Then
        Exit Sub
    End If
    
    Dim s As KantanLine
    
    Set s = New KantanLine
    
    s.Text = txtName.Text
    s.OutSideLine = Val(imgOutSide.tag)
    s.VHeadBorderLine = Val(imgVHead.tag)
    s.HHeadBorderLine = Val(imgHHead.tag)
    s.VInsideLine = Val(imgVLine.tag)
    s.HInsideLine = Val(imgHLine.tag)

    s.HHeadLineCount = Val(txtHead.Text)
    s.VHeadLineCount = Val(txtCol.Text)
    
    s.HeadColor = lblHead.BackColor
    s.EvenColor = lblEven.BackColor
    
    s.EnableEvenColor = chkLineColor.value
    s.EnableHogan = chkHoganMode.value
    
    s.EnableHRepeat = chkHRepeat.value
    s.EnableVRepeat = chkVRepeat.value
    s.AuthoHogan = chkAutoHogan.value
    
    If Val(txtHoganJudgeLine.Text) = 0 Then
        s.HoganJudgeLineCount = 1
    Else
        s.HoganJudgeLineCount = Val(txtHoganJudgeLine.Text)
    End If
    
    imgPreview.Picture = editKantan(s, xlBitmap)
    
    '編集結果をリストに設定
    If lstStamp.ListCount = 0 Then
        Exit Sub
    End If

    i = lstStamp.ListIndex

    mblnRefresh = True
    
    lstStamp.List(i, C_TEXT) = s.Text
    strBuf = serialize(s)
    lstStamp.List(i, C_DATA) = strBuf
    
    lblEvenLabel.enabled = chkLineColor.value
    lblEven.enabled = chkLineColor.value
    
    chkAutoHogan.enabled = chkHoganMode.value
    
    txtHoganJudgeLine.enabled = Not (chkAutoHogan.value)
    spnHoganJudgeLine.enabled = Not (chkAutoHogan.value)
    lblJudge.enabled = Not (chkAutoHogan.value)

    chkHRepeat.enabled = (Val(txtHead.Text) > 1)
    chkVRepeat.enabled = (Val(txtCol.Text) > 1)

    mblnRefresh = False

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 線スタイル取得
'------------------------------------------------------------------------------------------------------------------------
Private Function getBorder() As Long
    
    Dim i As Long
    
    getBorder = 2
    
    For i = 1 To colBorder.count
        If colBorder(i).BorderStyle = fmBorderStyleSingle Then
            getBorder = i
            Exit Function
        End If
    Next

End Function
'------------------------------------------------------------------------------------------------------------------------
' 線スタイル設定
'------------------------------------------------------------------------------------------------------------------------
Private Sub setBorder(ByVal lngNo As Long)

    Dim i As Long
    
    For i = 1 To colBorder.count

        If i = lngNo Then
            colBorder(i).BorderStyle = fmBorderStyleSingle
        Else
            colBorder(i).BorderStyle = fmBorderStyleNone
        End If

    Next

End Sub
'------------------------------------------------------------------------------------------------------------------------
' 線スタイル設定
'------------------------------------------------------------------------------------------------------------------------
Private Function setPicture(ByRef obj As image, ByVal lngPos As Long)

    Select Case lngPos
        Case 1
            Set obj.Picture = Image1.Picture
        Case 2
            Set obj.Picture = Image2.Picture
        Case 3
            Set obj.Picture = Image3.Picture
        Case 4
            Set obj.Picture = Image4.Picture
        Case 5
            Set obj.Picture = Image5.Picture
        Case 6
            Set obj.Picture = Image6.Picture
        Case 7
            Set obj.Picture = Image7.Picture
        Case 8
            Set obj.Picture = Image8.Picture
        Case 9
            Set obj.Picture = Image9.Picture
        Case 10
            Set obj.Picture = Image10.Picture
        Case 11
            Set obj.Picture = Image11.Picture
        Case 12
            Set obj.Picture = Image12.Picture
        Case 13
            Set obj.Picture = Image13.Picture
        Case 14
            Set obj.Picture = Image14.Picture
    End Select
    obj.tag = lngPos
    
    obj.visible = False
    obj.visible = True
    
    dispPreview
    
End Function

