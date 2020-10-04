VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSectionList 
   Caption         =   "段落番号リスト"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090.001
   OleObjectBlob   =   "frmSectionList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSectionList"
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
Private mColSection1 As New Collection
Private mColSection2 As New Collection
Private mColSection3 As New Collection
Private mColSection4 As New Collection
Private mColSection5 As New Collection
Private mColSection6 As New Collection

Private mColLevel01 As New Collection
Private mColLevel02 As New Collection
Private mColLevel03 As New Collection
Private mColLevel04 As New Collection
Private mColLevel05 As New Collection
Private mColLevel06 As New Collection



Private Sub cmdOk_Click()


    
    Unload Me
End Sub

Private Sub cmdSetting01_Click()

    Dim ret As Collection
    Dim i As Long
    
    Set ret = frmSectionEx.Start(mColSection1)
    If ret Is Nothing Then
    Else
        Set mColSection1 = ret
        For i = 1 To 6
            If mColSection1.Count < i Then
                mColLevel01(i).Caption = ""
            Else
                If mColSection1(i).classObj Is Nothing Then
                    mColLevel01(i).Caption = ""
                Else
                    mColLevel01(i).Caption = mColSection1(i).classObj.SectionLevelName(i)
                End If
            End If
        Next
        setSectionSetting "01", mColSection1
    End If
End Sub

Private Sub cmdSetting02_Click()

    Dim ret As Collection
    Dim i As Long
    
    Set ret = frmSectionEx.Start(mColSection2)
    If ret Is Nothing Then
    Else
        Set mColSection2 = ret
        For i = 1 To 6
            If mColSection2.Count < i Then
                mColLevel02(i).Caption = ""
            Else
                If mColSection2(i).classObj Is Nothing Then
                    mColLevel02(i).Caption = ""
                Else
                    mColLevel02(i).Caption = mColSection2(i).classObj.SectionLevelName(i)
                End If
            End If
        Next
        setSectionSetting "02", mColSection2
    End If
End Sub

Private Sub cmdSetting03_Click()

    Dim ret As Collection
    Dim i As Long
    
    Set ret = frmSectionEx.Start(mColSection3)
    If ret Is Nothing Then
    Else
        Set mColSection3 = ret
        For i = 1 To 6
            If mColSection3.Count < i Then
                mColLevel03(i).Caption = ""
            Else
                If mColSection3(i).classObj Is Nothing Then
                    mColLevel03(i).Caption = ""
                Else
                    mColLevel03(i).Caption = mColSection3(i).classObj.SectionLevelName(i)
                End If
            End If
        Next
        setSectionSetting "03", mColSection3
    End If
    
End Sub

Private Sub cmdSetting04_Click()
    Dim ret As Collection
    Dim i As Long
    
    Set ret = frmSectionEx.Start(mColSection4)
    If ret Is Nothing Then
    Else
        Set mColSection4 = ret
        For i = 1 To 6
            If mColSection4.Count < i Then
                mColLevel04(i).Caption = ""
            Else
                If mColSection4(i).classObj Is Nothing Then
                    mColLevel04(i).Caption = ""
                Else
                    mColLevel04(i).Caption = mColSection4(i).classObj.SectionLevelName(i)
                End If
            End If
        Next
        setSectionSetting "04", mColSection4
    End If

End Sub

Private Sub cmdSetting05_Click()
    Dim ret As Collection
    Dim i As Long
    
    Set ret = frmSectionEx.Start(mColSection5)
    If ret Is Nothing Then
    Else
        Set mColSection5 = ret
        For i = 1 To 6
            If mColSection5.Count < i Then
                mColLevel05(i).Caption = ""
            Else
                If mColSection5(i).classObj Is Nothing Then
                    mColLevel05(i).Caption = ""
                Else
                    mColLevel05(i).Caption = mColSection5(i).classObj.SectionLevelName(i)
                End If
            End If
        Next
        setSectionSetting "05", mColSection5
    End If

End Sub

Private Sub cmdSetting06_Click()
    Dim ret As Collection
    Dim i As Long
    
    Set ret = frmSectionEx.Start(mColSection6)
    If ret Is Nothing Then
    Else
        Set mColSection6 = ret
        For i = 1 To 6
            If mColSection6.Count < i Then
                mColLevel06(i).Caption = ""
            Else
                If mColSection6(i).classObj Is Nothing Then
                    mColLevel06(i).Caption = ""
                Else
                    mColLevel06(i).Caption = mColSection6(i).classObj.SectionLevelName(i)
                End If
            End If
        Next
        setSectionSetting "06", mColSection6
    End If
End Sub


Private Sub UserForm_Initialize()

    Dim i As Long
    Dim obj As Object
    
    mColLevel01.Add lblLevel0101, "01"
    mColLevel01.Add lblLevel0102, "02"
    mColLevel01.Add lblLevel0103, "03"
    mColLevel01.Add lblLevel0104, "04"
    mColLevel01.Add lblLevel0105, "05"
    mColLevel01.Add lblLevel0106, "06"
    
    mColLevel02.Add lblLevel0201, "01"
    mColLevel02.Add lblLevel0202, "02"
    mColLevel02.Add lblLevel0203, "03"
    mColLevel02.Add lblLevel0204, "04"
    mColLevel02.Add lblLevel0205, "05"
    mColLevel02.Add lblLevel0206, "06"
    
    mColLevel03.Add lblLevel0301, "01"
    mColLevel03.Add lblLevel0302, "02"
    mColLevel03.Add lblLevel0303, "03"
    mColLevel03.Add lblLevel0304, "04"
    mColLevel03.Add lblLevel0305, "05"
    mColLevel03.Add lblLevel0306, "06"
    
    mColLevel04.Add lblLevel0401, "01"
    mColLevel04.Add lblLevel0402, "02"
    mColLevel04.Add lblLevel0403, "03"
    mColLevel04.Add lblLevel0404, "04"
    mColLevel04.Add lblLevel0405, "05"
    mColLevel04.Add lblLevel0406, "06"
    
    mColLevel05.Add lblLevel0501, "01"
    mColLevel05.Add lblLevel0502, "02"
    mColLevel05.Add lblLevel0503, "03"
    mColLevel05.Add lblLevel0504, "04"
    mColLevel05.Add lblLevel0505, "05"
    mColLevel05.Add lblLevel0506, "06"
    
    mColLevel06.Add lblLevel0601, "01"
    mColLevel06.Add lblLevel0602, "02"
    mColLevel06.Add lblLevel0603, "03"
    mColLevel06.Add lblLevel0604, "04"
    mColLevel06.Add lblLevel0605, "05"
    mColLevel06.Add lblLevel0606, "06"
    
    Dim strClass As String
    Dim ss As SectionStructDTO
    Dim col As Collection
    
    
    
    'カスタム（１）
    Set mColSection1 = rlxGetSectionSetting("01")
    For i = 1 To 6
        If mColSection1.Count < i Then
            mColLevel01(i).Caption = ""
        Else
            If mColSection1(i).classObj Is Nothing Then
                mColLevel01(i).Caption = ""
            Else
                mColLevel01(i).Caption = mColSection1(i).classObj.SectionLevelName(i)
            End If
        End If
    Next
    
    'カスタム（２）
    Set mColSection2 = rlxGetSectionSetting("02")
    For i = 1 To 6
        If mColSection2.Count < i Then
            mColLevel02(i).Caption = ""
        Else
            If mColSection2(i).classObj Is Nothing Then
                mColLevel02(i).Caption = ""
            Else
                mColLevel02(i).Caption = mColSection2(i).classObj.SectionLevelName(i)
            End If
        End If
    Next
    
    'カスタム（３）
    Set mColSection3 = rlxGetSectionSetting("03")
    For i = 1 To 6
        If mColSection3.Count < i Then
            mColLevel03(i).Caption = ""
        Else
            If mColSection3(i).classObj Is Nothing Then
                mColLevel03(i).Caption = ""
            Else
                mColLevel03(i).Caption = mColSection3(i).classObj.SectionLevelName(i)
            End If
        End If
    Next
    
    'カスタム（４）
    Set mColSection4 = rlxGetSectionSetting("04")
    For i = 1 To 6
        If mColSection4.Count < i Then
            mColLevel04(i).Caption = ""
        Else
            If mColSection4(i).classObj Is Nothing Then
                mColLevel04(i).Caption = ""
            Else
                mColLevel04(i).Caption = mColSection4(i).classObj.SectionLevelName(i)
            End If
        End If
    Next
    
    'カスタム（５）
    Set mColSection5 = rlxGetSectionSetting("05")
    For i = 1 To 6
        If mColSection5.Count < i Then
            mColLevel05(i).Caption = ""
        Else
            If mColSection5(i).classObj Is Nothing Then
                mColLevel05(i).Caption = ""
            Else
                mColLevel05(i).Caption = mColSection5(i).classObj.SectionLevelName(i)
            End If
        End If
    Next
    
    'カスタム（６）
    Set mColSection6 = rlxGetSectionSetting("06")
    For i = 1 To 6
        If mColSection6.Count < i Then
            mColLevel06(i).Caption = ""
        Else
            If mColSection6(i).classObj Is Nothing Then
                mColLevel06(i).Caption = ""
            Else
                mColLevel06(i).Caption = mColSection6(i).classObj.SectionLevelName(i)
            End If
        End If
    Next

    
    Dim strPos As String
    strPos = GetSetting(C_TITLE, "Section", "pos", "1")
    
    Select Case strPos
        Case "2"
            lblTile02.BackColor = &HC0FFFF
        Case "3"
            lblTile03.BackColor = &HC0FFFF
        Case "4"
            lblTile04.BackColor = &HC0FFFF
        Case "5"
            lblTile05.BackColor = &HC0FFFF
        Case "6"
            lblTile06.BackColor = &HC0FFFF
        Case Else
            lblTile01.BackColor = &HC0FFFF
    End Select


End Sub

Private Sub UserForm_Terminate()

    Dim strPos As String
    
    strPos = GetSetting(C_TITLE, "Section", "pos", "1")
    
    Select Case strPos
        Case "1"
            Set mColSection = mColSection1
        Case "2"
            Set mColSection = mColSection2
        Case "3"
            Set mColSection = mColSection3
        Case "4"
            Set mColSection = mColSection4
        Case "5"
            Set mColSection = mColSection5
        Case "6"
            Set mColSection = mColSection6
    End Select
    
'    setSectionSetting "01", mColSection1
'    setSectionSetting "02", mColSection2
'    setSectionSetting "03", mColSection3
'    setSectionSetting "04", mColSection4
'    setSectionSetting "05", mColSection5
'    setSectionSetting "06", mColSection6
    
End Sub
