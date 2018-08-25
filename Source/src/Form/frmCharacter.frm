VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCharacter 
   Caption         =   "指定文字の文字修飾"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   OleObjectBlob   =   "frmCharacter.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdDef_Click()
    Modify True
End Sub
Private Sub cmdRun_Click()
    Modify False
End Sub
Private Sub Modify(ByVal def As Boolean)

    Dim v As Variant
    Dim s As Variant
    Dim r As Range
    Dim lngStart As Long
    Dim lngCount As Long
    Dim lngCng As Long

    Dim strBuf As String
    Dim rr As Range
    
    strBuf = Replace(txtText.Text, vbCrLf, vbLf)
    v = Split(strBuf, vbLf)
              
    Select Case True
        Case optSheet.Value
            Set rr = SpecialCellsEx(ActiveSheet.UsedRange)
        Case optSelection.Value
            Set rr = SpecialCellsEx(Selection)
    End Select
    
    On Error GoTo e
    Application.ScreenUpdating = False
    
    lngCount = 0
    lngCng = 0
    
    StartBar "", rr.count
    For Each r In rr
    
        For Each s In v
        
            '空の場合パス
            If Len(s) = 0 Then
                Exit For
            End If
        
            lngStart = InStr(r.Value, s)
            Do Until lngStart = 0
            
                If def Then
                    With r.Characters(lngStart, Len(s)).Font
                        .ColorIndex = xlAutomatic
                        .Bold = False
                        .Italic = False
                        .Underline = False
                    End With
                Else
                    With r.Characters(lngStart, Len(s)).Font
                        .Color = lblColor.BackColor
                        .Bold = cmdBold.Value
                        .Italic = cmdItalic.Value
                        .Underline = cmdUnderline.Value
                    End With
                End If
                
                lngCng = lngCng + 1
                
                lngStart = InStr(lngStart + 1, r.Value, s)
                
            Loop
        Next
        lngCount = lngCount + 1
        ReportBar lngCount
    Next
    StopBar
    
    Call SaveSetting(C_TITLE, "Character", "Text", txtText.Text)
    Call SaveSetting(C_TITLE, "Character", "Bold", cmdBold.Value)
    Call SaveSetting(C_TITLE, "Character", "Italic", cmdItalic.Value)
    Call SaveSetting(C_TITLE, "Character", "Underline", cmdUnderline.Value)
    Call SaveSetting(C_TITLE, "Character", "Color", CLng(lblColor.BackColor))
    Call SaveSetting(C_TITLE, "Character", "ActiveSheet", optSheet.Value)
    
    lblStatus.Caption = lngCng & "件処理しました。"
e:
    Application.ScreenUpdating = True
    
End Sub



Private Sub lblColor_Click()

    Dim lngColor As Long
    Dim Result As VbMsgBoxResult

    lngColor = lblColor.BackColor

    Result = frmColor.Start(lngColor)

    If Result = vbOK Then
        lblColor.BackColor = lngColor
    End If
    
End Sub


Private Sub UserForm_Initialize()

    txtText.Text = GetSetting(C_TITLE, "Character", "Text", "")
    cmdBold.Value = GetSetting(C_TITLE, "Character", "Bold", False)
    cmdItalic.Value = GetSetting(C_TITLE, "Character", "Italic", False)
    cmdUnderline.Value = GetSetting(C_TITLE, "Character", "Underline", False)
    
    If GetSetting(C_TITLE, "Character", "ActiveSheet", True) Then
        optSheet.Value = True
    Else
        optSelection.Value = True
    End If
    
    lblColor.BackColor = CLng(GetSetting(C_TITLE, "Character", "Color", vbRed))
    
    lblStatus.Caption = "文字修飾を行う文字列と文字属性を入力してください。"
    

End Sub
Private Sub StartBar(ByVal strMsg As String, ByVal lngMax As Long)

    lblBar.visible = True
    lblBar.width = 0
    lblBar.Caption = strMsg
    
    lblStatus.Caption = strMsg
    lblStatus.Tag = strMsg
    lblBar.Tag = lngMax
    
    'ReportBar 1

End Sub
Private Sub ReportBar(ByVal lngPos As Long)

    Dim dblPercent As Double
    
    dblPercent = (lngPos / Val(lblBar.Tag))

    lblBar.width = lblStatus.width * dblPercent
    lblBar.Caption = lblStatus.Tag & " 処理中です..." & Fix(dblPercent * 100) & "%"
    lblStatus.Caption = lblBar.Caption
    DoEvents

End Sub

Private Sub StopBar()
    
    lblBar.visible = False
    lblBar.width = 0
    lblBar.Caption = ""
    
    lblStatus.Caption = ""

End Sub

