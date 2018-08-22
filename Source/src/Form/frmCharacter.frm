VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCharacter 
   Caption         =   "指定文字の文字修飾"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390.001
   OleObjectBlob   =   "frmCharacter.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()

    Dim v As Variant
    Dim s As Variant
    Dim r As Range
    Dim lngStart As Long

    Dim strBuf As String
    Dim rr As Range
    
    strBuf = Replace(txtText.Text, vbCrLf, vbLf)
    v = Split(strBuf, vbLf)
              
    Set rr = SpecialCellsEx(ActiveSheet.UsedRange)
    
    On Error GoTo e
    Application.ScreenUpdating = False
    
    For Each r In rr
    
        For Each s In v
        
            '空の場合パス
            If Len(s) = 0 Then
                Exit For
            End If
        
            lngStart = InStr(r.Value, s)
            Do Until lngStart = 0
            
                With r.Characters(lngStart, Len(s)).Font
                    .Color = lblColor.BackColor
                    .Bold = cmdBold.Value
                    .Italic = cmdItalic.Value
                    .Underline = cmdUnderline.Value
                End With
                
                lngStart = InStr(lngStart + 1, r.Value, s)
            Loop
        Next
    
    Next
    
    Call SaveSetting(C_TITLE, "Character", "Text", txtText.Text)
    Call SaveSetting(C_TITLE, "Character", "Bold", cmdBold.Value)
    Call SaveSetting(C_TITLE, "Character", "Italic", cmdItalic.Value)
    Call SaveSetting(C_TITLE, "Character", "Underline", cmdUnderline.Value)
    Call SaveSetting(C_TITLE, "Character", "Color", CLng(lblColor.BackColor))
    
e:
    Application.ScreenUpdating = True
'    Unload Me
    
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
    
    lblColor.BackColor = CLng(GetSetting(C_TITLE, "Character", "Color", vbRed))

End Sub
