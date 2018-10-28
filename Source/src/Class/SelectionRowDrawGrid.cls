VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionRowDrawGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 宣言
'------------------------------------------------------------------------------------------------------------------------
Private WithEvents SFWork As SelectionRowFrameWork
Attribute SFWork.VB_VarHelpID = -1

'------------------------------------------------------------------------------------------------------------------------
' メンバ変数宣言部(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private mrngEven As Range
Private mrngOdd As Range

Private mlngHeadColor As Long
Private mlngEvenColor As Long

Private mblnCustom As Boolean

Private mlngHead As Long
Private mlngCol As Long

Private mblnHogan As Boolean


'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 作成
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    'デフォルト１行
    mlngHead = 1
    mlngCol = 0
    Set SFWork = New SelectionRowFrameWork
End Sub
'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork 開放
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set SFWork = Nothing
End Sub
'------------------------------------------------------------------------------------------------------------------------
' SelectionFrameWork Run
'------------------------------------------------------------------------------------------------------------------------
Public Sub Run()
    SFWork.Run
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 初期処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionInit(rArea As Areas, Cancel As Boolean, Undo As Boolean)
    
    Dim Result As VbMsgBoxResult
    Dim lngHead As Long
    Dim lngHeadColor As Long
    Dim lngEvenColor As Long
    Dim blnHoganMode As Boolean
    
    On Error GoTo e
    
    Undo = True
    
    If mblnCustom Then
    Else
        Exit Sub
    End If
    
    lngHeadColor = mlngHeadColor
    lngEvenColor = mlngEvenColor
    blnHoganMode = False
    
    Result = frmGrid.Start(mlngHead, mlngCol, lngHeadColor, lngEvenColor, blnHoganMode)
    
    If Result = vbOK Then
        
        mlngHeadColor = lngHeadColor
        mlngEvenColor = lngEvenColor
        mblnHogan = blnHoganMode
        
    Else
        Cancel = True
        Exit Sub
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 前処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionBegin(rArea As Range, Cancel As Boolean)

    Dim lngCols As Long
    Dim lngRows As Long
    
    On Error GoTo e
    
    rArea.Interior.Color = xlNone

    ''外周の線を引く
    With rArea.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rArea.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rArea.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rArea.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    lngRows = rArea.Rows.Count
    lngCols = rArea.Columns.Count
        
    ''内部縦線
    If lngCols > 1 Then
        If mblnHogan Then
        
            Dim lngCol As Long
            Dim lngRow As Long
            
            If mlngHead = 0 Then
                lngRow = 1
            Else
                lngRow = mlngHead
            End If
            
            With rArea.Borders(xlInsideVertical)
                .LineStyle = xlNone
            End With
            
            For lngCol = 2 To rArea.Columns.Count
            
                If rArea.Cells(lngRow, lngCol).Value <> "" Then
                    With rArea.Columns(lngCol).Borders(xlLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
                End If
            
            Next
                    
        Else
            With rArea.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End If
    End If
    ''内部横線
    If lngRows > 1 Then
        With rArea.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlNone
        End With
    End If
    
    ''内部横線
    If lngRows > mlngHead + 1 Then

        Dim r As Range
        Set r = Range(rArea(mlngHead + 1, 1), rArea(lngRows, lngCols))

        With r.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
    End If
    
    'ヘッダ
    If mlngHead = 0 Then
    Else
        If lngRows < mlngHead Then
            mlngHead = lngRows
        End If
        
        Set r = Range(rArea(1, 1), rArea(mlngHead, lngCols))
    
        With r.Interior
            If mlngHeadColor = -1 Then
                .ColorIndex = xlNone
            Else
                .Color = mlngHeadColor
            End If
            .Pattern = xlSolid
        End With
        
        With r.Borders(xlInsideHorizontal)
'            .LineStyle = xlNone
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With r.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
    Cancel = True
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 主処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionMain(r As Range, ByVal Row As Long, Cancel As Boolean)
        
    On Error GoTo e
    
    If mlngEvenColor = -1 Then
        Set mrngEven = Nothing
        Set mrngOdd = Nothing
    Else
        Select Case True
            Case Row <= mlngHead
            Case (Row - mlngHead) Mod 2 = 0
                If mrngEven Is Nothing Then
                    Set mrngEven = r
                Else
                    Set mrngEven = Union(mrngEven, r)
                End If
                
            Case (Row - mlngHead) Mod 2 = 1
                If mrngOdd Is Nothing Then
                    Set mrngOdd = r
                Else
                    Set mrngOdd = Union(mrngOdd, r)
                End If
                
        End Select
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
    Cancel = True
End Sub

'------------------------------------------------------------------------------------------------------------------------
' 後処理(UOC)
'------------------------------------------------------------------------------------------------------------------------
Private Sub SFWork_SelectionFinal(rArea As Range)
    
    On Error GoTo e
    
    '偶数ライン
    If mrngEven Is Nothing Then
    Else
        With mrngEven.Interior
            If mlngEvenColor = -1 Then
                .ColorIndex = xlNone
            Else
                .Color = mlngEvenColor
            End If
            .Pattern = xlSolid
        End With
    End If
    Set mrngEven = Nothing
    
    '奇数ライン
    If mrngOdd Is Nothing Then
    Else
        With mrngOdd.Interior
            .ColorIndex = xlNone
            .Pattern = xlSolid
        End With
    End If
    Set mrngOdd = Nothing

    Dim lngCols As Long
    Dim lngRows As Long
    
    lngRows = rArea.Rows.Count
    lngCols = rArea.Columns.Count
    
    Dim r As Range
    
    '列ヘッダ
    If mlngCol = 0 Then
    Else
        If lngCols < mlngCol Then
            mlngCol = lngCols
        End If
        
        Set r = Range(rArea(1, 1), rArea(lngRows, mlngCol))
    
        With r.Interior
            If mlngHeadColor = -1 Then
                .ColorIndex = xlNone
            Else
                .Color = mlngHeadColor
            End If
            .Pattern = xlSolid
        End With
        
        With r.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With r.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub

Public Property Let HeadColor(ByVal lngColor As Long)

    mlngHeadColor = lngColor

End Property
Public Property Let EvenColor(ByVal lngColor As Long)

    mlngEvenColor = lngColor

End Property
Public Property Let Custom(ByVal blnCustom As Long)

    mblnCustom = blnCustom

End Property
Public Property Let HeadLine(ByVal lngHeadLine As Long)

    mlngHead = lngHeadLine

End Property
Public Property Let ColLine(ByVal lngColLine As Long)

    mlngCol = lngColLine

End Property
Public Property Let HoganMode(ByVal blnHogan As Boolean)

    mblnHogan = blnHogan

End Property
