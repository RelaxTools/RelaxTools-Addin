Attribute VB_Name = "basItem"
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
Option Private Module

Public mColAllItem As Collection
'--------------------------------------------------------------
'　段落番号用オブジェクト生成
'--------------------------------------------------------------
Public Sub createAllItemObject()

    Set mColAllItem = New Collection

    'すべてのセクション
    mColAllItem.Add New itemPoint
    mColAllItem.Add New itemCircleB
    mColAllItem.Add New itemCircleW
    mColAllItem.Add New itemDiaB
    mColAllItem.Add New itemDiaW
    mColAllItem.Add New itemRevTriB
    mColAllItem.Add New itemRevTriW
    mColAllItem.Add New itemSquareB
    mColAllItem.Add New itemSquareW
    mColAllItem.Add New itemNum1
    mColAllItem.Add New itemImp
    mColAllItem.Add New itemDouble
    mColAllItem.Add New itemStarB
    mColAllItem.Add New itemStarW
    mColAllItem.Add New itemSime
    mColAllItem.Add New itemDblR
    mColAllItem.Add New itemNumExp
    mColAllItem.Add New itemNumC
    mColAllItem.Add New itemNumA

End Sub
'--------------------------------------------------------------
'　段落番号取得
'--------------------------------------------------------------
Function rlxGetItemNoAny(ByVal strBuf As String) As String
Attribute rlxGetItemNoAny.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetItemNoAny.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim i As Long
    Dim lngCnt As Long
    Dim strSecNo As String
    Dim blnFind As Boolean
    
    blnFind = False

    lngCnt = Len(strBuf)
    If lngCnt = 0 Then
        Exit Function
    End If
    
    strSecNo = ""
    
    If mColAllItem Is Nothing Then
        Call createAllItemObject
    End If

    For i = 1 To mColAllItem.count
    
        strSecNo = mColAllItem(i).SectionNumber(strBuf)
        If Len(strSecNo) > 0 Then
            Exit For
        End If
        
    Next

    rlxGetItemNoAny = strSecNo

End Function
'--------------------------------------------------------------
'　段落番号取得
'--------------------------------------------------------------
Private Function getItemNo(ByVal strBuf As String, ByVal strItemName As String) As String

    Dim obj As Object

    Set obj = rlxGetItemObject(strItemName)

    getItemNo = obj.SectionNumber(strBuf)
    
    Set obj = Nothing
    
End Function
'--------------------------------------------------------------
'　段落番号があるかどうか判定
'--------------------------------------------------------------
Function rlxHasItemNo(ByVal strBuf As String, ByVal strItemName As String) As Boolean
Attribute rlxHasItemNo.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxHasItemNo.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim i As Long
    Dim lngCnt As Long
    
    rlxHasItemNo = False
    
    lngCnt = Len(strBuf)
    If lngCnt = 0 Then
        Exit Function
    End If
    
    If Len(getItemNo(strBuf, strItemName)) > 0 Then
        rlxHasItemNo = True
    End If
    
End Function
'--------------------------------------------------------------
'　次番号の取得
'--------------------------------------------------------------
Function rlxGetItemNext(ByVal strBuf As String, ByVal strItemName As String) As String
Attribute rlxGetItemNext.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetItemNext.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim obj As Object
    Dim lngIndentLevel As Long

    Set obj = rlxGetItemObject(strItemName)

    '次番号の取得
    rlxGetItemNext = obj.NextNumber(strBuf, lngIndentLevel)
    
    Set obj = Nothing

End Function
'--------------------------------------------------------------
'　段落番号の設定
'--------------------------------------------------------------
Sub setItemNo(ByRef r As Range, ByVal strNewNo As String)

    Dim lngPos As Long
    Dim obj As Object

    If VarType(r.Value) = vbString Then
        r.Characters(0, 0).Insert strNewNo
    Else
        If Len(r.PrefixCharacter) > 0 Then
            r.Value = r.PrefixCharacter & strNewNo & r.Value
        Else
            r.Value = strNewNo & r.Value
        End If
    End If
    
End Sub
'--------------------------------------------------------------
'　段落番号の削除
'--------------------------------------------------------------
Sub delItemNo(ByRef r As Range)

    Dim strSecNo As String
    Dim lngPos As Long
    Dim obj As Object

    '現在の段落番号を取得（レベルにかかわらない）
    strSecNo = rlxGetItemNoAny(r.Value)
    If VarType(r.Value) = vbString Then
        If Len(strSecNo) > 0 Then
            r.Characters(1, Len(strSecNo)).Delete
        End If
    Else
        If Len(strSecNo) > 0 Then
            If Len(r.PrefixCharacter) > 0 Then
                r.Value = r.PrefixCharacter & Mid$(r.Value, Len(strSecNo) + 1)
            Else
                r.Value = Mid$(r.Value, Len(strSecNo) + 1)
            End If
        End If
    End If

End Sub
'--------------------------------------------------------------
'　クラス名からオブジェクトを取得
'--------------------------------------------------------------
Function rlxGetItemObject(ByVal className As String) As Object
Attribute rlxGetItemObject.VB_Description = "ワークシート関数として使用できません。"
Attribute rlxGetItemObject.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim ret As Object
    Dim obj As Object
    Set ret = Nothing
    
    If mColAllItem Is Nothing Then
        Call createAllItemObject
    End If
    
    For Each obj In mColAllItem
    
        If className = obj.Class Then
            Set ret = obj
            Exit For
        End If
    
    Next

    Set rlxGetItemObject = ret

End Function

