VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mCol As Collection
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
Sub Push(ByVal strBuf As String)
    mCol.Add strBuf
End Sub

Function Pop() As String
    If mCol.count = 0 Then
        Pop = ""
    Else
        Pop = mCol(mCol.count)
        mCol.Remove mCol.count
    End If
End Function

Function Value() As String
    If mCol.count = 0 Then
        Value = ""
    Else
        Value = mCol(mCol.count)
    End If
End Function

Function Eof() As Boolean
    If mCol.count = 0 Then
        Eof = True
    Else
        Eof = False
    End If
End Function
