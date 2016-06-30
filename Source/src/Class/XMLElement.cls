VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EnumElementType
    StartTag
    EndTag
    StandAlone
    Header
    Data
    CData
    Comment
End Enum


Public Element As String
Public ElementName As String
Public ElementType As EnumElementType


