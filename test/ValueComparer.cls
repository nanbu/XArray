VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Compare(a, b) As Integer
    If a < b Then
        Compare = -1
    ElseIf a > b Then
        Compare = 1
    Else
        Compare = 0
    End If
End Function
