VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Compare(a, b) As Integer
    If a Is Nothing Then
        Compare = -1
    ElseIf b Is Nothing Then
        Compare = 1
    ElseIf a.Count < b.Count Then
        Compare = -1
    ElseIf a.Count > b.Count Then
        Compare = 1
    Else
        Compare = 0
    End If
End Function

