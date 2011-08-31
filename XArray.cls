VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MyArray
Private MyCount As Long
Private MyCapacity As Long
Private MyCapacityStep As Long
Private MyCompareMode As VbCompareMethod

Private Sub Class_Initialize()
    MyCapacityStep = 1024
    MyCount = 0
    MyCapacity = MyCapacityStep
    MyCompareMode = vbBinaryCompare
    ReDim MyArray(MyCapacity - 1)
End Sub

Private Sub Class_Terminate()
    Erase MyArray
End Sub

Private Function CountUp()
    MyCount = MyCount + 1
    If MyCount > MyCapacity Then
        MyCapacity = MyCapacity + MyCapacityStep
        ReDim Preserve MyArray(MyCapacity - 1)
    End If
End Function

Private Function CountDown()
    MyArray(MyCount - 1) = Empty
    MyCount = MyCount - 1
    If MyCount + MyCapacityStep <= MyCapacity Then
        MyCapacity = MyCapacity - MyCapacityStep
        If MyCapacity <= MyCapacityStep Then
            MyCapacity = MyCapacityStep
        End If
        ReDim Preserve MyArray(MyCapacity - 1)
    End If
End Function

Private Function SetValue(Target, Val)
    If IsObject(Val) Then
        Set Target = Val
    Else
        Target = Val
    End If
End Function

Property Get Count() As Long
    Count = MyCount
End Property

Function Add(Val)
    Call CountUp
    Call SetValue(MyArray(MyCount - 1), Val)
End Function

Public Property Get Item(Index As Long) As Variant
    If Index >= MyCount Then
        Err.Raise 9, "XArray.Item(Index)"
    Else
        Call SetValue(Item, MyArray(Index))
    End If
End Property

Property Let Item(Index As Long, Value)
    If Index >= MyCount Then
        Err.Raise 9, "XArray.Item(Index, Value)"
    Else
        MyArray(Index) = Value
    End If
End Property

Property Set Item(Index As Long, Value)
    If Index >= MyCount Then
        Err.Raise 9, "XArray.Item(Index, Value)"
    Else
        Set MyArray(Index) = Value
    End If
End Property

Property Get Items() As Variant
    Dim Result As Variant
    If MyCount <= 0 Then
        Result = Array()
    Else
        Result = MyArray
        ReDim Preserve Result(MyCount - 1)
    End If
    Items = Result
End Property

Property Get IndexOf(Val, Optional Comparer) As Long
    Dim Result As Long
    Result = -1
    Dim i As Long
    If IsMissing(Comparer) Then
        If MyCompareMode = vbTextCompare Then
            For i = 0 To MyCount - 1
                If StrComp(MyArray(i), Val, vbTextCompare) = 0 Then
                    Result = i
                    Exit For
                End If
            Next
        Else
            For i = 0 To MyCount - 1
                If MyArray(i) = Val Then
                    Result = i
                    Exit For
                End If
            Next
        End If
    Else
        For i = 0 To MyCount - 1
            If Comparer.Compare(MyArray(i), Val) = 0 Then
                Result = i
                Exit For
            End If
        Next
    End If
    IndexOf = Result
End Property

Property Get Exists(Val, Optional Comparer) As Boolean
    Exists = (IndexOf(Val, Comparer) >= 0)
End Property

Function Insert(Index As Long, Value As Variant)
    If Index > MyCount Or Index < 0 Then
        Err.Raise 9, "XArray.Insert(Index, Value)"
    Else
        Call CountUp
        Dim i As Long
        For i = MyCount - 1 To Index + 1 Step -1
            Call SetValue(MyArray(i), MyArray(i - 1))
        Next
        Call SetValue(MyArray(Index), Value)
    End If
End Function

Function Remove(Index As Long)
    If Index >= MyCount Then
        Err.Raise 9, "XArray.Remove(index)"
    Else
        Dim i As Long
        If IsObject(MyArray(i)) Then
            Set MyArray(i) = Nothing
        End If
        For i = Index To MyCount - 2
            Call SetValue(MyArray(i), MyArray(i + 1))
        Next
        Call CountDown
    End If
End Function

Function Exchange(Index1 As Long, Index2 As Long)
    If Index1 < 0 Or Index1 >= MyCount Or Index2 < 0 Or Index2 >= MyCount Then
        Err.Raise 9, "XArray.Exchange(Index1, Index2)"
        Exit Function
    End If
    Dim Temp
    Call SetValue(Temp, MyArray(Index1))
    Call SetValue(MyArray(Index1), MyArray(Index2))
    Call SetValue(MyArray(Index2), Temp)
End Function

Function Reverse()
    Dim Result
    ReDim Result(MyCapacity - 1)
    Dim i As Long, n As Long
    i = 0
    For n = MyCount - 1 To 0 Step -1
        Call SetValue(Result(i), MyArray(n))
        i = i + 1
    Next
    Erase MyArray
    MyArray = Result
End Function

Function Sort(Optional Comparer)
    Dim Min As Long
    Dim Max As Long
    Min = 0
    Max = MyCount - 1
    Call QuickSort(MyArray, Min, Max, Comparer)
End Function

Private Function QuickSort(Target, Min As Long, Max As Long, Optional Comparer)
    If Min >= Max Then
        Exit Function
    End If

    Dim Middle As Long
    Dim Index1 As Long
    Dim Index2 As Long
    Dim Temp

    Middle = (Min + Max) \ 2
    Call SetValue(Temp, Target(Middle))
    Call SetValue(Target(Middle), Target(Min))
    Index2 = Min
    Index1 = Min + 1
    If IsMissing(Comparer) Then
        If MyCompareMode = vbTextCompare Then
            Do While Index1 <= Max
                If StrComp(Target(Index1), Temp, vbTextCompare) < 0 Then
                    Index2 = Index2 + 1
                    Call Exchange(Index1, Index2)
                End If
                Index1 = Index1 + 1
            Loop
        Else
            Do While Index1 <= Max
                If Target(Index1) < Temp Then
                    Index2 = Index2 + 1
                    Call Exchange(Index1, Index2)
                End If
                Index1 = Index1 + 1
            Loop
        End If
    Else
        Do While Index1 <= Max
            If Comparer.Compare(Target(Index1), Temp) < 0 Then
                Index2 = Index2 + 1
                Call Exchange(Index1, Index2)
            End If
            Index1 = Index1 + 1
        Loop
    End If
    Call SetValue(Target(Min), Target(Index2))
    Call SetValue(Target(Index2), Temp)
    Call QuickSort(Target, Min, Index2 - 1, Comparer)
    Call QuickSort(Target, Index2 + 1, Max, Comparer)
End Function

Function Clone() As XArray
    Dim Result As XArray
    Set Result = New XArray
    Dim i As Long
    For i = 0 To MyCount - 1
        Result.Add MyArray(i)
    Next
    Set Clone = Result
End Function

Property Let CompareMode(Mode As VbCompareMethod)
    MyCompareMode = Mode
End Property

Property Get CompareMode() As VbCompareMethod
    CompareMode = MyCompareMode
End Property
