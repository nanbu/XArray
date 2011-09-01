Attribute VB_Name = "XArrayTest"
Option Explicit

Sub Test()
    Dim StartTime As Currency
    Dim EndTime As Currency
    Debug.Print "**** XArrayTest.Test ****"
    StartTime = Timer()
    Call TestAdd
    Call TestInsert
    Call TestRemove
    Call TestItem
    Call TestCount
    Call TestExists
    Call TestIndexOf
    Call TestExchange
    Call TestSort
    Call TestReverse
    Call TestClone
    Call TestItems
    EndTime = Timer() - StartTime
    Debug.Print "Total: " & EndTime & " seconds. "
End Sub

Sub TestAdd()
    Dim StartTime As Currency
    Dim EndTime As Currency
    Dim a As New XArray
    Dim i As Long
    Dim Count As Long
    
    Count = 100000
    StartTime = Timer()
    For i = 0 To Count - 1
        a.Add i
    Next
    
    EndTime = Timer() - StartTime
    Debug.Print EndTime
    Debug.Assert EndTime < 10
    Debug.Assert a.Count = Count
End Sub

Sub TestInsert()
    Dim a As New XArray
    
    a.Add 1
    a.Add 2
    a.Add 3

    a.Insert 0, "A"
    Debug.Assert a.Count = 4
    Debug.Assert a.Item(0) = "A"
    Debug.Assert a.Item(1) = 1
    Debug.Assert a.Item(2) = 2
    Debug.Assert a.Item(3) = 3
    
    a.Insert a.Count - 1, "B"
    Debug.Assert a.Count = 5
    Debug.Assert a.Item(0) = "A"
    Debug.Assert a.Item(1) = 1
    Debug.Assert a.Item(2) = 2
    Debug.Assert a.Item(3) = "B"
    Debug.Assert a.Item(4) = 3
    
    a.Insert 2, "C"
    Debug.Assert a.Count = 6
    Debug.Assert a.Item(0) = "A"
    Debug.Assert a.Item(1) = 1
    Debug.Assert a.Item(2) = "C"
    Debug.Assert a.Item(3) = 2
    Debug.Assert a.Item(4) = "B"
    Debug.Assert a.Item(5) = 3
    
    Dim b As New XArray
    b.Insert 0, "A"
    Debug.Assert b.Count = 1
    Debug.Assert b.Item(0) = "A"
    
    Dim StartTime As Currency
    Dim EndTime As Currency
    Dim c As New XArray
    Dim i As Long
    Dim Count As Long
    
    Count = 100000
    For i = 0 To Count - 1
        c.Add i
    Next
    
    StartTime = Timer()
    c.Insert 0, "A"
    EndTime = Timer() - StartTime
    Debug.Print EndTime
    Debug.Assert EndTime < 10
    Debug.Assert c.Item(0) = "A"
    Debug.Assert c.Item(1) = 0
    Debug.Assert c.Item(Count) = Count - 1
    Debug.Assert c.Count = Count + 1
End Sub

Sub TestRemove()
    Dim StartTime As Currency
    Dim EndTime As Currency
    Dim a As New XArray
    Dim i As Long
    Dim Count As Long
    
    a.Add "A"
    a.Add "B"
    a.Add "C"
    a.Add "D"
    a.Add "E"
    Debug.Assert a.Count = 5
    
    a.Remove 4
    Debug.Assert a.Count = 4
    Debug.Assert a.Item(0) = "A"
    Debug.Assert a.Item(1) = "B"
    Debug.Assert a.Item(2) = "C"
    Debug.Assert a.Item(3) = "D"
    
    a.Remove 0
    Debug.Assert a.Count = 3
    Debug.Assert a.Item(0) = "B"
    Debug.Assert a.Item(1) = "C"
    Debug.Assert a.Item(2) = "D"
    
    a.Remove 1
    Debug.Assert a.Count = 2
    Debug.Assert a.Item(0) = "B"
    Debug.Assert a.Item(1) = "D"
    
    a.Remove 1
    Debug.Assert a.Count = 1
    Debug.Assert a.Item(0) = "B"
    
    a.Remove 0
    Debug.Assert a.Count = 0
    Count = 1000
    For i = 0 To Count - 1
        a.Add i
    Next
    StartTime = Timer()
    
    For i = 0 To Count - 1
        a.Remove 0
    Next
    EndTime = Timer() - StartTime
    Debug.Assert a.Count = 0
    Debug.Print EndTime
    Debug.Assert EndTime < 10
End Sub

Sub TestItem()
    Dim a As New XArray
    
    a.Add "A"
    Debug.Assert a.Item(0) = "A"
    Debug.Assert a.Count = 1
    
    a.Item(0) = "Z"
    Debug.Assert a.Item(0) = "Z"
    Debug.Assert a.Count = 1
    
    Set a.Item(0) = Nothing
    Debug.Assert a.Item(0) Is Nothing
    
    a.Item(0) = True
    Debug.Assert a.Item(0) = True
    
    Set a.Item(0) = New Collection
    Debug.Assert TypeName(a.Item(0)) = TypeName(New Collection)
    
    a.Item(0) = Empty
    Debug.Assert IsEmpty(a.Item(0))
    
    Dim StartTime As Currency
    Dim EndTime As Currency
    Dim i As Long
    Dim Count As Long
    Dim b As New XArray
    
    Count = 10000
    For i = 0 To Count - 1
        b.Add i
    Next
    
    StartTime = Timer()
    For i = 0 To Count - 1
        Debug.Assert b.Item(i) = i
    Next
    EndTime = Timer() - StartTime
    Debug.Assert b.Count = Count
    Debug.Print EndTime
    Debug.Assert EndTime < 10
End Sub

Sub TestCount()
    Dim a As New XArray
    Debug.Assert a.Count = 0
    a.Add "A"
    Debug.Assert a.Count = 1
    a.Add "B"
    Debug.Assert a.Count = 2
    a.Add "C"
    Debug.Assert a.Count = 3
    a.Remove 2
    Debug.Assert a.Count = 2
    a.Remove 0
    Debug.Assert a.Count = 1
End Sub

Sub TestExists()
    Dim a As New XArray
    a.Add 10
    a.Add 2
    a.Add 0
    a.Add 2
    a.Add 5
    a.CompareMode = vbBinaryCompare
    Debug.Assert a.Exists(10)
    Debug.Assert a.Exists(2)
    Debug.Assert a.Exists(0)
    Debug.Assert a.Exists(5)
    Debug.Assert Not a.Exists("10")
    Debug.Assert Not a.Exists("2")
    Debug.Assert Not a.Exists("0")
    Debug.Assert Not a.Exists("5")
    a.CompareMode = vbTextCompare
    Debug.Assert a.Exists("10")
    Debug.Assert a.Exists("2")
    Debug.Assert a.Exists("0")
    Debug.Assert a.Exists("5")
    
    Dim Comp As CountComparer
    Dim b As New XArray
    Dim Collection1 As New Collection
    Dim Collection2 As New Collection
    Dim Collection3 As New Collection
    Dim SearchCollection1 As New Collection
    Dim SearchCollection2 As New Collection
    Dim SearchCollection3 As New Collection
    Dim NonMatchCollection As New Collection
    
    Collection1.Add "A"
    Collection1.Add "B"
    Call Collection2.Count
    Collection3.Add "A"
    b.Add Collection1
    b.Add Collection2
    b.Add Collection3
    
    SearchCollection1.Add "X"
    SearchCollection1.Add "Y"
    Debug.Assert b.Exists(SearchCollection1, New CountComparer)

    Call SearchCollection2.Count
    Debug.Assert b.Exists(SearchCollection2, New CountComparer)

    SearchCollection3.Add "Z"
    Debug.Assert b.Exists(SearchCollection3, New CountComparer)

    NonMatchCollection.Add "1"
    NonMatchCollection.Add "2"
    NonMatchCollection.Add "3"
    NonMatchCollection.Add "4"
    Debug.Assert Not b.Exists(NonMatchCollection, New CountComparer)

    Dim StartTime As Currency
    Dim EndTime As Currency
    Dim c As New XArray
    Dim i As Long
    Dim Count As Long
    
    Count = 100000
    For i = 0 To Count - 1
        c.Add i
    Next
    
    StartTime = Timer()
    Debug.Assert Not c.Exists("Not Exists")
    EndTime = Timer() - StartTime
    
    Debug.Print EndTime
    Debug.Assert EndTime < 10
End Sub

Sub TestIndexOf()
    Dim a As New XArray
    a.Add 10
    a.Add 2
    a.Add 0
    a.Add 2
    a.Add 5
    
    a.CompareMode = vbBinaryCompare
    Debug.Assert a.IndexOf(10) = 0
    Debug.Assert a.IndexOf(2) = 1
    Debug.Assert a.IndexOf(0) = 2
    Debug.Assert a.IndexOf(5) = 4
    Debug.Assert a.IndexOf("10") = -1
    Debug.Assert a.IndexOf("2") = -1
    Debug.Assert a.IndexOf("0") = -1
    Debug.Assert a.IndexOf("5") = -1
    
    a.CompareMode = vbTextCompare
    Debug.Assert a.IndexOf("10") = 0
    Debug.Assert a.IndexOf("2") = 1
    Debug.Assert a.IndexOf("0") = 2
    Debug.Assert a.IndexOf("5") = 4
    
    Dim Comp As CountComparer
    Dim b As New XArray
    Dim Collection1 As New Collection
    Dim Collection2 As New Collection
    Dim Collection3 As New Collection
    Dim SearchCollection1 As New Collection
    Dim SearchCollection2 As New Collection
    Dim SearchCollection3 As New Collection
    Dim NonMatchCollection As New Collection
    
    Collection1.Add "A"
    Collection1.Add "B"
    Call Collection2.Count
    Collection3.Add "A"
    b.Add Collection1
    b.Add Collection2
    b.Add Collection3
    
    SearchCollection1.Add "X"
    SearchCollection1.Add "Y"
    Debug.Assert b.IndexOf(SearchCollection1, New CountComparer) = 0

    Call SearchCollection2.Count
    Debug.Assert b.IndexOf(SearchCollection2, New CountComparer) = 1

    SearchCollection3.Add "Z"
    Debug.Assert b.IndexOf(SearchCollection3, New CountComparer) = 2

    NonMatchCollection.Add "1"
    NonMatchCollection.Add "2"
    NonMatchCollection.Add "3"
    NonMatchCollection.Add "4"
    Debug.Assert b.IndexOf(NonMatchCollection, New CountComparer) = -1

    Dim StartTime As Currency
    Dim EndTime As Currency
    Dim c As New XArray
    Dim i As Long
    Dim Count As Long
    
    Count = 100000
    For i = 0 To Count - 1
        c.Add i
    Next
    
    StartTime = Timer()
    Debug.Assert c.IndexOf("Not Exists") = -1
    EndTime = Timer() - StartTime
    
    Debug.Print EndTime
    Debug.Assert EndTime < 10
End Sub

Sub TestExchange()
    Dim a As New XArray
    
    a.Add "A"
    a.Add "B"
    a.Add "C"

    a.Exchange 0, 1
    Debug.Assert a.Count = 3
    Debug.Assert a.Item(0) = "B"
    Debug.Assert a.Item(1) = "A"
    Debug.Assert a.Item(2) = "C"
    
    a.Exchange 0, 2
    Debug.Assert a.Item(0) = "C"
    Debug.Assert a.Item(1) = "A"
    Debug.Assert a.Item(2) = "B"
    
    a.Exchange 1, 2
    Debug.Assert a.Item(0) = "C"
    Debug.Assert a.Item(1) = "B"
    Debug.Assert a.Item(2) = "A"
    
    a.Exchange 2, 0
    Debug.Assert a.Item(0) = "A"
    Debug.Assert a.Item(1) = "B"
    Debug.Assert a.Item(2) = "C"
    
    Dim b As New XArray
    b.Add "Z"
    b.Exchange 0, 0
    Debug.Assert b.Count = 1
    Debug.Assert b.Item(0) = "Z"
    
    Dim StartTime As Currency
    Dim EndTime As Currency
    Dim c As New XArray
    Dim i As Long
    Dim Count As Long
    
    Count = 100000
    For i = 0 To Count - 1
        c.Add i
    Next
    
    StartTime = Timer()
    c.Exchange 0, Count - 1
    EndTime = Timer() - StartTime
    Debug.Print EndTime
    Debug.Assert EndTime < 10
    Debug.Assert c.Item(0) = Count - 1
    Debug.Assert c.Item(Count - 1) = 0
End Sub

Sub TestSort()
    Dim a As New XArray
    
    a.Add 10
    a.Add 12
    a.Add 0
    a.Add 1
    a.Add 5
    a.Sort
    Debug.Assert a.Item(0) = 0
    Debug.Assert a.Item(1) = 1
    Debug.Assert a.Item(2) = 5
    Debug.Assert a.Item(3) = 10
    Debug.Assert a.Item(4) = 12
    
    a.CompareMode = vbTextCompare
    a.Sort
    Debug.Assert a.Item(0) = 0
    Debug.Assert a.Item(1) = 1
    Debug.Assert a.Item(2) = 10
    Debug.Assert a.Item(3) = 12
    Debug.Assert a.Item(4) = 5
    
    a.CompareMode = vbBinaryCompare
    a.Sort
    Debug.Assert a.Item(0) = 0
    Debug.Assert a.Item(1) = 1
    Debug.Assert a.Item(2) = 5
    Debug.Assert a.Item(3) = 10
    Debug.Assert a.Item(4) = 12
    
    a.CompareMode = vbTextCompare
    a.Sort
    Debug.Assert a.Item(0) = 0
    Debug.Assert a.Item(1) = 1
    Debug.Assert a.Item(2) = 10
    Debug.Assert a.Item(3) = 12
    Debug.Assert a.Item(4) = 5
    a.CompareMode = vbBinaryCompare
    a.Sort New ValueComparer
    Debug.Assert a.Item(0) = 0
    Debug.Assert a.Item(1) = 1
    Debug.Assert a.Item(2) = 5
    Debug.Assert a.Item(3) = 10
    Debug.Assert a.Item(4) = 12
    
    Dim b As New XArray
    Dim Collection1 As New Collection
    Dim Collection2 As New Collection
    Dim Collection3 As New Collection

    b.Add Nothing
    b.Sort New CountComparer
    
    Collection1.Add "A"
    Collection1.Add "B"
    Call Collection2.Count
    Collection3.Add "A"
    b.Add Collection1
    b.Add Collection2
    b.Add Collection3
    b.Add Nothing
    b.Sort New CountComparer
    Debug.Assert b.Item(0) Is Nothing
    Debug.Assert b.Item(1) Is Nothing
    Debug.Assert b.Item(2) Is Collection2
    Debug.Assert b.Item(3) Is Collection3
    Debug.Assert b.Item(4) Is Collection1
    
    Dim c As New XArray
    c.Sort
    Debug.Assert c.Count = 0
    
    Dim d As New XArray
    Dim i As Long
    Dim Count As Long
    Dim StartTime As Currency
    Dim EndTime As Currency
    
    Count = 10000
    For i = Count - 1 To 0 Step -1
        d.Add i
    Next
    StartTime = Timer()
    d.Sort
    EndTime = Timer() - StartTime
    Debug.Print EndTime
    Debug.Assert EndTime < 10
End Sub

Sub TestReverse()
    Dim a As New XArray
    
    a.Add "A"
    a.Add "B"
    a.Add "C"
    a.Add "D"
    a.Add "E"
    a.Reverse
    
    Debug.Assert a.Item(0) = "E"
    Debug.Assert a.Item(1) = "D"
    Debug.Assert a.Item(2) = "C"
    Debug.Assert a.Item(3) = "B"
    Debug.Assert a.Item(4) = "A"
    
    Dim b As New XArray
    b.Reverse
    Debug.Assert b.Count = 0
End Sub

Sub TestClone()
    Dim a As New XArray
    Dim b As XArray
    
    a.Add "A"
    a.Add "B"
    a.Add "C"
    a.Add "D"
    a.Add "E"
    Debug.Assert a.Item(0) = "A"
    Debug.Assert a.Item(1) = "B"
    Debug.Assert a.Item(2) = "C"
    Debug.Assert a.Item(3) = "D"
    Debug.Assert a.Item(4) = "E"
    
    Set b = a.Clone
    
    Debug.Assert b.Item(0) = "A"
    Debug.Assert b.Item(1) = "B"
    Debug.Assert b.Item(2) = "C"
    Debug.Assert b.Item(3) = "D"
    Debug.Assert b.Item(4) = "E"
    
    Debug.Assert Not a Is b

    Dim StartTime As Currency
    Dim EndTime As Currency
    Dim c As New XArray
    Dim d As XArray
    Dim i As Long
    Dim Count As Long
    
    Count = 100000
    For i = 0 To Count - 1
        c.Add i
    Next
    
    StartTime = Timer()
    Set d = c.Clone
    EndTime = Timer() - StartTime
    Debug.Print EndTime
    Debug.Assert EndTime < 10
    Debug.Assert c.Count = Count
End Sub

Sub TestItems()
    Dim a As New XArray
    Dim Item
    Dim Count As Long
    a.Add "A"
    a.Add "B"
    a.Add "C"
    Count = 0
    For Each Item In a.Items
        Count = Count + 1
        Debug.Assert Item = a.Item(Count - 1)
    Next
    Debug.Assert Count = a.Count
    
    Dim b As New XArray
    Dim i As Long
    Dim Values
    Dim StartTime As Currency
    Dim EndTime As Currency
    Count = 10000
    For i = 0 To Count - 1
        b.Add i
    Next
    StartTime = Timer()
    Values = b.Items
    EndTime = Timer() - StartTime
    Debug.Print EndTime
    Debug.Assert EndTime < 10
    Debug.Assert IsArray(Values)
    Debug.Assert UBound(Values) = Count - 1
End Sub
