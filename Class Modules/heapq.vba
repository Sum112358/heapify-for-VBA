Option ClassModule
'Class Module: HeapqNum
Private Heap() As Variant
Private last As Integer

Public Sub heapInitalize()
    'Initalize the heap as 1,5 Array
    ReDim Heap(0 To 1)
    Dim zero(4) As Integer
    zero(0) = 0
    zero(1) = 0
    zero(2) = 0
    zero(3) = 0
    zero(4) = 0
    Heap(0) = zero 'Set the first value of the heap to be zero
    last = 0
End Sub

Private Sub checkSize()
    'Check the size of the heap and make it larger if needed
    currentSize = checkSizeArray(Heap)
    If currentSize <= last + 1 Then
        ReDim Preserve Heap(0 To currentSize * 2)
    End If
    
End Sub

Public Sub insert(X)
    'Add an element to the heap
    checkSize
    last = last + 1    
    Heap(last) = X
    swim
End Sub

Public Sub swim()
    'Find the appropriate location for the new element in the heap to maintain heapify
    'O(lg N) Operation
    current = last
    parentE = Floor(last / 2)
    Do While parentE > 0 And customCompare(Heap(current), Heap(parentE))
        exchangeKey (current), (parentE)
        current = parentE
        parentE = Floor(parentE / 2)
    Loop
End Sub

Public Function pop() As Variant
    'Extract the smallest value from the heap and move the last element to the top
    'Let the last element sink downwards
    If checkSizeArray(Heap) >= 1 Then
       Dim answer() As Variant
       ReDim answer(0 To 1)
       answer(0) = Heap(1)
       exchangeKey 1, (last)
       last = last - 1
       sink
       pop = answer(0)
    End If
End Function

Public Sub sink()
    'Maintain the heap property by "sinking" the top value through the heap
    'O(lg N) Operation
    current = 1
    answer = True
    Do While answer = True And (current * 2) <= last
        childL = current * 2
        childR = current * 2 + 1
        nxt = childL
        If childR <= last Then
            If customCompare(Heap(childR), Heap(current)) Then
                nxt = childR
            End If
        End If
        If customCompare(Heap(nxt), Heap(current)) Then
            exchangeKey (current), (nxt)
            current = nxt
        Else
            answer = False
        End If
    Loop
End Sub

Private Function customCompare(a, b) As Boolean
    'Custom Compare Operator to Check two arrays in the format (D,W,I,B,P)
    'Returns True if a < b
    'Minimize D then minimize W if not possible return False
    'Examples:
    'customCompare((1/3,3,1,0,"A1"),(1/2,2,1,17,"B7")) --> True
    'customCompare((1/3,3,1,0,"A1"),(1/3,2,1,17,"B7")) --> False
    If a(0) < b(0) Then
        customCompare = True
    ElseIf a(0) = b(0) Then
        If a(1) < b(1) Then
            customCompare = True
        Else
            customCompare = False
        End If
    Else
        customCompare = False
    End If
End Function

Private Sub exchangeKey(ByVal key1 As Integer, ByVal key2 As Integer)
    'Switch two objects within the Heap
    Value1 = Heap(key1)
    Value2 = Heap(key2)
    Heap(key1) = Value2
    Heap(key2) = Value1
End Sub

Public Function MaxValue() As Variant
    'Get the MaxValue of the heap
    MaxValue = Heap(1)
End Function

Public Function getHeap() As Variant
    'Return the Entire Heap
    getHeap = Heap
End Function

Public Function checkSizeArray(X)
    'Check the size of an Array and return the size
    checkSizeArray = UBound(X) - LBound(X) + 1
End Function

Public Function Floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    'Returns the floor of the value passed to the function
    'X is the value you want to round
    'is the multiple to which you want to round
    Floor = Int(X / Factor) * Factor
End Function








