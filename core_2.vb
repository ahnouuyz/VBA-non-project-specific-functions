' ============================================================================
' Core Functions
' ============================================================================

Function inList(value, list) As Boolean
' Check if the given value exists in the given list.
'
    Dim V2
    
    For Each V2 In list
        If value = V2 Then
            inList = True
            Exit Function
        End If
    Next V2
End Function



Function swapValues(value1, value2)
    Dim V2
    
    V2 = value1
    value1 = value2
    value2 = V2
End Function

Function quickSortArr1D(Arr1D, Optional r1 = -1, Optional rN = -1, Optional cKey = -1)
' Sort a 1D array in ascending order.
'
    Dim tr1 As Long
    Dim trN As Long
    Dim pivotValue
    
    If Not IsArray(Arr1D) Then
        MsgBox "Not an array, cannot be sorted."
        Exit Function
    End If
    
    If r1 = -1 Then r1 = LBound(Arr1D)
    If rN = -1 Then rN = UBound(Arr1D)
    tr1 = r1
    trN = rN
    
    If cKey <> -1 Then
        pivotValue = Arr1D((r1 + rN) \ 2)(cKey)
        
        Do While tr1 <= trN
            Do While Arr1D(tr1)(cKey) < pivotValue And tr1 < rN
                tr1 = tr1 + 1
            Loop
            
            Do While Arr1D(trN)(cKey) > pivotValue And trN > r1
                trN = trN - 1
            Loop
            
            If tr1 <= trN Then
                swapValues Arr1D(tr1), Arr1D(trN)
                tr1 = tr1 + 1
                trN = trN - 1
            End If
        Loop
    Else
        pivotValue = Arr1D((r1 + rN) \ 2)
        
        Do While tr1 <= trN
            Do While Arr1D(tr1) < pivotValue And tr1 < rN
                tr1 = tr1 + 1
            Loop
            
            Do While Arr1D(trN) > pivotValue And trN > r1
                trN = trN - 1
            Loop
            
            If tr1 <= trN Then
                swapValues Arr1D(tr1), Arr1D(trN)
                tr1 = tr1 + 1
                trN = trN - 1
            End If
        Loop
    End If
    
    If r1 < trN Then quickSortArr1D Arr1D, r1, trN, cKey
    If tr1 < rN Then quickSortArr1D Arr1D, tr1, rN, cKey
End Function

Function reverseArr1D(Arr1D, Optional r1 = -1, Optional rN = -1)
    Dim L2 As Long
    Dim nSwaps As Long
    
    If r1 = -1 Then r1 = LBound(Arr1D)
    If rN = -1 Then rN = UBound(Arr1D)
    
    nSwaps = (rN - r1 + 1) \ 2
    
    For L2 = 0 To nSwaps
        swapValues Arr1D(r1 + L2), Arr1D(rN - L2)
    Next L2
End Function

' ===============================================================================
' Not frequently used yet.
' ===============================================================================



Function Arr2DToArrArr(Arr2D, Optional byCols = False) As Variant()
    Dim L2 As Long
    Dim L3 As Long
    Dim Dim1 As Long
    Dim Dim2 As Long
    Dim Arr1D() As Variant
    Dim ArrArr() As Variant
    
    Dim1 = 1 - (byCols = True)
    Dim2 = 1 - (Not byCols = True)
    
    ReDim ArrArr(LBound(Arr2D, Dim1) To UBound(Arr2D, Dim1))
    ReDim Arr1D(LBound(Arr2D, Dim2) To UBound(Arr2D, Dim2))
    
    For L2 = LBound(Arr2D, Dim1) To UBound(Arr2D, Dim1)
        For L3 = LBound(Arr2D, Dim2) To UBound(Arr2D, Dim2)
            If Not byCols Then
                Arr1D(L3) = Arr2D(L2, L3)
            ElseIf byCols Then
                Arr1D(L3) = Arr2D(L3, L2)
            End If
        Next L3
        
        ArrArr(L2) = Arr1D
    Next L2
    
    Arr2DToArrArr = ArrArr
End Function

Function ReindexArr2D(Arr2D, Optional r1 = 0, Optional c1 = 0) As Variant()
' Reassign the base indices (default = 0) for both dimensions of a 2D array.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim rN As Long
    Dim cN As Long
    
    rN = UBound(Arr2D) - LBound(Arr2D) + 1 ' Number of rows
    cN = UBound(Arr2D, 2) - LBound(Arr2D, 2) + 1 ' Number of columns
    
    ReDim NewArr2D(r1 To rN + r1 - 1, c1 To cN + c1 - 1) As Variant
    
    For L2 = r1 To rN + r1 - 1
        For L3 = c1 To cN = c1 - 1
            NewArr2D(L2, L3) = Arr2D(LBound(Arr2D) + L2 - r1, LBound(Arr2D, 2) + L3 - c1)
        Next L3
    Next L2
    
    ReindexArr2D = NewArr2D
End Function
