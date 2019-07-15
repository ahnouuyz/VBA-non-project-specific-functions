' ============================================================================
' Core Functions
' ============================================================================

Function getTable(Optional dropBlankRows = True) As Variant()
    Dim hasFirst As Boolean
    Dim row As Long
    Dim Arr2D() As Variant
    Dim ArrArr() As Variant
    
    Arr2D = ActiveSheet.UsedRange.value2
    
    If dropBlankRows Then
        hasFirst = False
        For row = LBound(Arr2D) To UBound(Arr2D)
            If rowNotBlank(Arr2D, row) Then
                appendRow ArrArr, Arr2D, row, hasFirst
            End If
        Next row
        Arr2D = ArrArrToArr2D(ArrArr)
    End If
    
    getTable = Arr2D
End Function

Function appendRow(ArrArr, Arr2D, row, hasFirst)
    Dim col As Long
    ReDim Arr1D(LBound(Arr2D, 2) To UBound(Arr2D, 2))
    
    For col = LBound(Arr1D) To UBound(Arr1D)
        Arr1D(col) = Arr2D(row, col)
    Next col
    
    If hasFirst Then
        ReDim Preserve ArrArr(1 To UBound(ArrArr) + 1)
    Else
        ReDim ArrArr(1 To 1)
        hasFirst = True
    End If
    
    ArrArr(UBound(ArrArr)) = Arr1D
End Function

Function rowNotBlank(Arr2D, row) As Boolean
    Dim col As Long
    
    For col = LBound(Arr2D, 2) To UBound(Arr2D, 2)
        If LenB(Arr2D(row, col)) Then
            rowNotBlank = True
            Exit Function
        End If
    Next col
    rowNotBlank = False
End Function

Function ArrArrToArr2D(ArrArr) As Variant()
    Dim row As Long
    Dim col As Long
    Dim c1 As Long
    Dim cN As Long
    
    c1 = 2147483647
    cN = 0
    For row = LBound(ArrArr) To UBound(ArrArr)
        If Not IsEmpty(ArrArr(row)) Then
            If c1 > LBound(ArrArr(row)) Then c1 = LBound(ArrArr(row))
            If cN < UBound(ArrArr(row)) Then cN = UBound(ArrArr(row))
        End If
    Next row
    
    ReDim Arr2D(LBound(ArrArr) To UBound(ArrArr), c1 To cN) As Variant
    
    For row = LBound(ArrArr) To UBound(ArrArr)
        If Not IsEmpty(ArrArr(row)) Then
            For col = LBound(ArrArr(row)) To UBound(ArrArr(row))
                Arr2D(row, col) = ArrArr(row)(col)
            Next col
        End If
    Next row
    
    ArrArrToArr2D = Arr2D
End Function

Function PrintArr2D(Arr2D, Optional r1 = 1, Optional c1 = 1)
    Dim nrows As Long
    Dim ncols As Long
    
    nrows = UBound(Arr2D) - LBound(Arr2D) + 1
    ncols = UBound(Arr2D, 2) - LBound(Arr2D, 2) + 1
    Cells(r1, c1).Resize(nrows, ncols).value2 = Arr2D
End Function



Function createButton(PositionRng, ButtonText, Optional OnAction = "doNothing")
    With PositionRng
        ActiveSheet.Buttons.Add(.Left, .Top, .Width, .Height).Text = ButtonText
    End With
    
    With ActiveSheet
        With .Shapes(.Shapes.count)
            .Placement = xlMoveAndSize
            .OnAction = OnAction
        End With
    End With
End Function

Function doNothing()
    MsgBox "Nothing was done.", , "Do Nothing"
    Debug.Print "Nothing was done."
End Function

Private Sub deleteButtons()
    ActiveSheet.Buttons.Delete
End Sub
