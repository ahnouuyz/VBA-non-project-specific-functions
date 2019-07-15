' ============================================================================
' Core Functions
' ============================================================================

Function getTable(Optional dropBlankRows = True) As Variant()
    Dim row As Long
    Dim Arr1D() As Variant
    Dim Arr2D() As Variant
    Dim ArrArr() As Variant
    
    Arr2D = ActiveSheet.UsedRange.value2
    
    If dropBlankRows Then
        hasFirst = False
        For row = LBound(Arr2D) To UBound(Arr2D)
            Arr1D = getRow(Arr2D, row)
            If Not arrIsEmpty(Arr1D) Then
                appendVal ArrArr, Arr1D
            End If
        Next row
        Arr2D = ArrArrToArr2D(ArrArr)
    End If
    
    getTable = Arr2D
End Function

Function appendVal(Arr(), val, Optional base = 1)
    If (Not Arr) <> -1 Then
        ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
    Else
        ReDim Arr(base To base)
    End If
    
    Arr(UBound(Arr)) = val
End Function

Function getRow(Arr2D(), row)
    Dim col As Long
    ReDim Arr1D(LBound(Arr2D, 2) To UBound(Arr2D, 2))
    
    For col = LBound(Arr1D) To UBound(Arr1D)
        Arr1D(col) = Arr2D(row, col)
    Next col
    
    getRow = Arr1D
End Function

Function arrIsEmpty(Arr1D()) As Boolean
    Dim V2 As Variant
    
    For Each V2 In Arr1D
        If LenB(V2) Then
            arrIsEmpty = False
            Exit Function
        End If
    Next V2
    arrIsEmpty = True
End Function



Function ArrArrToArr2D(ArrArr()) As Variant()
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



Function numberToLetters(ByVal num) As String
    Dim remainder As Long
    Dim letters As String
    
    While num > 0
        remainder = ((num - 1) Mod 26)
        letters = Chr(remainder + Asc("A")) & letters
        num = (num - remainder) \ 26
    Wend
    
    numberToLetters = letters
End Function

Function lettersToNumber(ByVal letters) As Long
    Dim power As Long
    Dim number As Long
    
    letters = UCase(letters)
    For power = 0 To Len(letters) - 1
        number = number + (Asc(Right(letters, 1)) - Asc("A") + 1) * (26 ^ power)
        letters = Left(letters, Len(letters) - 1)
    Next power
    
    lettersToNumber = number
End Function
