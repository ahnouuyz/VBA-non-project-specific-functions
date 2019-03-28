' ================================================================================================
' Linear Algebra Algorithms
'
' Purpose:
'     Provide a convenient platform for matrix operations.
'
' Requirements:
'     Microsoft® Excel® 2016 MSO (16.0.9029.2106) 32-bit (Microsoft Corp., Redmond, WA, USA).
'
' Usage:
'     Enter matrices in a spreadsheet, and click buttons to perform the desired operations.
'     Hopefully it's as simple as that.
'
' Limitations:
'     Accuracy.
'     Not all functions available.
'
' Version:
'     2018/08/01
'
' Author:
'     ©2018 Zhuoyuan (Roscoe) Lai, ahnouuyz@gmail.com
' ================================================================================================
' Available functions:
'     Create identity matrix
'     Transpose matrix
'     Calculate trace of matrix (if square)
'     Calculate determinant of matrix (if square)
'     Invert matrix (if invertible)
'     Matrix multiplication (if allowed)
'     Matrix power
'     Eigendecomposition of a symmetric matrix
'     One matrix element-wise operation with a scalar (add, mult, div, power)
'     (Reduced) row echelon form of matrix
' Unused:
'     Calculate cofactors
'     Calculate adjugate matrix
'     Kronecker product
' Not available:
'     Two matrix element-wise operations
'     Concatenation (not so important, can be performed on spreadsheet)
' ================================================================================================

' To Do:
'     Complex calculator
'     0 ^ (-ve)
'     (-ve) ^ 0.2
'     Bland-Altman plot
'     outRow and outCol system, and matrixPrint

' Limit scanning to the 1st 1000 diagonal lines.
Private Const firstDiag As Long = 1
Private Const lastDiag As Long = 1000 ' Can go up to 16384 and even 1048567 (please don't try)

' Positioning of output matrices.
Private Const outRow1 As Long = 3 ' 1st row for output
Private Const outCol1 As Long = 1 ' 1st column for output
Private Const outCol2 As Long = outCol1 + 1 ' 2nd column for output (top-left cell of matrix)

Private Const roundValues As Boolean = False
Private Const roundDecimalPlaces As Long = 11
Private deleteSheetUnlock As Boolean

Private Sub symMatEigenDecom()
' Calculate the eigenvalues and eigenvectors of a symmetric matrix.
'
    Dim L2 As Long
    Dim outRow2 As Long
    Dim outRow3 As Long
    Dim outRow4 As Long
    Dim outCol2 As Long
    Dim outCol3 As Long
    Dim inputRows As Long
    Dim inputArr As Variant
    Dim inputEigen As Variant
    
    ActiveSheet.Copy , ActiveSheet
    inputArr = arrScanTopLeft
    
    If matrixIsInvalid(inputArr) Then Exit Sub
    
    If matrixIsSymmetric(inputArr) Then
        inputEigen = callEigenJK(inputArr)
    Else
        MsgBox "Not a symmetric matrix. Exiting script."
        deleteSheet True
        Exit Sub
    End If
    
    inputRows = UBound(inputArr) - LBound(inputArr) + 1
    
    ' Delete everything else from the sheet.
    ActiveSheet.UsedRange.Clear
    ActiveSheet.Cells.Borders.LineStyle = xlNone
    
    outRow2 = outRow1 + inputRows + 2
    outRow3 = outRow2 + inputRows + 1
    outRow4 = outRow3 + inputRows + 2
    
    outCol2 = outCol1 + 2
    outCol3 = outCol2 + inputRows ' inputCols = inputRows
    
    For L2 = 1 To inputRows
        Cells(outRow1 - 1, outCol2 - 1 + L2).Value2 = "eVect " & L2
        Cells(outRow2 - 1, outCol2 - 1 + L2).Value2 = "eVal " & L2
        Cells(outRow3 - 1 + L2, outCol2 - 1).Value2 = "eVect " & L2
    Next L2
    
    matrixPrint inputEigen(1), "T", outRow1, outCol1, outCol2, outCol3
    matrixPrint inputEigen(2), "D", outRow2, outCol1, outCol2, outCol3
    matrixPrint inputEigen(3), "trans(T)", outRow3, outCol1, outCol2, outCol3
    matrixPrint inputArr, "A", outRow4, outCol1, outCol2, outCol3
    
    Rows(outRow4 - 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Function numToLetter(num As Long) As String
'
'
    Dim tempNum As Long
    Dim letterIndex As Long
    Dim letterString As String
    
    tempNum = num
    Do
        letterIndex = ((tempNum - 1) Mod 26)
        letterString = Chr(letterIndex + 65) & letterString
        tempNum = (tempNum - letterIndex) \ 26
    Loop While tempNum > 0
    
    numToLetter = letterString
End Function

Private Sub nMatrixProduct(Optional numOfMatrices As Long = 2)
' Enter desired number of matrices in a spreadsheet.
' Pay attention to the order, place matrices to the left closer to the top-left corner.
' Return the product if it exists.
'
    Dim L2 As Long
    Dim outCol3 As Long
    Dim maxColArr As Long
    Dim prodStr As String
    Dim arrNameStr As String
    Dim arrNames() As String
    
    ReDim arrOfArrs(1 To numOfMatrices + 1) As Variant
    ReDim rowsArr(1 To numOfMatrices + 1) As Long
    ReDim colsArr(1 To numOfMatrices + 1) As Long
    ReDim outRowArr(1 To numOfMatrices + 1) As Long
    
    ActiveSheet.Copy , ActiveSheet
    
    ' Reserve 1st index for the product.
    For L2 = 2 To numOfMatrices + 1
        arrOfArrs(L2) = arrScanTopLeft
        
        If matrixIsInvalid(arrOfArrs(L2), L2 - 1) Then Exit Sub
        
        ' Measure the dimensions of the matrix.
        rowsArr(L2) = UBound(arrOfArrs(L2)) - LBound(arrOfArrs(L2)) + 1
        colsArr(L2) = UBound(arrOfArrs(L2), 2) - LBound(arrOfArrs(L2), 2) + 1
        If maxColArr < colsArr(L2) Then maxColArr = colsArr(L2)
        
        ' Perform the matrix multiplication.
        If L2 = 2 Then
            arrOfArrs(1) = arrOfArrs(2)
        Else
            If IsArray(arrOfArrs(1)) Then
                arrOfArrs(1) = matrixMultiply(arrOfArrs(1), arrOfArrs(L2))
            End If
        End If
        
        arrNameStr = arrNameStr & "," & numToLetter(L2 - 1)
        prodStr = prodStr & numToLetter(L2 - 1)
    Next L2
    
    arrNameStr = "," & prodStr & arrNameStr
    arrNames = Split(arrNameStr, ",")
    
    ' Delete everything else from the sheet.
    ActiveSheet.UsedRange.Clear
    ActiveSheet.Cells.Borders.LineStyle = xlNone
    
    ' Determine output columns and rows, and print output.
    outRowArr(1) = outRow1
    outCol3 = outCol2 + maxColArr
    
    If matrixFinalProductExists(colsArr, rowsArr, numOfMatrices) Then
        rowsArr(1) = UBound(arrOfArrs(1)) - LBound(arrOfArrs(1)) + 1
        matrixPrint arrOfArrs(1), arrNames(1), outRowArr(1), outCol1, outCol2, outCol3
    Else
        rowsArr(1) = 1
        Cells(outRowArr(1), outCol1).Value2 = arrNames(1)
        Cells(outRowArr(1), outCol2).Value2 = arrOfArrs(1)
    End If
    
    For L2 = 2 To numOfMatrices + 1
        outRowArr(L2) = outRowArr(L2 - 1) + rowsArr(L2 - 1) + 1 - (L2 = 2)
        matrixPrint arrOfArrs(L2), arrNames(L2), outRowArr(L2), outCol1, outCol2, outCol3
    Next L2
    
    Rows(outRowArr(2) - 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Sub oneMatrixTrans()
' Transpose the first detectable matrix on the spreadsheet.
'
    Dim outRow2 As Long
    Dim outCol3 As Long
    Dim inputRows As Long
    Dim inputCols As Long
    Dim inputArr As Variant
    Dim inputTrans() As Double
    
    ActiveSheet.Copy , ActiveSheet
    inputArr = arrScanTopLeft
    
    If matrixIsInvalid(inputArr) Then Exit Sub
    
    inputRows = UBound(inputArr) - LBound(inputArr) + 1
    inputCols = UBound(inputArr, 2) - LBound(inputArr, 2) + 1
    
    inputTrans = matrixTranspose(inputArr)
    
    ' Delete everything else from the sheet.
    ActiveSheet.UsedRange.Clear
    ActiveSheet.Cells.Borders.LineStyle = xlNone
    
    outRow2 = outRow1 + inputCols + 2 ' transRows = inputCols
    
    If inputCols < inputRows Then inputCols = inputRows
    outCol3 = outCol2 + inputCols
    
    matrixPrint inputTrans, "trans(A)", outRow1, outCol1, outCol2, outCol3
    matrixPrint inputArr, "A", outRow2, outCol1, outCol2, outCol3
    
    Rows(outRow2 - 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Sub oneMatrixInv()
' Invert the first detectable matrix on the spreadsheet, if invertible.
'
    Dim outRow2 As Long
    Dim outCol3 As Long
    Dim inputRows As Long
    Dim inputCols As Long
    Dim inputArr As Variant
    Dim inputInvDet As Variant
    
    ActiveSheet.Copy , ActiveSheet
    inputArr = arrScanTopLeft
    
    If matrixIsInvalid(inputArr) Then Exit Sub
    
    inputRows = UBound(inputArr) - LBound(inputArr) + 1
    inputCols = UBound(inputArr, 2) - LBound(inputArr, 2) + 1
    
    ' Delete everything else from the sheet.
    ActiveSheet.UsedRange.Clear
    ActiveSheet.Cells.Borders.LineStyle = xlNone
    
    inputInvDet = matrixGauss(inputArr)
    
    outCol3 = outCol2 + inputCols
    
    If inputRows = inputCols And inputInvDet(4) <> 0 Then
        outRow2 = outRow1 + inputRows + 2
        matrixPrint inputInvDet(3), "inv(A)", outRow1, outCol1, outCol2, outCol3
    Else
        outRow2 = outRow1 + 1 + 2
        Cells(outRow1, outCol1).Value2 = "inv(A)"
        Cells(outRow1, outCol2).Value2 = inputInvDet(3)
    End If
    
    matrixPrint inputArr, "A", outRow2, outCol1, outCol2, outCol3
    
    Rows(outRow2 - 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Sub oneMatrixPower(Optional matPow As Long = 1)
' Raise the first detectable matrix on the spreadsheet to the desired power, if possible.
'
    Dim outRow2 As Long
    Dim outCol3 As Long
    Dim inputRows As Long
    Dim inputCols As Long
    Dim inputArr As Variant
    Dim inputPow As Variant
    
    ActiveSheet.Copy , ActiveSheet
    inputArr = arrScanTopLeft
    
    If matrixIsInvalid(inputArr) Then Exit Sub
    
    inputRows = UBound(inputArr) - LBound(inputArr) + 1
    inputCols = UBound(inputArr, 2) - LBound(inputArr, 2) + 1
    
    inputPow = matrixPower(inputArr, matPow)
    
    ' Delete everything else from the sheet.
    ActiveSheet.UsedRange.Clear
    ActiveSheet.Cells.Borders.LineStyle = xlNone
    
    outCol3 = outCol2 + inputCols ' Input has to be a square matrix
    
    If inputRows = inputCols And matPow > -1 Then
        ' Nonnegative powers.
        outRow2 = outRow1 + inputRows + 2
        matrixPrint inputPow, "A^" & matPow, outRow1, outCol1, outCol2, outCol3
    ElseIf inputRows = inputCols And matrixDeterminant(inputArr) <> 0 Then
        ' Negative powers, must be invertible.
        outRow2 = outRow1 + inputRows + 2
        matrixPrint inputPow, "A^" & matPow, outRow1, outCol1, outCol2, outCol3
    Else
        outRow2 = outRow1 + 1 + 2
        Cells(outRow1, outCol1).Value2 = "A^" & matPow
        Cells(outRow1, outCol2).Value2 = inputPow
    End If
    
    matrixPrint inputArr, "A", outRow2, outCol1, outCol2, outCol3
    
    Rows(outRow2 - 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Sub oneMatrixElementwise(opVal As Long, numVal As Double)
' Perform an operation with a number on each element in a matrix.
'
    Dim outRow2 As Long
    Dim outCol3 As Long
    Dim inputRows As Long
    Dim inputCols As Long
    Dim inputArr As Variant
    Dim inputOp() As Double
    Dim opStr() As String
    
    opStr = Split(",(el+,(el*,(el/,(el^", ",")
    
    ' Create a copy of the input, because the scanner function will delete arrays.
    ActiveSheet.Copy , ActiveSheet
    inputArr = arrScanTopLeft
    
    If matrixIsInvalid(inputArr) Then Exit Sub
    
    inputRows = UBound(inputArr) - LBound(inputArr) + 1
    inputCols = UBound(inputArr, 2) - LBound(inputArr, 2) + 1
    
    ' Delete everything else from the sheet.
    ActiveSheet.UsedRange.Clear
    ActiveSheet.Cells.Borders.LineStyle = xlNone
    
    inputOp = matrixElementWise(inputArr, opVal, numVal)
    
    outCol3 = outCol2 + inputCols
    outRow2 = outRow1 + inputRows + 2
    
    matrixPrint inputOp, "A" & opStr(opVal) & numVal & ")", outRow1, outCol1, outCol2, outCol3
    matrixPrint inputArr, "A", outRow2, outCol1, outCol2, outCol3
    
    Rows(outRow2 - 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Sub oneMatrix_R_REF()
' Perform an operation with a number on each element in a matrix.
'
    Dim outRow2 As Long
    Dim outRow3 As Long
    Dim outCol3 As Long
    Dim inputRows As Long
    Dim inputCols As Long
    Dim inputArr As Variant
    Dim inputR_REF As Variant
    
    ' Create a copy of the input, because the scanner function will delete arrays.
    ActiveSheet.Copy , ActiveSheet
    inputArr = arrScanTopLeft
    
    If matrixIsInvalid(inputArr) Then Exit Sub
    
    inputRows = UBound(inputArr) - LBound(inputArr) + 1
    inputCols = UBound(inputArr, 2) - LBound(inputArr, 2) + 1
    
    ' Delete everything else from the sheet.
    ActiveSheet.UsedRange.Clear
    ActiveSheet.Cells.Borders.LineStyle = xlNone
    
    inputR_REF = matrixGauss(inputArr)
    
    outCol3 = outCol2 + inputCols
    outRow2 = outRow1 + inputRows + 1
    outRow3 = outRow2 + inputRows + 2
    
    matrixPrint inputR_REF(1), "REF(A)", outRow1, outCol1, outCol2, outCol3
    matrixPrint inputR_REF(2), "RREF(A)", outRow2, outCol1, outCol2, outCol3
    matrixPrint inputArr, "A", outRow3, outCol1, outCol2, outCol3
    
    Rows(outRow3 - 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

' ================================================================================================
' Function Library
' ================================================================================================

Function arrScanRC() As Long()
' Scan the active spreadsheet.
' Find the number closest to the top-left corner of the spreadsheet that is not alone.
' Returns [row number, column number] for that cell.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim r1c1(1 To 2) As Long
    
    For L2 = firstDiag + 1 To lastDiag + 1
        For L3 = 1 To L2 - 1
            If L3 <= 16384 Then
                If LenB(Cells(L2 - L3, L3).Value2) <> 0 Then
                    If IsNumeric(Cells(L2 - L3, L3).Value2) Then
                        If Not isSingleNumber(L2 - L3, L3) Then
                            r1c1(1) = L2 - L3
                            r1c1(2) = L3
                            
                            If L3 > 1 Then
                                ' The one exception.
                                ' Reject if number found is the trace of the matrix.
                                If Cells(L2 - L3, L3 - 1).Value2 = "Trace" Then
                                    r1c1(1) = 0
                                    r1c1(2) = 0
                                Else
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    End If
                End If
            Else
                Exit For
            End If
        Next L3
        
        If r1c1(1) <> 0 And r1c1(2) <> 0 Then Exit For
    Next L2
    
    arrScanRC = r1c1
End Function

Function arrScanTopLeft() As Variant
' Find the array closest to the top-left corner of the spreadsheet.
' Store array in memory and delete from spreadsheet.
'
    Dim rN As Long
    Dim cN As Long
    Dim maxRow As Long
    Dim maxCol As Long
    Dim r1c1() As Long
    
    ' Find coordinates of first cell.
    r1c1 = arrScanRC
    
    If r1c1(1) = 0 Or r1c1(2) = 0 Then
        MsgBox "No array of numbers were found in the first " & lastDiag & " diagonals."
        Exit Function
    End If
    
    ' Find coordinates of last cell.
    rN = r1c1(1)
    cN = r1c1(2)
    
    ' Scan down 1st column, across rows, find longest row.
    Do While isNonEmptyNum(Cells(rN, r1c1(2)).Value2)
        Do While isNonEmptyNum(Cells(rN, cN + 1).Value2)
            cN = cN + 1
        Loop
        If maxCol < cN Then maxCol = cN
        cN = r1c1(2) ' Reset
        rN = rN + 1
    Loop
    
    ' Scan across 1st row, down columns, find longest column.
    rN = r1c1(1) ' Reset
    Do While isNonEmptyNum(Cells(r1c1(1), cN).Value2)
        Do While isNonEmptyNum(Cells(rN + 1, cN).Value2)
            rN = rN + 1
        Loop
        If maxRow < rN Then maxRow = rN
        rN = r1c1(1) ' Reset
        cN = cN + 1
    Loop
    
    arrScanTopLeft = Range(Cells(r1c1(1), r1c1(2)), Cells(maxRow, maxCol)).Value2
    Range(Cells(r1c1(1), r1c1(2)), Cells(maxRow, maxCol)).Clear ' Delete array from sheet
End Function

' ================================================================================================
' Checking Functions
' ================================================================================================

Function isSingleNumber(iR, jC) As Boolean
' Check if the cells to the right and below are empty or contain non-numeric values.
'
    Dim notNumRight As Boolean
    Dim notNumBottom As Boolean
    
    notNumRight = LenB(Cells(iR, jC + 1)) = 0 Or Not IsNumeric(Cells(iR, jC + 1))
    notNumBottom = LenB(Cells(iR + 1, jC)) = 0 Or Not IsNumeric(Cells(iR + 1, jC))
    
    If notNumRight And notNumBottom Then isSingleNumber = True
End Function

Function isNonEmptyNum(ByRef inputVal) As Boolean
    If LenB(inputVal) <> 0 And IsNumeric(inputVal) Then
        isNonEmptyNum = True
    End If
End Function

Function isInteger(ByRef inputVal, Optional posInt = False) As Boolean
    If IsNumeric(inputVal) Then
        If inputVal = CLng(inputVal) Then
            If posInt Then
                If inputVal > 0 Then isInteger = True
            Else
                isInteger = True
            End If
        End If
    End If
End Function

Function matrixIsInvalid(ByRef inputArr, Optional arrNum = -1) As Boolean
' Check if arrScanTopLeft picked up an invalid matrix.
'
    If arrNum = -1 Then
        If Not IsArray(inputArr) Then
            MsgBox "No arrays found. Exiting script."
            matrixIsInvalid = True
        ElseIf matrixHasNonNum(inputArr) Then
            MsgBox "Empty or non-numeric cell detected. Exiting script."
            matrixIsInvalid = True
        End If
    Else
        If Not IsArray(inputArr) Then
            MsgBox "Array " & arrNum & " not found. Exiting script."
            matrixIsInvalid = True
        ElseIf matrixHasNonNum(inputArr) Then
            MsgBox "Empty or non-numeric cell detected in Array " & arrNum & ". Exiting script."
            matrixIsInvalid = True
        End If
    End If
    
    If matrixIsInvalid = True Then deleteSheet True
End Function

Function matrixHasNonNum(ByRef inputArr) As Boolean
' Check for empty or non-numeric values in matrix.
'
    For Each arrVal In inputArr
        If Not isNonEmptyNum(arrVal) Then
            matrixHasNonNum = True
            Exit Function
        End If
    Next arrVal
    matrixHasNonNum = False
End Function

Function matrixIsSquare(ByRef inArr) As Boolean
    If UBound(inArr) - LBound(inArr) = UBound(inArr, 2) - LBound(inArr, 2) Then
        matrixIsSquare = True
    End If
End Function

Function matrixIsSymmetric(ByRef inputArr) As Boolean
    Dim L2 As Long
    Dim L3 As Long
    
    If matrixIsSquare(inputArr) Then
        For L2 = LBound(inputArr) To UBound(inputArr)
            For L3 = L2 + 1 To UBound(inputArr)
                If inputArr(L2, L3) <> inputArr(L3, L2) Then
                    matrixIsSymmetric = False
                    Exit Function
                End If
            Next L3
        Next L2
    Else
        matrixIsSymmetric = False
        Exit Function
    End If
    matrixIsSymmetric = True
End Function

Function matrixFinalProductExists(ByRef colsArr, ByRef rowsArr, numOfMatrices) As Boolean
' Check if matrix multiplication is possible for the given matrices.
' Not a generalizable function (just take a look at the input parameters).
'
    Dim L2 As Long
    Dim validCount As Long
    
    For L2 = 2 To numOfMatrices
        If colsArr(L2) = rowsArr(L2 + 1) Then
            validCount = validCount + 1
        End If
    Next L2
    
    If validCount = numOfMatrices - 1 Then
        matrixFinalProductExists = True
    Else
        matrixFinalProductExists = False
    End If
End Function



Function matrixPrint(ByRef inArr, matName, r1, c1, c2, outCol3)
' Print matrix onto spreadsheet.
' outCol3 is one column after the widest matrix.
'
    Dim rN As Long
    Dim cN As Long
    Dim numOfRows As Long
    Dim numOfCols As Long
    
    numOfRows = UBound(inArr) - LBound(inArr) + 1
    numOfCols = UBound(inArr, 2) - LBound(inArr, 2) + 1
    
    rN = r1 + numOfRows - 1
    cN = c2 + numOfCols - 1
    
    Cells(r1, c1).Value2 = matName
    
    With Range(Cells(r1, c2), Cells(rN, cN))
        If roundValues Then
            .Value2 = matrixRoundOff(inArr)
        Else
            .Value2 = inArr
        End If
        
        .NumberFormat = "# ??/??"
        .BorderAround xlContinuous
    End With
    
    ' Do not print trace and determinant for vectors or scalars.
    If numOfRows > 1 And numOfCols > 1 Then
        Cells(rN - 1, outCol3).Value2 = "Trace" ' Remember, used as exclusion criteria!
        Cells(rN, outCol3).Value2 = "Determ"
        
        If numOfRows = numOfCols Then
            Cells(rN - 1, outCol3 + 1).Value2 = matrixTrace(inArr)
            Cells(rN, outCol3 + 1).Value2 = matrixDeterminant(inArr)
        Else
            Cells(rN - 1, outCol3 + 1).Value2 = "Undefined"
            Cells(rN, outCol3 + 1).Value2 = "Undefined"
        End If
    End If
End Function

Function matrixRoundOff(ByRef inputArr)
' Round off values in a 2D matrix to set number of decimal places.
'
    Dim iRow As Long
    Dim iCol As Long
    
    For iRow = LBound(inputArr) To UBound(inputArr)
        For iCol = LBound(inputArr, 2) To UBound(inputArr, 2)
            inputArr(iRow, iCol) = Round(inputArr(iRow, iCol), roundDecimalPlaces)
        Next iCol
    Next iRow
    
    matrixRoundOff = inputArr
End Function

Function callEigenJK(ByRef inArr) As Variant
' Calls EIGEN_JK function.
' Recalculate eigenvalues with the correct sign, D = T(trans)AT.
' Effectively removes the positive definite constraint of the original function.
' Returns output as [T, D, T(trans)].
'
    Dim L2 As Long
    Dim L3 As Long
    Dim numOfRows As Long
    Dim jkOut As Variant
    Dim outArr(1 To 3) As Variant
    
    jkOut = EIGEN_JK(inArr)
    numOfRows = UBound(inArr) - LBound(inArr) + 1
    
    ReDim matrixT(1 To numOfRows, 1 To numOfRows) As Double
    ReDim matrixTt(1 To numOfRows, 1 To numOfRows) As Double
    ReDim matrixTtAT(1 To numOfRows, 1 To numOfRows) As Double
    ReDim matrixD(1 To numOfRows, 1 To numOfRows) As Double
    
    For L2 = 1 To numOfRows
        For L3 = 1 To numOfRows
            matrixT(L2, L3) = jkOut(L2, L3 + 1)
        Next L3
    Next L2
    
    matrixTt = matrixTranspose(matrixT)
    matrixTtAT = matrixMultiply(matrixMultiply(matrixTt, inArr), matrixT)
    
    ' TtAT is actually the diagonal matrix of eigenvalues.
    ' Redefine to remove off-diagonal errors.
    For L2 = 1 To numOfRows
        matrixD(L2, L2) = matrixTtAT(L2, L2)
    Next L2
    
    outArr(1) = matrixT
    outArr(2) = matrixD
    outArr(3) = matrixTt
    
    callEigenJK = outArr
End Function

Function EIGEN_JK(ByRef M As Variant) As Variant
' A Function That Computes the Eigenvalues and Eigenvectors For a Real Symmetric Matrix
' http://www.freevbcode.com/ShowCode.asp?ID=9209

'***************************************************************************
'**  Function computes the eigenvalues and eigenvectors for a real        **
'**  symmetric positive definite matrix using the "JK Method".  The       **
'**  first column of the return matrix contains the eigenvalues and       **
'**  the rest of the p+1 columns contain the eigenvectors.                **
'**  See:                                                                 **
'**  KAISER,H.F. (1972) "THE JK METHOD: A PROCEDURE FOR FINDING THE       **
'**  EIGENVALUES OF A REAL SYMMETRIC MATRIX", The Computer Journal,       **
'**  VOL.15, 271-273.                                                     **
'***************************************************************************

Dim A() As Variant, Ematrix() As Double
Dim i As Long, j As Long, k As Long, iter As Long, p As Long
Dim den As Double, hold As Double, Sin_ As Double, num As Double
Dim Sin2 As Double, Cos2 As Double, Cos_ As Double, Test As Double
Dim Tan2 As Double, Cot2 As Double, tmp As Double
Const eps As Double = 1E-16
    
    On Error GoTo EndProc
    
    A = M
    p = UBound(A, 1)
    ReDim Ematrix(1 To p, 1 To p + 1)
    
    For iter = 1 To 15
        
        'Orthogonalize pairs of columns in upper off diag
        For j = 1 To p - 1
            For k = j + 1 To p
                
                den = 0#
                num = 0#
                'Perform single plane rotation
                For i = 1 To p
                    num = num + 2 * A(i, j) * A(i, k)   ': numerator eq. 11
                    den = den + (A(i, j) + A(i, k)) * _
                        (A(i, j) - A(i, k))             ': denominator eq. 11
                Next i
                
                'Skip rotation if aij is zero and correct ordering
                If Abs(num) < eps And den >= 0 Then Exit For
                
                'Perform Rotation
                If Abs(num) <= Abs(den) Then
                    Tan2 = Abs(num) / Abs(den)          ': eq. 11
                    Cos2 = 1 / Sqr(1 + Tan2 * Tan2)     ': eq. 12
                    Sin2 = Tan2 * Cos2                  ': eq. 13
                Else
                    Cot2 = Abs(den) / Abs(num)          ': eq. 16
                    Sin2 = 1 / Sqr(1 + Cot2 * Cot2)     ': eq. 17
                    Cos2 = Cot2 * Sin2                  ': eq. 18
                End If
                
                Cos_ = Sqr((1 + Cos2) / 2)              ': eq. 14/19
                Sin_ = Sin2 / (2 * Cos_)                ': eq. 15/20
                
                If den < 0 Then
                    tmp = Cos_
                    Cos_ = Sin_                         ': table 21
                    Sin_ = tmp
                End If
                
                Sin_ = Sgn(num) * Sin_                  ': sign table 21
                
                'Rotate
                For i = 1 To p
                    tmp = A(i, j)
                    A(i, j) = tmp * Cos_ + A(i, k) * Sin_
                    A(i, k) = -tmp * Sin_ + A(i, k) * Cos_
                Next i
                
            Next k
        Next j
        
        'Test for convergence
        Test = Application.SumSq(A)
        If Abs(Test - hold) < eps And iter > 5 Then Exit For
        hold = Test
    Next iter
    
    If iter = 16 Then MsgBox "JK Iteration has not converged."
    
    'Compute eigenvalues/eigenvectors
    For j = 1 To p
        'Compute eigenvalues
        For k = 1 To p
            Ematrix(j, 1) = Ematrix(j, 1) + A(k, j) ^ 2
        Next k
        Ematrix(j, 1) = Sqr(Ematrix(j, 1))
        
        'Normalize eigenvectors
        For i = 1 To p
            If Ematrix(j, 1) <= 0 Then
                Ematrix(i, j + 1) = 0
            Else
                Ematrix(i, j + 1) = A(i, j) / Ematrix(j, 1)
            End If
        Next i
    Next j
        
    EIGEN_JK = Ematrix
    
    Exit Function
    
EndProc:
    MsgBox prompt:="Error in function EIGEN_JK!" & vbCr & vbCr & _
        "Error: " & Err.Description & ".", Buttons:=48, _
        Title:="Run time error!"
End Function

Function matrixElementWise(ByRef inputArr, opVal As Long, numVal As Double) As Double()
' Calculator for element-wise operations.
' Add, multiply, divide, or exponentiate every cell in a matrix by a scalar.
' Mode:
'     1: Element-wise addition
'     2: Element-wise multiplication (scalar multiplication)
'     3: Element-wise division
'     4: Element-wise power
'
    Dim L2 As Long
    Dim L3 As Long
    
    ReDim resultArr( _
        LBound(inputArr) To UBound(inputArr), _
        LBound(inputArr, 2) To UBound(inputArr, 2)) As Double
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        For L3 = LBound(inputArr, 2) To UBound(inputArr, 2)
            If opVal = 1 Then ' Addition
                resultArr(L2, L3) = inputArr(L2, L3) + numVal
            ElseIf opVal = 2 Then ' Multiplication
                resultArr(L2, L3) = inputArr(L2, L3) * numVal
            ElseIf opVal = 3 Then ' Division
                If numVal <> 0 Then
                    resultArr(L2, L3) = inputArr(L2, L3) / numVal
                End If
            ElseIf opVal = 4 Then ' Power
                If inputArr(L2, L3) <> 0 Then
                    resultArr(L2, L3) = inputArr(L2, L3) ^ numVal
                End If
            End If
        Next L3
    Next L2
    
    matrixElementWise = resultArr
End Function

Function matrixPower(ByRef inputArr, powerN As Long) As Variant
' Calculate the n-th power of a matrix.
'
    Dim L2 As Long
    Dim invArr As Variant
    Dim resultArr As Variant
    
    If matrixIsSquare(inputArr) Then
        If powerN > 0 Then
            resultArr = inputArr
            
            For L2 = 2 To powerN
                resultArr = matrixMultiply(resultArr, inputArr)
            Next L2
        ElseIf powerN = 0 Then
            resultArr = createIdentityMatrix(UBound(inputArr) - LBound(inputArr) + 1)
        ElseIf powerN < 0 Then
            invArr = matrixGauss(inputArr)
            
            If invArr(4) <> 0 Then
                resultArr = matrixPower(invArr(3), -powerN)
            Else
                resultArr = "Undefined"
            End If
        End If
        
        matrixPower = resultArr
    Else
        matrixPower = "Undefined"
    End If
End Function

Function matrixDeterminant(ByRef inArr) As Variant
    matrixDeterminant = matrixGauss(inArr, 4)
End Function

Function matrixGauss(ByRef inArr, Optional singVal = 0) As Variant
' Gauss(-Jordan) elimination with partial pivoting.
' Returns an array of [REF(A), RREF(A), inv(A), det(A)].
'
    Dim L2 As Long
    Dim L3 As Long
    Dim sRow As Long
    Dim sCol As Long
    Dim pivotRow As Long
    Dim pivotVal As Double
    Dim rowMult As Double
    Dim findInv As Boolean
    Dim findDet As Boolean
    Dim priArr As Variant
    Dim matrixDet As Variant
    Dim matrixREF As Variant
    Dim matrixRREF As Variant
    Dim matrixInv As Variant
    Dim outArr(1 To 4) As Variant
    
    priArr = inArr
    If matrixIsSquare(priArr) Then
        findInv = True
        findDet = True
        matrixInv = createIdentityMatrix(UBound(priArr) - LBound(priArr) + 1)
        matrixDet = 1
    Else
        If singVal = 3 Or singVal = 4 Then
            matrixGauss = "Undefined"
            Exit Function
        End If
        
        matrixInv = "Undefined"
        matrixDet = "Undefined"
    End If
    
    ' Forward elimination.
    sRow = LBound(priArr)
    sCol = LBound(priArr, 2)
    Do While sRow <= UBound(priArr) And sCol <= UBound(priArr, 2)
        ' Partial pivoting, find the largest absolute value in the column.
        pivotVal = 0
        For L2 = sRow To UBound(priArr)
            If Abs(pivotVal) < Abs(priArr(L2, sCol)) Then
                pivotVal = priArr(L2, sCol)
                pivotRow = L2
            End If
        Next L2
        
        If pivotVal = 0 Then
            If singVal = 3 Then
                matrixGauss = "Undefined"
                Exit Function
            ElseIf singVal = 4 Then
                matrixGauss = 0
                Exit Function
            End If
            
            ' Stop further calculations.
            findInv = False
            findDet = False
            
            matrixInv = "Undefined"
            
            If matrixIsSquare(priArr) Then
                matrixDet = 0
            Else
                matrixDet = "Undefined"
            End If
            
            sCol = sCol + 1
        Else
            If findDet Then matrixDet = matrixDet * pivotVal
            
            ' Swap rows if necessary.
            If pivotRow <> sRow Then
                For L2 = sCol To UBound(priArr, 2)
                    swapValues priArr(pivotRow, L2), priArr(sRow, L2)
                Next L2
                
                If findDet Then matrixDet = matrixDet * -1
                If findInv Then swapRows matrixInv, pivotRow, sRow
            End If
            
            ' Subtract multiples of current row from all lower rows.
            For L2 = sRow + 1 To UBound(priArr)
                rowMult = priArr(L2, sCol) / pivotVal
                priArr(L2, sCol) = 0

                For L3 = sCol + 1 To UBound(priArr, 2)
                    priArr(L2, L3) = priArr(L2, L3) - (rowMult * priArr(sRow, L3))
                Next L3
                
                If findInv Then
                    For L3 = LBound(priArr, 2) To UBound(priArr, 2)
                        matrixInv(L2, L3) = matrixInv(L2, L3) - (rowMult * matrixInv(sRow, L3))
                    Next L3
                End If
            Next L2
            
            sRow = sRow + 1
            sCol = sCol + 1
        End If
    Loop
    
    If singVal = 1 Then
        matrixGauss = priArr
        Exit Function
    ElseIf singVal = 4 Then
        matrixGauss = matrixDet
        Exit Function
    End If
    
    matrixREF = priArr
    
    ' Backward elimination.
    sRow = UBound(priArr)
    sCol = UBound(priArr, 2)
    Do While sRow >= LBound(priArr) And sCol >= LBound(priArr, 2)
        pivotVal = 0
        For L2 = LBound(priArr, 2) To UBound(priArr, 2)
            If priArr(sRow, L2) <> 0 Then
                pivotVal = priArr(sRow, L2)
                sCol = L2
                Exit For
            End If
        Next L2
        
        If pivotVal = 0 Then
            sRow = sRow - 1
        Else
            ' Divide current row by pivot value.
            priArr(sRow, sCol) = 1
            For L2 = LBound(priArr, 2) To UBound(priArr, 2)
                If L2 > sCol Then priArr(sRow, L2) = priArr(sRow, L2) / pivotVal
                If findInv Then matrixInv(sRow, L2) = matrixInv(sRow, L2) / pivotVal
            Next L2
            
            ' Subtract multiples of current row from all upper rows.
            For L2 = sRow - 1 To LBound(priArr) Step -1
                rowMult = priArr(L2, sCol)
                priArr(L2, sCol) = 0

                For L3 = UBound(priArr, 2) To sCol + 1 Step -1
                    priArr(L2, L3) = priArr(L2, L3) - (rowMult * priArr(sRow, L3))
                Next L3
                
                If findInv Then
                    For L3 = UBound(priArr, 2) To LBound(priArr, 2) Step -1
                        matrixInv(L2, L3) = matrixInv(L2, L3) - (rowMult * matrixInv(sRow, L3))
                    Next L3
                End If
            Next L2
            
            sRow = sRow - 1
        End If
    Loop
    
    matrixRREF = priArr
    
    If singVal = 2 Then
        matrixGauss = priArr
    ElseIf singVal = 3 Then
        matrixGauss = matrixInv
    ElseIf singVal = 0 Then
        outArr(1) = matrixREF
        outArr(2) = matrixRREF
        outArr(3) = matrixInv
        outArr(4) = matrixDet
        matrixGauss = outArr
    End If
End Function

Function matrixTrace(ByRef inputArr) As Variant
' Calculate the trace of a 2D square matrix.
'
    Dim L2 As Long
    
    ' Defensive programming.
    If Not matrixIsSquare(inputArr) Then
        matrixTrace = "Undefined"
        Exit Function
    End If
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        matrixTrace = matrixTrace + inputArr(L2, L2)
    Next L2
End Function

Function matrixMultiply(ByRef lArr, ByRef rArr) As Variant
' Multiply two 2D matrices.
' Remember: Matrix multiplication is noncommutative.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim leftCols As Long
    Dim rightRows As Long
    
    leftCols = UBound(lArr, 2) - LBound(lArr, 2) + 1
    rightRows = UBound(rArr) - LBound(rArr) + 1
    
    If leftCols <> rightRows Then
        matrixMultiply = "Undefined"
        Exit Function
    End If
    
    ' Result array would have as many rows as left array and as many columns as right array.
    ReDim outArr(LBound(lArr) To UBound(lArr), LBound(rArr, 2) To UBound(rArr, 2)) As Double
    
    For L2 = LBound(outArr) To UBound(outArr)
        For L3 = LBound(outArr, 2) To UBound(outArr, 2)
            For L4 = LBound(rArr) To UBound(rArr)
                outArr(L2, L3) = outArr(L2, L3) + (lArr(L2, L4) * rArr(L4, L3))
            Next L4
        Next L3
    Next L2
    
    matrixMultiply = outArr
End Function

Function matrixTranspose(ByRef inArr) As Double()
' Transpose a 2D matrix.
'
    Dim inRow As Long
    Dim inCol As Long
    ReDim outArr(LBound(inArr, 2) To UBound(inArr, 2), LBound(inArr) To UBound(inArr)) As Double
    
    For inRow = LBound(inArr) To UBound(inArr)
        For inCol = LBound(inArr, 2) To UBound(inArr, 2)
            outArr(inCol, inRow) = inArr(inRow, inCol)
        Next inCol
    Next inRow
    
    matrixTranspose = outArr
End Function

Function createIdentityMatrix(ByRef sizeN As Long) As Double()
' Create a size N identity matrix.
'
    Dim L2 As Long
    ReDim identityArr(1 To sizeN, 1 To sizeN) As Double
    
    For L2 = 1 To sizeN
        identityArr(L2, L2) = 1
    Next L2
    
    createIdentityMatrix = identityArr
End Function

Function swapRows(ByRef inputArr, Row1, Row2)
    Dim L2 As Long
    For L2 = LBound(inputArr, 2) To UBound(inputArr, 2)
        swapValues inputArr(Row1, L2), inputArr(Row2, L2)
    Next L2
End Function

Function swapValues(ByRef Variable1, ByRef Variable2)
    tempVar = Variable1
    Variable1 = Variable2
    Variable2 = tempVar
End Function

' ================================================================================================
' Unused Functions
' ================================================================================================

Function matrixAdjugate(ByRef inputArr) As Variant
' Calculate the adjugate (classical adjoint) matrix of a square matrix.
'
    Dim L2 As Long
    Dim L3 As Long
    
    ' Defensive programming.
    If Not matrixIsSquare(inputArr) Then
        matrixAdjugate = "Undefined"
        Exit Function
    End If
    
    ReDim cofactMat( _
        LBound(inputArr) To UBound(inputArr), _
        LBound(inputArr, 2) To UBound(inputArr, 2)) As Double
    
    ' Calculate the cofactor matrix.
    For L2 = LBound(cofactMat) To UBound(cofactMat)
        For L3 = LBound(cofactMat, 2) To UBound(cofactMat, 2)
            cofactMat(L2, L3) = singleCofactor(inputArr, L2, L3)
        Next L3
    Next L2
    
    matrixAdjugate = matrixTranspose(cofactMat)
End Function

Function singleCofactor(ByRef inputArr, iRow, jCol) As Double
' Calculate the cofactor of A(i,j).
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim L5 As Long
    
    ReDim tempArr( _
        LBound(inputArr) To UBound(inputArr) - 1, _
        LBound(inputArr, 2) To UBound(inputArr, 2) - 1) As Double
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        If L2 <> iRow Then
            L4 = L4 + 1
            For L3 = LBound(inputArr, 2) To UBound(inputArr, 2)
                If L3 <> jCol Then
                    L5 = L5 + 1
                    tempArr(L4, L5) = inputArr(L2, L3)
                End If
            Next L3
            L5 = 0
        End If
    Next L2
    
    If (iRow + jCol) Mod 2 = 0 Then
        singleCofactor = matrixDeterminant(tempArr)
    Else
        singleCofactor = -matrixDeterminant(tempArr)
    End If
End Function

Function KroneckerProduct(ByRef leftArr, ByRef rightArr) As Double()
' Reference:
'     Langville, A. N., & Stewart, W. J. (2004).
'     The Kronecker product and stochastic automata networks
'     Journal of Computational and Applied Mathematics, 167(2), 429-447.
'
' Calculate the Kronecker product/Zehfuss product of two matrices.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim L5 As Long
    Dim leftRows As Long
    Dim leftCols As Long
    Dim rightRows As Long
    Dim rightCols As Long
    Dim blockRow As Long
    Dim blockCol As Long
    
    leftRows = UBound(leftArr, 1) - LBound(leftArr, 1) + 1
    leftCols = UBound(leftArr, 2) - LBound(leftArr, 2) + 1
    rightRows = UBound(rightArr, 1) - LBound(rightArr, 1) + 1
    rightCols = UBound(rightArr, 2) - LBound(rightArr, 2) + 1
    
    ReDim outputArr(1 To leftRows * rightRows, 1 To leftCols * rightCols) As Double
    
    For L2 = 1 To leftRows
        blockRow = (L2 - 1) * rightRows
        For L3 = 1 To leftCols
            blockCol = (L3 - 1) * rightCols
            For L4 = 1 To rightRows
                For L5 = 1 To rightCols
                    outputArr(blockRow + L4, blockCol + L5) = leftArr(L2, L3) * rightArr(L4, L5)
                Next L5
            Next L4
        Next L3
    Next L2
    
    KroneckerProduct = outputArr
End Function

' ================================================================================================
' Procedural Functions/Subroutines - May need to be changed
' ================================================================================================

Sub createButtonsAndInstructions()
'
    Dim L2 As Long
    Dim instructionString() As String
    
    Sheets.Add Sheets(1)
    
    instructionString = Split( _
        "Linear Algebra Algorithms" & vbTab & _
        vbNullString & vbTab & _
        "Instructions:" & vbTab & _
        "    Create a new sheet or book to enter data in (or use this sheet)." & vbTab & _
        "    Enter 1-2 matrices near the top-left corner of the spreadsheet." & vbTab & _
        "        Within the first 500x500 cells would be safe." & vbTab & _
        "        The matrix closer to the top-left corner would be read first." & vbTab & _
        "    If the buttons are not visible:" & vbTab & _
        "        Press Alt+F8, select the 'createButtons' macro and run it." & vbTab & _
        "    Use the buttons to perform operations on the matrices." & vbTab & _
        "    Enter matrices first BEFORE clicking the buttons!" & vbTab & _
        "    There should be at least one cell spacing between matrices." & vbTab & _
        "    Note: The 'Reposition buttons' button moves them to the selected cell." & vbTab & _
        "    Have fun!", vbTab)
    
    For L2 = 0 To UBound(instructionString)
        Cells(1 + L2, 1) = instructionString(L2)
    Next L2
    
    callCreateButtons
End Sub

Function createButton(PositionRng, ButtonText, Optional CallFunction = "doNothing")
    With PositionRng
        ActiveSheet.Buttons.Add(.Left, .Top, .Width, .Height).Text = ButtonText
    End With
    
    With ActiveSheet
        With .Shapes(.Shapes.Count)
            .Placement = xlMoveAndSize
            .OnAction = CallFunction
        End With
    End With
End Function

Function doNothing()
    
End Function

Sub callCreateButtons(Optional FirstPosition = "L4:O5")
'
    Const numOfButtons As Long = 17
    
    Dim L2 As Long
    Dim buttonLabels() As String
    Dim macroList() As String
    
    buttonLabels = Split( _
        "," & _
        "Transpose single matrix," & _
        "Invert single matrix," & _
        "Matrix power single matrix," & _
        "Multiply two matrices," & _
        "Multiply three matrices," & _
        "Multiply multiple matrices," & _
        "Eigendecompose symmetric matrix," & _
        "Element-wise operation single matrix," & _
        "(Reduced) Row echelon form," & _
        "Previous sheet," & _
        "Next sheet," & _
        "Delete sheet warning on," & _
        "Delete sheet warning off," & _
        "Delete sheet," & _
        "Delete all but 1st sheet," & _
        "Reposition buttons," & _
        "Delete buttons", ",")
    
    macroList = Split( _
        "," & _
        "oneMatrixTrans," & _
        "oneMatrixInv," & _
        "callOneMatrixPower," & _
        "twoMatrixProd," & _
        "threeMatrixProd," & _
        "multiMatrixProd," & _
        "symMatEigenDecom," & _
        "callOneMatrixElementwise," & _
        "oneMatrix_R_REF," & _
        "previousSheet," & _
        "nextSheet," & _
        "deleteSheetLockOn," & _
        "deleteSheetLockOff," & _
        "deleteSheet," & _
        "deleteAllSheets," & _
        "repositionButtons," & _
        "deleteButtons,", ",")
    
    For L2 = 1 To numOfButtons
        createButton Range(FirstPosition).Offset((L2 - 1) * 2, 0), buttonLabels(L2), macroList(L2)
    Next L2
End Sub

Private Sub callOneMatrixPower()
'
    Do
        matrixPow = InputBox("Please enter the power to raise to:", "Select power", 2)
        
        If LenB(matrixPow) = 0 Then
            MsgBox "No changes made. Exiting script.": Exit Sub
        ElseIf Not IsNumeric(matrixPow) Then
            MsgBox "Please enter a number.", , "Invalid entry"
        ElseIf Not isInteger(matrixPow) Then
            MsgBox "Please enter an integer.", , "Invalid entry"
        End If
    Loop While Not IsNumeric(matrixPow) Or Not isInteger(matrixPow)
    
    oneMatrixPower CLng(matrixPow)
End Sub

Private Sub multiMatrixProd()
' Only allow positive integer input.
'
    Do
        matrixCount = InputBox( _
            "Please enter number of matrices to multiply:", _
            "Input number of matrices", 2)
        
        If LenB(matrixCount) = 0 Then
            MsgBox "No changes made. Exiting script.": Exit Sub
        ElseIf Not IsNumeric(matrixCount) Then
            MsgBox "Please enter a number.", , "Invalid entry"
        ElseIf Not isInteger(matrixCount, True) Then
            MsgBox "Please enter a positive integer.", , "Invalid entry"
        End If
    Loop While Not IsNumeric(matrixCount) Or Not isInteger(matrixCount, True)
    
    nMatrixProduct CLng(matrixCount)
End Sub

Private Sub callOneMatrixElementwise()
' Only allow 1, 2, 3, 4 for opVal.
' Only allow non-zero numVal for opVal 3 (division).
' Only allow non-zero integer numVal for opVal 4 (power).
'
    Dim loopCheck As Boolean
    Dim promptStr() As String
    
    promptStr = Split( _
        ",Please enter number to add by:" & _
        ",Please enter number to multiply by:" & _
        ",Please enter number to divide by (0 not accepted):" & _
        ",Please enter a non-zero integer power (0 not accepted):", ",")
    
    Do
        opVal = InputBox( _
            "Please select an element-wise operation:" & vbCr & _
            "  1. Addition" & vbCr & _
            "  2. Multiplication" & vbCr & _
            "  3. Division" & vbCr & _
            "  4. Power", "Select operator", 1)
        
        loopCheck = opVal <> 1 And opVal <> 2 And opVal <> 3 And opVal <> 4
        
        If LenB(opVal) = 0 Then
            MsgBox "No changes made. Exiting script.": Exit Sub
        ElseIf loopCheck Then
            MsgBox "Please enter a valid number.", , "Invalid entry"
        End If
    Loop While loopCheck
    
    Do
        numVal = InputBox(promptStr(opVal), "Enter value", 1)
        
        If LenB(numVal) = 0 Then
            MsgBox "No changes made. Exiting script.": Exit Sub
        ElseIf Not IsNumeric(numVal) Then
            MsgBox "Please enter a number.", , "Invalid entry"
        ElseIf opVal = 3 Then
            If numVal = 0 Then MsgBox "Please enter a non-zero number.", , "Invalid entry"
        ElseIf opVal = 4 Then
            If Not isInteger(numVal) Or numVal = 0 Then
                MsgBox "Please enter a non-zero integer.", , "Invalid entry"
            End If
        End If
        
        loopCheck = _
            Not IsNumeric(numVal) Or _
            (opVal = 3 And numVal = 0) Or _
            (opVal = 4 And Not isInteger(numVal))
    Loop While loopCheck
    
    oneMatrixElementwise CLng(opVal), CDbl(numVal)
End Sub

' ================================================================================================
' Procedural Functions/Subroutines - Not likely to be changed
' ================================================================================================

Sub createButtons()
    callCreateButtons
End Sub

Private Sub twoMatrixProd()
    nMatrixProduct 2
End Sub

Private Sub threeMatrixProd()
    nMatrixProduct 3
End Sub

Private Sub previousSheet()
    If ActiveSheet.Index > 1 Then
        ActiveSheet.Previous.Select
    End If
End Sub

Private Sub nextSheet()
    If ActiveSheet.Index < Sheets.Count Then
        ActiveSheet.Next.Select
    End If
End Sub

Private Sub deleteSheetLockOn()
    deleteSheetUnlock = False
End Sub

Private Sub deleteSheetLockOff()
    deleteSheetUnlock = True
End Sub

Private Sub deleteSheet(Optional lockOverride As Boolean = False)
    If Sheets.Count > 1 Then
        If deleteSheetUnlock Or lockOverride Then
            Application.DisplayAlerts = False
            ActiveSheet.Delete
            Application.DisplayAlerts = True
        Else
            ActiveSheet.Delete
        End If
    End If
End Sub

Private Sub deleteAllSheets()
    finalConfirmation = MsgBox("Are you sure?", vbYesNo, "Delete all but first sheet")
    If finalConfirmation = vbYes Then
        Application.DisplayAlerts = False
        Do While Sheets.Count > 1
            Sheets(Sheets.Count).Delete
        Loop
        Application.DisplayAlerts = True
    End If
End Sub

Private Sub repositionButtons()
    ActiveSheet.Buttons.Delete
    callCreateButtons ActiveCell.Resize(2, 4).Address
End Sub

Sub deleteButtons()
    ActiveSheet.Buttons.Delete
End Sub

' ================================================================================================
' Secret Projects
' ================================================================================================

Function sqrtComplex(ByRef inputVal) As Variant
' Enable calculation of the square root of negative numbers.
' Return complex numbers if necessary.
'
    If inputVal < 0 Then
        sqrtComplex = Sqr(-inputVal) & "i"
    Else
        sqrtComplex = Sqr(inputVal)
    End If
End Function

Private Sub plotBlandAltman()
' Receive array of repeated measures.
' Calculate means and differences of repeated measures, mean and SD of differences, LoA.
'
    ' First row for printing output.
    Const outRow1 As Long = 2
    
    Dim L2 As Long
    Dim numOfRows As Long
    Dim numOfCols As Long
    Dim diffMean As Double
    Dim diffSD As Double
    Dim loaSpan As Double
    Dim loaLBound As Double
    Dim loaUBound As Double
    Dim inputArr As Variant
    
    ' Create a copy of the input, because the scanner function will delete arrays.
    ActiveSheet.Copy , ActiveSheet
    inputArr = arrScanTopLeft
    
    If Not IsArray(inputArr) Then
        MsgBox "No arrays found. Exiting script."
        Exit Sub
    ElseIf Not matrixHasNonNum(inputArr) Then
        MsgBox "Empty or non-numeric cell detected. Exiting script."
        Exit Sub
    End If
    
    ' Delete everything else from the sheet.
    ActiveSheet.UsedRange.Clear
    
    ReDim meanArr(LBound(inputArr) To UBound(inputArr)) As Double
    ReDim diffArr(LBound(inputArr) To UBound(inputArr)) As Double
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        meanArr(L2) = (inputArr(L2, LBound(inputArr, 2)) + inputArr(L2, UBound(inputArr, 2))) / 2
        diffArr(L2) = inputArr(L2, LBound(inputArr, 2)) - inputArr(L2, UBound(inputArr, 2))
    Next L2
    
    diffMean = Application.Average(diffArr)
    diffSD = Application.StDev(diffArr)
    loaSpan = diffSD * Application.NormSInv(0.975)
    loaLBound = diffMean - loaSpan
    loaUBound = diffMean + loaSpan
    
    numOfRows = UBound(inputArr) - LBound(inputArr) + 1
    numOfCols = UBound(inputArr, 2) - LBound(inputArr, 2) + 1
    
    Cells(outRow1 - 1, outCol1).Value2 = "A"
    Cells(outRow1 - 1, outCol1 + 1).Value2 = "B"
    Cells(outRow1 - 1, outCol1 + 2).Value2 = "Mean(AB)"
    Cells(outRow1 - 1, outCol1 + 3).Value2 = "A-B"
    
    Range(Cells(outRow1, outCol1), Cells(outRow1 + numOfRows - 1, outCol1 + 1)).Value2 = inputArr
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        Cells(outRow1 + L2 - 1, outCol1 + 2) = meanArr(L2)
        Cells(outRow1 + L2 - 1, outCol1 + 3) = diffArr(L2)
    Next L2
    
    Cells(outRow1 - 1, outCol1 + 5).Value2 = "xMinMax"
    Cells(outRow1 - 1, outCol1 + 6).Value2 = "meanDiff"
    Cells(outRow1 - 1, outCol1 + 7).Value2 = "loaLower"
    Cells(outRow1 - 1, outCol1 + 8).Value2 = "loaUpper"
    
    Cells(outRow1, outCol1 + 5).Value2 = Application.Min(meanArr)
    Cells(outRow1 + 1, outCol1 + 5).Value2 = Application.Max(meanArr)
    Cells(outRow1, outCol1 + 6).Value2 = diffMean
    Cells(outRow1 + 1, outCol1 + 6).Value2 = diffMean
    Cells(outRow1, outCol1 + 7).Value2 = loaLBound
    Cells(outRow1 + 1, outCol1 + 7).Value2 = loaLBound
    Cells(outRow1, outCol1 + 8).Value2 = loaUBound
    Cells(outRow1 + 1, outCol1 + 8).Value2 = loaUBound
End Sub

Function calcMeanSd1D(ByRef inputArr) As Double()
'
    Dim L2 As Long
    Dim valCount As Long
    Dim valSum As Double
    Dim valCenSumSq As Double
    Dim outputArr(1 To 2) As Double
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        valCount = valCount + 1
        valSum = valSum + inputArr(L2)
    Next L2
    
    outputArr(1) = valSum / valCount ' Mean
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        valCenSumSq = valCenSumSq + ((inputArr(L2) - outputArr(1)) ^ 2)
    Next L2
    
    outputArr(2) = Sqr(valCenSumSq / (valCount - 1)) ' SD (with Bessel's correction)
    
    calcMeanSd1D = outputArr
End Function
