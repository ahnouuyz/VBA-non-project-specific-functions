' ================================================================================================
' Math Functions
'     Functions NOT available in Excel
' ================================================================================================

Function fnCalcCV(Mean, SD) As Variant
' Calculate the coefficient of variation.
'
    If Mean <> 0 Then
        fnCalcCV = SD / Mean
    Else
        fnCalcCV = errDiv0
    End If
End Function

Function fnHodgesLehmann(ByRef inputArr) As Variant
' References:
'     Hodges, J. L., & Lehmann, E. L. (1963).
'     Estimates of Location Based on Rank Tests.
'     Ann. Math. Statist., 34(2), pp. 598-611.
'
'     Lehmann, E. L. (1963).
'     Nonparametric Confidence Intervals for a Shift Parameter.
'     The Annals of Mathematical Statistics, 34(4), pp. 1507-1512.
'
' Receive an array of numeric values.
' Return Hodges-Lehmann estimate for the median, IQR, 95CI lo, and 95CI up.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim arrLen As Long
    Dim walshLen As Long
    Dim ci95LoIndex As Long
    Dim ci95UpIndex As Long
    Dim wilcoxonMean As Double
    Dim wilcoxonVariance As Double
    Dim ci95CritZ As Double
    Dim ci95span As Double
    Dim outputArr(0 To 3) As Variant
    
    arrLen = UBound(inputArr) - LBound(inputArr) + 1
    
    ' Calculate Walsh averages for all non-repeated pairs.
    ReDim walschArr(1 To 1) As Double
    For L2 = LBound(inputArr) To UBound(inputArr)
        For L3 = L2 To UBound(inputArr)
            walshLen = walshLen + 1
            ReDim Preserve walschArr(1 To walshLen) As Double
            walschArr(walshLen) = (inputArr(L3) + inputArr(L2)) / 2
        Next L3
    Next L2
    
    wilcoxonMean = walshLen / 2
    wilcoxonVariance = wilcoxonMean * (2 * arrLen + 1) / 6
    ci95CritZ = stdNormCdfInv(1 - typeOneErr / 2 / numOfComps) ' Bonferroni-correction
    ci95span = ci95CritZ * Sqr(wilcoxonVariance)
    ci95LoIndex = Int(wilcoxonMean - ci95span) + 1 ' <~~ 17 (for n = 16)
    ci95UpIndex = Int(wilcoxonMean + ci95span) + 1 ' <~~ 120 (for n = 16)
    
    outputArr(0) = fnArrQuartile(walschArr, 2) ' Median
    outputArr(1) = fnArrQuartile(walschArr, 3) - fnArrQuartile(walschArr, 1) ' IQR
    
    If ci95LoIndex < 1 Or ci95UpIndex > walshLen Then
        Debug.Print "Error: Walsh index for 95%CI out of bounds (sample size may be too small)."
        MsgBox "Error: Walsh index for 95%CI out of bounds (sample size may be too small)."
        outputArr(2) = "NA"
        outputArr(3) = "NA"
    End If
    
    outputArr(2) = QuickSelect(walschArr, 1, walshLen, ci95LoIndex) ' 95CI lower bound
    outputArr(3) = QuickSelect(walschArr, 1, walshLen, ci95UpIndex) ' 95CI upper bound
    
    fnHodgesLehmann = outputArr
End Function

Function fnCohensD(dEstimator, dSpread, dCorrel) As Variant
' Reference:
'     Dunlap, W. P., Cortina, J. M., Vaslow, J. B., & Burke, M. J. (1996).
'     Meta-analysis of experiments with matched groups or repeated measures designs.
'     Psychological Methods, 1(2), pp. 170-177.
'
'     Knudson, D. (2009).
'     Significant and meaningful effects in sports biomechanics research.
'     Sports Biomechanics, 8(1), pp. 96-104.
'
'     Hopkins, W. G. (2002).
'     A scale of magnitudes for effect statistics.
'     Retrieved Date Accessed, 2017.
'         <0.10 ~~> Trivial difference.
'         0.10 ~~> Small difference.
'         0.20 ~~> Medium difference.
'         0.60 ~~> Large difference.
'         1.20 ~~> Very large difference.
'
' Effect size estimation using Cohen's d for repeated measures.
' The correlation between two conditions is taken into account.
'
    Const correlDecimalPlaces As Long = 5
    
    If dSpread <> 0 And IsNumeric(dCorrel) Then
        fnCohensD = dEstimator / dSpread * Sqr(2 * (1 - Round(dCorrel, correlDecimalPlaces)))
    Else
        fnCohensD = errDiv0
    End If
End Function

Function stdNormCdf(ByVal zScore As Double) As Double
' Source:
'     https://www.johndcook.com/blog/python_phi/
'
' Standard normal cumulative distribution function (Phi).
' Receive Z score.
' Return area under standard normal distribution to the left of Z score.
'
' Max error = 6.96877223704817E-08 (at Z = -0.0638389) cf. Excel's .NormSDist function.
'
    Const a1 As Double = 0.254829592
    Const a2 As Double = -0.284496736
    Const a3 As Double = 1.421413741
    Const a4 As Double = -1.453152027
    Const a5 As Double = 1.061405429
    Const p As Double = 0.3275911
    
    Dim Sign As Long
    Dim t As Double
    Dim y As Double
    
    ' Save the sign of zScore.
    Sign = Sgn(zScore)
    zScore = Abs(zScore) / Sqr(2)

    ' Abramowitz and Stegun Formula 7.1.26.
    t = 1 / (1 + p * zScore)
    y = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Exp(-zScore * zScore)

    stdNormCdf = 0.5 * (1 + Sign * y)
End Function

Function stdNormCdfInv(ByVal p As Double) As Double
' Source:
'     Adapted for Microsoft Visual Basic from Peter Acklam's
'     "An algorithm for computing the inverse normal cumulative distribution function"
'     (http://home.online.no/~pjacklam/notes/invnorm/)
'     by John Herrero (3-Jan-03)
'
' Max error = 3.92294685624961E-09 (at p = 0.0002262) cf. Excel's .NormSInv function.
'
    ' Define coefficients in rational approximations.
    Const a5 As Double = -39.6968302866538
    Const a4 As Double = 220.946098424521
    Const a3 As Double = -275.928510446969
    Const a2 As Double = 138.357751867269
    Const a1 As Double = -30.6647980661472
    Const a0 As Double = 2.50662827745924
    
    Const b5 As Double = -54.4760987982241
    Const b4 As Double = 161.585836858041
    Const b3 As Double = -155.698979859887
    Const b2 As Double = 66.8013118877197
    Const b1 As Double = -13.2806815528857
    
    Const c5 As Double = -7.78489400243029E-03
    Const c4 As Double = -0.322396458041136
    Const c3 As Double = -2.40075827716184
    Const c2 As Double = -2.54973253934373
    Const c1 As Double = 4.37466414146497
    Const c0 As Double = 2.93816398269878
    
    Const d4 As Double = 7.78469570904146E-03
    Const d3 As Double = 0.32246712907004
    Const d2 As Double = 2.445134137143
    Const d1 As Double = 3.75440866190742
    
    ' Define break-points.
    Const p_low As Double = 0.02425
    Const p_high As Double = 1 - p_low
    
    ' Define work variables.
    Dim q As Double
    Dim r As Double
    
    ' If argument out of bounds, raise error.
    If p <= 0 Or p >= 1 Then Err.Raise 5
    
    If p < p_low Then
      ' Rational approximation for lower region.
      q = Sqr(-2 * Log(p))
      stdNormCdfInv = (((((c5 * q + c4) * q + c3) * q + c2) * q + c1) * q + c0) / _
        ((((d4 * q + d3) * q + d2) * q + d1) * q + 1)
    ElseIf p <= p_high Then
      ' Rational approximation for lower region.
      q = p - 0.5
      r = q * q
      stdNormCdfInv = (((((a5 * r + a4) * r + a3) * r + a2) * r + a1) * r + a0) * q / _
        (((((b5 * r + b4) * r + b3) * r + b2) * r + b1) * r + 1)
    ElseIf p < 1 Then
      ' Rational approximation for upper region.
      q = Sqr(-2 * Log(1 - p))
      stdNormCdfInv = -(((((c5 * q + c4) * q + c3) * q + c2) * q + c1) * q + c0) / _
        ((((d4 * q + d3) * q + d2) * q + d1) * q + 1)
    End If
End Function

' ================================================================================================
' Statistics: Two-Way Repeated Measures ANOVA
' ================================================================================================

Function rmANOVATwoWay(ByRef inArr3D) As Variant
' Input will be a 3-dimensional array:
'     Dimension 1 ~~> Subject
'     Dimension 2 ~~> Mask
'     Dimension 3 ~~> Shoe
' [N layers (subjects) x L rows (masks) x K columns (shoes)]
' Repeated measures ANOVA for main effects (mask, shoe) and interaction (mask*shoe).
'
' GG and HF estimates for sphericity for interaction effect are unavailable at the moment.
'
    Dim L2 As Long
    Dim L3 As Long
    
    Dim grandMean As Double
    Dim ssTotal As Double
    Dim ssErrFactor1x2 As Double
    Dim dfErrFactor1x2 As Long
    Dim msErrFactor1x2 As Double
    Dim dfAdj As Double
    
    Dim df1D(1 To 3) As Long
    Dim df2D(1 To 3) As Long
    Dim ss1D(1 To 3) As Double
    Dim ss2D(1 To 3) As Double
    Dim ms1D(2 To 3) As Double
    Dim ms2D(1 To 3) As Double
    Dim sphereEst() As Double
    Dim outArr(1 To 3, 0 To 7) As Variant
    
    grandMean = fnArrMean(inArr3D)
    
    ' Calculate sum of squares and 1D degrees of freedom.
    For L2 = 1 To 3
        ' ss1D(1) = ssSubjects (subject)
        ' ss1D(2) = ssFactor1 (mask)
        ' ss1D(3) = ssFactor2 (shoe)
        ss1D(L2) = ssMainEffect(inArr3D, grandMean, L2)
        
        ' ss2D(1) = ssFactor1x2 (mask*shoe) [2*3]
        ' ss2D(2) = ssErrFactor2 = ssFactor2xSubjects (subject*shoe) [1*3]
        ' ss2D(3) = ssErrFactor1 = ssFactor1xSubjects (subject*mask) [1*2]
        ss2D(L2) = ssInteractions(inArr3D, grandMean, L2)
        
        ' df1D(1) = dfSubjects (subjects - 1)
        ' df1D(2) = dfFactor1 (masks - 1)
        ' df1D(3) = dfFactor2 (shoes - 1)
        df1D(L2) = UBound(inArr3D, L2) - LBound(inArr3D, L2)
    Next L2
    
    ' Calculate 2D degrees of freedom and mean squares.
    For L2 = 1 To 3
        ' df2D(1) = dfFactor1x2 = df1D(2) * df1D(3) (dfMasks*dfShoes)
        ' df2D(2) = dfErrFactor1 = df1D(1) * df1D(2) (dfSubjects*dfMasks)
        ' df2D(3) = dfErrFactor2 = df1D(1) * df1D(3) (dfSubjects*dfShoes)
        df2D(L2) = df1D(1 - (L2 = 1)) * df1D(3 + (L2 = 2))
        
        If L2 > 1 Then
            ' ms1D(2) = msFactor1 (mask)
            ' ms1D(3) = msFactor2 (shoe)
            ms1D(L2) = ss1D(L2) / df1D(L2)
        End If
        
        ' ms2D(1) = msFactor1x2 = ss2D(1) / df2D(1) (mask*shoe) [2*3]
        ' ms2D(2) = msErrFactor1 = ss2D(3) / df2D(2) (subject*shoe) [1*3]
        ' ms2D(3) = msErrFactor2 = ss2D(2) / df2D(3) (subject*mask) [1*2]
        ms2D(L2) = ss2D(1 - (L2 = 2) - (L2 > 1)) / df2D(L2)
        
        ' Prepare output array.
        For L3 = 0 To 7
            outArr(L2, L3) = "NA"
        Next L3
    Next L2
    
    ssTotal = fnArrCenSum(inArr3D, 2)
    
    ssErrFactor1x2 = ssTotal - ss1D(1) - ss1D(2) - ss2D(3) - ss1D(3) - ss2D(2) - ss2D(1)
    dfErrFactor1x2 = df1D(1) * df1D(2) * df1D(3)
    msErrFactor1x2 = ssErrFactor1x2 / dfErrFactor1x2
    
    ' Main effects.
    For L2 = 1 To 2
        If ms2D(L2 + 1) <> 0 Then
            ' L2 = 1, Subject by shoe, flatten dim3 (mask averaged)
            ' L2 = 2, Subject by mask, flatten dim2 (shoe averaged)
            sphereEst = sphereEps(covarMat(arr3Dto2D(inArr3D, 2 - (L2 = 1))))
            
            outArr(L2, 0) = sphereEst(2) ' eGG
            outArr(L2, 1) = sphereEst(3) ' eHF
            outArr(L2, 2) = sphereEst(1) ' eLB
            outArr(L2, 3) = ms1D(L2 + 1) / ms2D(L2 + 1) ' F
            outArr(L2, 4) = sphereEst(4) * df1D(L2 + 1) ' df1
            outArr(L2, 5) = sphereEst(4) * df2D(L2 + 1) ' df2
            outArr(L2, 6) = fDistProb(outArr(L2, 3), outArr(L2, 4), outArr(L2, 5)) ' p value
            outArr(L2, 7) = ss1D(L2 + 1) / (ss1D(L2 + 1) + ss2D(2 - (L2 = 1))) ' pEtaSq
        End If
    Next L2
    
    ' Interaction effect.
    If msErrFactor1x2 <> 0 Then
        dfAdj = 1 / df2D(1) ' Lower bound epsilon estimate
        
        outArr(3, 0) = "NA" ' eGG
        outArr(3, 1) = "NA" ' eHF
        outArr(3, 2) = dfAdj ' eLB
        outArr(3, 3) = ms2D(1) / msErrFactor1x2 ' F
        outArr(3, 4) = dfAdj * df2D(1) ' df1
        outArr(3, 5) = dfAdj * dfErrFactor1x2 ' df2
        outArr(3, 6) = fDistProb(outArr(3, 3), outArr(3, 4), outArr(3, 5)) ' p value
        outArr(3, 7) = ss2D(1) / (ss2D(1) + ssErrFactor1x2) ' pEtaSq
    End If
    
    rmANOVATwoWay = outArr
End Function

Function arr3Dto2D(ByRef inArr3D, flatDimension As Long) As Double()
' Convert a 3D array to a 2D array by flattening (taking the average of) one dimension.
'
    ' Defensive programming.
    If flatDimension < 1 Or flatDimension > 3 Then
        Debug.Print "Error: Invalid dimension specified for 2-way AOV."
        Err.Raise 5
        Exit Function
    End If
    
    Dim d1 As Long
    Dim d2 As Long
    Dim d3 As Long
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim dimLen1 As Long
    Dim lineSum As Double
    
    ' Set the dimensions.
    ' 1,23
    ' 2,13
    ' 3,12
    d1 = flatDimension
    d2 = 1 - (d1 = 1)
    d3 = 3 + (d1 = 3)
    
    dimLen1 = UBound(inArr3D, d1) - LBound(inArr3D, d1) + 1
    
    ReDim outArr2D( _
        LBound(inArr3D, d2) To UBound(inArr3D, d2), _
        LBound(inArr3D, d3) To UBound(inArr3D, d3)) As Double
    
    For L2 = LBound(inArr3D, d2) To UBound(inArr3D, d2)
        For L3 = LBound(inArr3D, d3) To UBound(inArr3D, d3)
            lineSum = 0
            For L4 = LBound(inArr3D, d1) To UBound(inArr3D, d1)
                If d1 = 1 Then
                    lineSum = lineSum + inArr3D(L4, L2, L3) ' Line of subjects
                ElseIf d1 = 2 Then
                    lineSum = lineSum + inArr3D(L2, L4, L3) ' Line of masks
                ElseIf d1 = 3 Then
                    lineSum = lineSum + inArr3D(L2, L3, L4) ' Line of shoes
                End If
            Next L4
            outArr2D(L2, L3) = lineSum / dimLen1
        Next L3
    Next L2
    
    arr3Dto2D = outArr2D
End Function

Function ssMainEffect(ByRef inArr3D, grandMean, mainFactor As Long) As Double
' Calculate the sum of squares for a main effect term.
'
    ' Defensive programming.
    If mainFactor < 1 Or mainFactor > 3 Then
        Debug.Print "Error: Invalid main factor specified for 2-way AOV."
        Err.Raise 5
        Exit Function
    End If
    
    Dim d1 As Long
    Dim d2 As Long
    Dim d3 As Long
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim dimLen2 As Long
    Dim dimLen3 As Long
    Dim planeCount As Long
    Dim planeSum As Double
    Dim planeMean As Double
    Dim ssPlane As Double
    
    ' Set the dimensions.
    ' 1,23
    ' 2,13
    ' 3,12
    d1 = mainFactor
    d2 = 1 - (d1 = 1)
    d3 = 3 + (d1 = 3)
    
    dimLen2 = UBound(inArr3D, d2) - LBound(inArr3D, d2) + 1
    dimLen3 = UBound(inArr3D, d3) - LBound(inArr3D, d3) + 1
    planeCount = dimLen2 * dimLen3 ' Number of values in 1 plane
    
    For L2 = LBound(inArr3D, d1) To UBound(inArr3D, d1)
        planeSum = 0
        For L3 = LBound(inArr3D, d2) To UBound(inArr3D, d2)
            For L4 = LBound(inArr3D, d3) To UBound(inArr3D, d3)
                If d1 = 1 Then
                    planeSum = planeSum + inArr3D(L2, L3, L4) ' Plane of 1 subject
                ElseIf d1 = 2 Then
                    planeSum = planeSum + inArr3D(L3, L2, L4) ' Plane of 1 mask
                ElseIf d1 = 3 Then
                    planeSum = planeSum + inArr3D(L3, L4, L2) ' Plane of 1 shoe
                End If
            Next L4
        Next L3
        planeMean = planeSum / planeCount
        ssPlane = ssPlane + ((planeMean - grandMean) ^ 2)
    Next L2
    
    ssMainEffect = ssPlane * planeCount
End Function

Function ssInteractions(ByRef inputArr, grandMean, thirdFactor) As Double
' Calculate the sum of squares for an interaction effect term.
'
    Dim d1 As Long
    Dim d2 As Long
    Dim d3 As Long
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim L5 As Long
    Dim dimLen1 As Long
    Dim dimLen2 As Long
    Dim dimLen3 As Long
    Dim lineSum As Double
    Dim planeSum1 As Double
    Dim planeSum2 As Double
    Dim lMean As Double
    Dim pMean1 As Double
    Dim pMean2 As Double
    
    ' Set the dimensions.
    ' 23,1
    ' 13,2
    ' 12,3
    d3 = thirdFactor
    d1 = 1 - (d3 = 1)
    d2 = 3 + (d3 = 3)
    
    dimLen1 = UBound(inputArr, d1) - LBound(inputArr, d1) + 1
    dimLen2 = UBound(inputArr, d2) - LBound(inputArr, d2) + 1
    dimLen3 = UBound(inputArr, d3) - LBound(inputArr, d3) + 1
    
    For L2 = LBound(inputArr, d1) To UBound(inputArr, d1)
        For L3 = LBound(inputArr, d2) To UBound(inputArr, d2)
            lineSum = 0
            planeSum1 = 0
            planeSum2 = 0
            
            For L4 = LBound(inputArr, d3) To UBound(inputArr, d3)
                If d3 = 1 Then
                    lineSum = lineSum + inputArr(L4, L2, L3) ' Line of subjects
                ElseIf d3 = 2 Then
                    lineSum = lineSum + inputArr(L2, L4, L3) ' Line of masks
                ElseIf d3 = 3 Then
                    lineSum = lineSum + inputArr(L2, L3, L4) ' Line of shoes
                End If
                
                For L5 = LBound(inputArr, d1) To UBound(inputArr, d1)
                    If d3 = 1 Then
                        planeSum1 = planeSum1 + inputArr(L4, L5, L3) ' Plane of shoes
                    ElseIf d3 = 2 Then
                        planeSum1 = planeSum1 + inputArr(L5, L4, L3) ' Plane of shoes
                    ElseIf d3 = 3 Then
                        planeSum1 = planeSum1 + inputArr(L5, L3, L4) ' Plane of masks
                    End If
                Next L5
                
                For L5 = LBound(inputArr, d2) To UBound(inputArr, d2)
                    If d3 = 1 Then
                        planeSum2 = planeSum2 + inputArr(L4, L2, L5) ' Plane of masks
                    ElseIf d3 = 2 Then
                        planeSum2 = planeSum2 + inputArr(L2, L4, L5) ' Plane of subjects
                    ElseIf d3 = 3 Then
                        planeSum2 = planeSum2 + inputArr(L2, L5, L4) ' Plane of subjects
                    End If
                Next L5
            Next L4
            
            lMean = lineSum / dimLen3
            pMean1 = planeSum1 / (dimLen1 * dimLen3)
            pMean2 = planeSum2 / (dimLen2 * dimLen3)
            
            ssInteractions = ssInteractions + ((lMean - pMean1 - pMean2 + grandMean) ^ 2)
        Next L3
    Next L2
    
    ssInteractions = ssInteractions * dimLen3
End Function

' ================================================================================================
' Statistics: One-Way Repeated Measures ANOVA
' ================================================================================================

Function fTableSumSq(ByRef inArr) As Double()
' Calculate ssRows, ssCols, and ssWithin for an F table.
'
' F table
' ==============================================
' Subj | Conditions          |        |        |
' ==============================================
'      | 01  02  03  ..   k  |  avgR  |        |
' ---- | --  --  --  --  --  | ------ |        |
'   01 |                     |        | ssW_1  |
'   02 |                     |        | ssW_2  |
'   03 |                     |        | ssW_3  |
'   04 |                     |        | ssW_4  |
'   .. |                     |        |  ...   |
'   .. |                     |        |  ...   |
'    n |                     |        | ssW_n  |
' ----------------------------------------------
' avgC |                     |  avgT  | ssC/n  |
' ----------------------------------------------
'      |                     | ssR/k  |        |
' ==============================================
'
' Used by:
'     1. One-way repeated measures ANOVA
'     2. Intraclass correlation coefficients (ICCs)
'     3. Friedman's test
'
' Output as [ssRow, ssCol, ssWit].
'
    Dim L2 As Long
    Dim L3 As Long
    Dim inRows As Long
    Dim inCols As Long
    Dim rowSum As Double
    Dim colSum As Double
    Dim rowMean As Double
    Dim colMean As Double
    Dim planeMean As Double
    Dim ssRow As Double
    Dim ssCol As Double
    Dim ssWit As Double
    Dim outArr(1 To 3) As Double
    
    inRows = UBound(inArr) - LBound(inArr) + 1
    inCols = UBound(inArr, 2) - LBound(inArr, 2) + 1
    
    planeMean = fnArrMean(inArr)
    
    For L2 = LBound(inArr) To UBound(inArr)
        rowSum = 0
        For L3 = LBound(inArr, 2) To UBound(inArr, 2)
            rowSum = rowSum + inArr(L2, L3)
        Next L3
        
        rowMean = rowSum / inCols
        ssRow = ssRow + (rowMean - planeMean) ^ 2
        
        For L3 = LBound(inArr, 2) To UBound(inArr, 2)
            ssWit = ssWit + (inArr(L2, L3) - rowMean) ^ 2
        Next L3
    Next L2
    ssRow = ssRow * inCols
    
    For L2 = LBound(inArr, 2) To UBound(inArr, 2)
        colSum = 0
        For L3 = LBound(inArr) To UBound(inArr)
            colSum = colSum + inArr(L3, L2)
        Next L3
        
        colMean = colSum / inRows
        ssCol = ssCol + (colMean - planeMean) ^ 2
    Next L2
    ssCol = ssCol * inRows
    
    outArr(1) = ssRow
    outArr(2) = ssCol
    outArr(3) = ssWit
    fTableSumSq = outArr
End Function

Function rmANOVA(ByRef inArr2D) As Variant
' Input is a [N rows (subjects) x K columns (shoes)] 2D array.
'
    Dim L2 As Long
    Dim dfRows As Long
    Dim dfCols As Long
    Dim msCol As Double
    Dim msErr As Double
    Dim dfAdj As Double
    Dim fTab() As Double
    Dim sphereEst As Variant
    Dim outArr(0 To 7) As Variant
    
    dfRows = UBound(inArr2D) - LBound(inArr2D)
    dfCols = UBound(inArr2D, 2) - LBound(inArr2D, 2)
    
    fTab = fTableSumSq(inArr2D)
    
    msCol = fTab(2) / dfCols ' ssCol / dfCols
    msErr = (fTab(3) - fTab(2)) / (dfRows * dfCols) ' (ssWithin - ssCol) / (dfRows * dfCols)
    
    If msErr <> 0 Then
        sphereEst = sphereEps(covarMat(inArr2D))
    
        dfAdj = sphereEst(4)
        
        outArr(0) = sphereEst(2) ' eGG
        outArr(1) = sphereEst(3) ' eHF
        outArr(2) = sphereEst(1) ' eLB
        outArr(3) = msCol / msErr ' F
        outArr(4) = dfAdj * dfCols ' df1
        outArr(5) = dfAdj * dfRows * dfCols ' df2
        outArr(6) = fDistProb(outArr(3), outArr(4), outArr(5)) ' p value
        outArr(7) = fTab(2) / fTab(3) ' pEtaSq = ssCol / ssWithin
    Else
        ' No variance in table.
        For L2 = 0 To 7
            outArr(L2) = "NA"
        Next L2
    End If
    
    rmANOVA = outArr
End Function

Function covarMat(ByRef inArr2D) As Double()
' Receive [n rows x k columns] array.
' Return [k x k] covariance matrix.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    
    ReDim tempArr1(LBound(inArr2D) To UBound(inArr2D)) As Double
    ReDim tempArr2(LBound(inArr2D) To UBound(inArr2D)) As Double
    ReDim outArr2D( _
        LBound(inArr2D, 2) To UBound(inArr2D, 2), _
        LBound(inArr2D, 2) To UBound(inArr2D, 2)) As Double
    
    For L2 = LBound(inArr2D, 2) To UBound(inArr2D, 2)
        For L3 = L2 To UBound(inArr2D, 2)
            For L4 = LBound(inArr2D) To UBound(inArr2D)
                tempArr1(L4) = inArr2D(L4, L2)
                tempArr2(L4) = inArr2D(L4, L3)
            Next L4
            
            outArr2D(L2, L3) = fnCovar(tempArr1, tempArr2)
            If L3 <> L2 Then outArr2D(L3, L2) = outArr2D(L2, L3) ' Symmetric matrix
        Next L3
    Next L2
    
    covarMat = outArr2D
End Function

Function sphereEps(ByRef inArr) As Double()
' Reference:
'     Girden, E. R. (1992).
'     ANOVA: Repeated measures.
'     Newbury Park, Calif: Sage Publications.
'         ~~> If the average of eGG and eHF is >0.75, then use eHF. Use eGG otherwise.
'
'     Greenhouse, S.W. & Geisser, S. (1959).
'     On methods in the analysis of profile data.
'     Psychometrika, 24(2), 95-112.
'
'     Huynh Huynh, & Feldt, L. (1976).
'     Estimation of the Box Correction for Degrees of Freedom from
'     Sample Data in Randomized Block and Split-Plot Designs.
'     Journal of Educational Statistics, 1(1), 69-82.
'
' Receive [k x k] covariance matrix.
' Output as [eLB, eGG, eHF, eRec].
'
    Dim L2 As Long
    Dim L3 As Long
    Dim inCols As Long
    Dim ssCell As Double
    Dim arrMean As Double
    Dim colSum As Double
    Dim colMean As Double
    Dim ssColMean As Double
    Dim diagSum As Double
    Dim diagMean As Double
    Dim eLB As Double
    Dim ggNum As Double
    Dim ggDen As Double
    Dim eGG As Double
    Dim hfNum As Double
    Dim hfDen As Double
    Dim eHF As Double
    Dim eRec As Double
    Dim outArr(1 To 4) As Double
    
    inCols = UBound(inArr) - LBound(inArr) + 1 ' covarCols = covarRows
    
    eLB = 1 / (inCols - 1) ' Lower bound estimate
    
    ssCell = fnArrRawSum(inArr, 2)
    arrMean = fnArrMean(inArr)
    
    For L2 = LBound(inArr) To UBound(inArr)
        colSum = 0
        For L3 = LBound(inArr) To UBound(inArr)
            colSum = colSum + inArr(L3, L2)
        Next L3
        
        colMean = colSum / inCols
        ssColMean = ssColMean + (colMean ^ 2)
        
        diagSum = diagSum + inArr(L2, L2) ' Sum of variances
    Next L2
    diagMean = diagSum / inCols ' Mean variance
    
    ' Greenhouse-Geisser.
    ggNum = (inCols * (diagMean - arrMean)) ^ 2
    ggDen = (inCols - 1) * (ssCell - 2 * inCols * ssColMean + (inCols * arrMean) ^ 2)
    eGG = ggNum / ggDen ' Greenhouse-Geisser
    
    ' Huynh-Feldt.
    hfNum = numOfSubjs * (inCols - 1) * eGG - 2
    hfDen = (inCols - 1) * (numOfSubjs - 1 - (inCols - 1) * eGG)
    eHF = hfNum / hfDen ' Huynh-Feldt
    If eHF > 1 Then eHF = 1
    
    If (eGG + eHF) / 2 > 0.75 Then
        eRec = eHF
    Else
        eRec = eGG
    End If
    
    outArr(1) = eLB
    outArr(2) = eGG
    outArr(3) = eHF
    outArr(4) = eRec
    sphereEps = outArr
End Function

Function fDistProb(fStat, df1, df2) As Variant
' Input F, df1, and df2 (fractional degrees of freedom accepted).
' Output p value (calculated using bilinear interpolation of harmonic/reciprocal dfs).
'
    If IsNumeric(fStat) And IsNumeric(df1) And IsNumeric(df2) Then
        
        ' Defensive programming.
        If fStat < 0 Or df1 < 0 Or df2 < 0 Then
            Debug.Print "Invalid parameter(s) detected for .FDIST function:"
            Debug.Print "F = " & fStat
            Debug.Print "df1 = " & df1
            Debug.Print "df2 = " & df2
            Err.Raise 5
            Exit Function
        End If
        
        Dim df11 As Long
        Dim df12 As Long
        Dim df21 As Long
        Dim df22 As Long
        Dim p11 As Double
        Dim p12 As Double
        Dim p21 As Double
        Dim p22 As Double
        Dim q1 As Double
        Dim q2 As Double
        Dim q3 As Double
        Dim q4 As Double
        Dim grandDenom As Double
        
        ' Find the nearest whole-numbered dfs to assign as lower and upper bounds.
        df11 = Int(df1) - (Int(df1) = 0) ' <~~ 1 if result is 0
        df12 = df11 + 1
        df21 = Int(df2) - (Int(df2) = 0) ' <~~ 1 if result is 0
        df22 = df21 + 1
        
        ' Find the boundary p values for the four combinations of dfs.
        With Application
            p11 = .FDist(fStat, df11, df21)
            p12 = .FDist(fStat, df11, df22)
            p21 = .FDist(fStat, df12, df21)
            p22 = .FDist(fStat, df12, df22)
        End With
        
        ' 2D linear interpolation using harmonic dfs.
        q1 = (1 / df12 - 1 / df1) * (1 / df22 - 1 / df2) * p11
        q2 = (1 / df12 - 1 / df1) * (1 / df2 - 1 / df21) * p12
        q3 = (1 / df1 - 1 / df11) * (1 / df22 - 1 / df2) * p21
        q4 = (1 / df1 - 1 / df11) * (1 / df2 - 1 / df21) * p22
        grandDenom = (1 / df12 - 1 / df11) * (1 / df22 - 1 / df21)
        
        fDistProb = (q1 + q2 + q3 + q4) / grandDenom
    Else
        fDistProb = "NA"
    End If
End Function

' ================================================================================================
' Statistics: Intraclass Correlation Coefficient
' ================================================================================================

Function meanICCs(ByRef allStepsArr) As Variant
' Calculate all average ICCs for k steps from all possible k-wise (n choose k) combinations.
' Input is a 2D array, with subjects by rows and steps by columns.
' Return array of averaged ICCs for 2, 3, 4, ..., n steps.
'
    Const deLim1 As String = "," ' Delimiter between items
    Const deLim2 As String = ";" ' Delimiter between combinations
    
    Dim iStepCombo As Long
    Dim iSubj As Long
    Dim stepNum As Long
    Dim numOfSteps As Long
    Dim maxNumOfSteps As Long
    Dim iccCount As Long
    Dim iccSum As Double
    Dim stepCombiArr() As String
    Dim stepNumArr() As String
    Dim icc3kSingle As Variant
    
    maxNumOfSteps = UBound(allStepsArr, 2) - LBound(allStepsArr, 2) + 1 ' Number of columns
    
    ReDim iccFinalArr(2 To maxNumOfSteps) As Variant
    
    For numOfSteps = 2 To maxNumOfSteps
        iccSum = 0
        iccCount = 0
        stepCombiArr = Split(nonRepeatCombis(Empty, Empty, maxNumOfSteps, numOfSteps), deLim2)
        
        ReDim subSetArr(LBound(allStepsArr) To UBound(allStepsArr), 1 To numOfSteps) As Double
        
        For iStepCombo = LBound(stepCombiArr) To UBound(stepCombiArr)
            stepNumArr = Split(stepCombiArr(iStepCombo), deLim1)
            
            ' Copy relevant data to the subset array.
            For iSubj = LBound(allStepsArr) To UBound(allStepsArr)
                For stepNum = 1 To numOfSteps
                    subSetArr(iSubj, stepNum) = allStepsArr(iSubj, stepNumArr(stepNum - 1))
                Next stepNum
            Next iSubj
            
            icc3kSingle = icc3k(subSetArr)
            
            If IsNumeric(icc3kSingle) Then
                iccSum = iccSum + icc3kSingle
                iccCount = iccCount + 1
            End If
        Next iStepCombo
        
        If iccCount <> 0 Then
            iccFinalArr(numOfSteps) = iccSum / iccCount
        Else
            iccFinalArr(numOfSteps) = "NA"
        End If
    Next numOfSteps
    
    meanICCs = iccFinalArr
End Function

Function nonRepeatCombis(sFinal, s2, maxNum, kItems, Optional i2 = 1) As String
' List all possible k-wise combinations without repeats from n (n choose k).
' Ascending lexicographical order (1,2,3;1,2,4;...;1,k-1,k;2,3,4;...;k-2,k-1,k).
' Returns a double-delimited string.
'
    Const deLim1 As String = "," ' Delimiter between items
    Const deLim2 As String = ";" ' Delimiter between combinations
    
    Dim L2 As Long
    Dim newStr As String
    Dim dummyStr As String
    
    If kItems = 0 Then
        ' Base case.
        s2 = Left(s2, Len(s2) - Len(deLim1)) ' Remove last deLim1
        sFinal = sFinal & (s2 & deLim2) ' Add most recent combination and delimiter
        Exit Function
    ElseIf kItems > maxNum Then
        ' Invalid case.
        Debug.Print "Invalid number of items - greater than maximum."
        Err.Raise 5
        Exit Function
    End If
    
    For L2 = i2 To maxNum - (kItems - 1)
        newStr = s2 & (L2 & deLim1)
        dummyStr = newStr & nonRepeatCombis(sFinal, newStr, maxNum, kItems - 1, L2 + 1)
    Next L2
    
    nonRepeatCombis = Left(sFinal, Len(sFinal) - Len(deLim2)) ' Remove last deLim2
End Function

Function icc3k(ByRef inArr2D) As Double
' References:
'     Bartko, J. J. (1976).
'     On various intraclass correlation reliability coefficients.
'     Psychological Bulletin, 83(5), 762–765.
'         ~~> Set -ve ICCs to 0 if taking the average of multiple ICCs.
'
'     Koo, T. K., & Li, M. Y. (2016).
'     A Guideline of Selecting and Reporting Intraclass
'     Correlation Coefficients for Reliability Research.
'     Journal of Chiropractic Medicine, 15(2), 155–163.
'         ~~> Flowchart for selecting the appropriate ICC.
'         ~~> Reference for the various ICC equations.
'
' Receives an array with [N rows (subjects) x K columns (conditions)].
' Calculate ICC(3,k), absolute agreement.
' Two-way mixed effects, absolute agreement, multiple raters/measurements.
'
    Dim dfRows As Long
    Dim dfCols As Long
    Dim msRow As Double
    Dim msCol As Double
    Dim msErr As Double
    Dim icc3kNum As Double
    Dim icc3kDenom As Double
    Dim fTab() As Double
    
    dfRows = UBound(inArr2D) - LBound(inArr2D)
    dfCols = UBound(inArr2D, 2) - LBound(inArr2D, 2)
    
    fTab = fTableSumSq(inArr2D)
    
    msRow = fTab(1) / dfRows ' ssRow / dfRows
    msCol = fTab(2) / dfCols ' ssCol / dfCols
    msErr = (fTab(3) - fTab(2)) / (dfRows * dfCols) ' (ssWithin - ssCol) / (dfRows * dfCols)
    
    ' ICC(3,k) equation.
    icc3kNum = msRow - msErr
    icc3kDenom = msRow + (msCol - msErr) / (dfRows + 1)
    
    If icc3kNum = 0 And icc3kDenom = 0 Then
        icc3k = 1
    ElseIf icc3kDenom <> 0 Then
        icc3k = icc3kNum / icc3kDenom
        If icc3k < 0 Then icc3k = 0
    Else
        Debug.Print "Division by zero."
        Err.Raise 5
    End If
End Function

' ================================================================================================
' Statistics: Nonparametric Tests
' ================================================================================================

Function rankFunction(iVal, ByRef numList) As Double
' Calculate the rank of a value from an array of values.
' Highest rank for highest value, average rank for ties.
'
    Const decimalPlaces As Long = 10
    
    Dim V2 As Variant
    Dim d2 As Double
    Dim d3 As Double
    
    rankFunction = 0.5
    d2 = Round(iVal, decimalPlaces)
    For Each V2 In numList
        d3 = Round(V2, decimalPlaces)
        If d2 > d3 Then
            rankFunction = rankFunction + 1
        ElseIf d2 = d3 Then
            rankFunction = rankFunction + 0.5
        End If
    Next V2
End Function

Function npTestFriedman(ByRef inArr) As Variant
' Reference:
'     Friedman, M. (1937).
'     The Use of Ranks to Avoid the Assumption of Normality Implicit in the Analysis of Variance.
'     Journal of the American Statistical Association, 32(200), 675-701.
'
'     Friedman, M. (1939).
'     A Correction: The Use of Ranks to Avoid the Assumption
'     of Normality Implicit in the Analysis of Variance.
'     Journal of the American Statistical Association, 34(205), 109-109.
'
'     Friedman, M. (1940).
'     A Comparison of Alternative Tests of Significance for the Problem of m Rankings.
'     The Annals of Mathematical Statistics, 11(1), 86-92.
'
'     Kendall, M., & Smith, B. (1939).
'     The Problem of m Rankings.
'     The Annals of Mathematical Statistics, 10(3), 275-287.
'
' Friedman's Test.
' Receive an array with [N rows (subjects) x K columns (conditions)].
' Output as [chiSq, df, p value, Kendall's W].
'
    Dim L2 As Long
    Dim L3 As Long
    Dim dfCols As Long
    Dim inRows As Long
    Dim inCols As Long
    Dim ssT As Double
    Dim ssE As Double
    Dim ssDev As Double
    Dim kwDenom As Double
    Dim fTab() As Double
    Dim outArr(0 To 3) As Variant
    
    ReDim ranksArr( _
        LBound(inArr) To UBound(inArr), _
        LBound(inArr, 2) To UBound(inArr, 2)) As Double
    ReDim tempRow( _
        LBound(inArr, 2) To UBound(inArr, 2)) As Double
    
    inRows = UBound(inArr, 1) - LBound(inArr, 1) + 1
    dfCols = UBound(inArr, 2) - LBound(inArr, 2)
    inCols = dfCols + 1
    
    ' Create array of ranks.
    For L2 = LBound(inArr) To UBound(inArr)
        For L3 = LBound(inArr, 2) To UBound(inArr, 2)
            tempRow(L3) = inArr(L2, L3)
        Next L3
        
        For L3 = LBound(tempRow) To UBound(tempRow)
            ranksArr(L2, L3) = rankFunction(tempRow(L3), tempRow)
        Next L3
    Next L2
    
    fTab = fTableSumSq(ranksArr)
    
    ssT = fTab(2) ' ssT = ssCol
    ssDev = ssT * inRows ' ssDev = ssCol * N
    ssE = fnArrCenSum(ranksArr, 2) / (inRows * dfCols)
    kwDenom = (((inCols * inCols) - 1) * inCols) * (inRows * inRows)
    
    If ssE <> 0 Then
        outArr(0) = ssT / ssE ' chiSq
        outArr(2) = Application.ChiDist(outArr(0), dfCols) ' p (Friedman)
    Else
        ' Ranks are all the same
        outArr(0) = "NA"
        outArr(2) = "NA"
    End If
    
    outArr(1) = dfCols ' df
    outArr(3) = 12 * ssDev / kwDenom ' Kendall's W
    
    npTestFriedman = outArr
End Function

Function npTestWilcoxonSR(ByRef inputArr) As Variant
' Reference:
'     Wilcoxon, F. (1945).
'     Individual Comparisons by Ranking Methods.
'     Biometrics Bulletin, 1(6), 80-83.
'
' Wilcoxon Signed-Rank Test.
' Receive an array of paired differences.
' Return an array of [Z, p value].
'
    Dim L2 As Long
    Dim newN As Long
    Dim tieCount As Long
    Dim tieAdjust As Long
    Dim tieCompare As Double
    Dim valRank As Double
    Dim sumSignRank As Double
    Dim wilcoxVar1 As Double
    Dim wilcoxVar2 As Double
    Dim zVal As Double
    Dim pVal As Double
    Dim V2 As Variant
    Dim outputArr(0 To 1) As Variant
    
    ReDim tempArr(1 To 1) As Double
    
    ' Count number of non-zero values.
    For Each V2 In inputArr
        If V2 <> 0 Then
            newN = newN + 1
            ReDim Preserve tempArr(1 To newN) As Double
            tempArr(newN) = V2
        End If
    Next V2
    
    ' Skip the test if all values are 0.
    If newN = 0 Then
        outputArr(0) = "NA"
        outputArr(1) = "NA"
        npTestWilcoxonSR = outputArr
        Exit Function
    End If
    
    QuickSort tempArr, 1, newN, 1
    
    ReDim absArr(1 To newN) As Double
    
    ' Store absolute values in new array (for ranking).
    For L2 = 1 To newN
        absArr(L2) = Abs(tempArr(L2))
    Next L2
    
    For L2 = 1 To newN
        valRank = rankFunction(absArr(L2), absArr)
        sumSignRank = sumSignRank + (valRank * Sgn(tempArr(L2)))
        
        ' Adjust variance for tied ranks.
        If tieCompare <> valRank Then
            tieCompare = valRank
            tieAdjust = tieAdjust + ((tieCount * tieCount - 1) * tieCount)
            tieCount = 0
        Else
            tieCount = tieCount + 1
        End If
    Next L2
    
    wilcoxVar1 = ((2 * newN + 3) * newN + 1) * newN / 6 ' <~~ N(N+1)(2N+1)/6
    wilcoxVar2 = wilcoxVar1 - tieAdjust / 12
    
    zVal = sumSignRank / Sqr(wilcoxVar2)
    If zVal > 0 Then zVal = -zVal ' <~~ Convert all Z scores to negative
    
    pVal = stdNormCdf(zVal) * 2 * numOfComps ' <~~ Two-tailed; Bonferroni correction
    If pVal > 1 Then pVal = 1
    
    outputArr(0) = zVal
    outputArr(1) = pVal
    
    npTestWilcoxonSR = outputArr
End Function

' ================================================================================================
' Statistics: Shapiro-Wilk Test of Normality
' ================================================================================================

Function shapiroWilkBelow25(ByVal inputArr) As Variant
' Reference:
'     Shapiro, S., & Wilk, M. (1965).
'     An Analysis of Variance Test for Normality (Complete Samples).
'     Biometrika, 52(3/4), 591-611.
'
' Shapiro-Wilk test (for sample sizes 25 and below only, for now).
' Receive array of individual means used to derive group mean.
' W values raised to 20th power for linear interpolation.
' Return p value.
'
    Const powerTransform As Long = 20
    
    Dim L2 As Long
    Dim kSW As Long
    Dim ssTotal As Double
    Dim bCoeff As Double
    Dim wStat As Double
    Dim w1 As Double
    Dim w2 As Double
    Dim p1 As Double
    Dim p2 As Double
    Dim pVal As Double
    Dim aCoeff() As Double
    Dim pDistr() As Double
    
    ' Limitation.
    If numOfSubjs > 25 Then
        Debug.Print "Shapiro-Wilk test only available for n = 25 or less at the moment. Sorry."
        shapiroWilkBelow25 = "NA"
        Exit Function
    End If
    
    ssTotal = fnArrCenSum(inputArr, 2)
    
    If ssTotal = 0 Then ' <~~ Zero variance - not normally distributed
        shapiroWilkBelow25 = 0
        Exit Function
    End If
    
    kSW = (numOfSubjs + 1) \ 2
    aCoeff = shapiroWilkTable5(numOfSubjs, kSW)
    pDistr = shapiroWilkTable6(numOfSubjs)
    
    QuickSort inputArr, LBound(inputArr), UBound(inputArr), 1
    
    For L2 = 1 To kSW
        If numOfSubjs Mod 2 = 0 Then
            ' Even number of subjects.
            bCoeff = bCoeff + (aCoeff(L2) * (inputArr(kSW * 2 + 1 - L2) - inputArr(L2)))
        Else
            ' Odd number of subjects.
            If L2 = kSW Then Exit For ' aCoeff(kSW) = 0 anyway
            bCoeff = bCoeff + (aCoeff(L2) * (inputArr(kSW * 2 - L2) - inputArr(L2)))
        End If
    Next L2
    
    wStat = bCoeff * bCoeff / ssTotal
    
    For L2 = 0 To 10
        If wStat = pDistr(L2, 0) Then
            pVal = pDistr(L2, 1)
            Exit For
        ElseIf wStat < pDistr(L2, 0) Then
            w1 = pDistr(L2 - 1, 0) ^ powerTransform
            w2 = pDistr(L2, 0) ^ powerTransform
            wStat = wStat ^ powerTransform
            p1 = pDistr(L2 - 1, 1)
            p2 = pDistr(L2, 1)
            pVal = p1 + (p2 - p1) * (wStat - w1) / (w2 - w1)
            Exit For
        End If
    Next L2
    
    If pVal < 0 Then pVal = 0
    shapiroWilkBelow25 = pVal
End Function

Function shapiroWilkTable5(nSubjs As Long, kSW As Long) As Double()
' Reference:
'     Shapiro, S., & Wilk, M. (1965).
'     An Analysis of Variance Test for Normality (Complete Samples).
'     Biometrika, 52(3/4), 591-611.
'
' Table 5 (partial).
'
    ReDim aCoeff(1 To kSW) As Double
    
    If nSubjs = 2 Then
        aCoeff(1) = 0.7071
    ElseIf nSubjs = 3 Then
        aCoeff(1) = 0.7071
        aCoeff(2) = 0
    ElseIf nSubjs = 4 Then
        aCoeff(1) = 0.6872
        aCoeff(2) = 0.1677
    ElseIf nSubjs = 5 Then
        aCoeff(1) = 0.6646
        aCoeff(2) = 0.2413
        aCoeff(3) = 0
    ElseIf nSubjs = 6 Then
        aCoeff(1) = 0.6431
        aCoeff(2) = 0.2806
        aCoeff(3) = 0.0875
    ElseIf nSubjs = 7 Then
        aCoeff(1) = 0.6233
        aCoeff(2) = 0.3031
        aCoeff(3) = 0.1401
        aCoeff(4) = 0
    ElseIf nSubjs = 8 Then
        aCoeff(1) = 0.6052
        aCoeff(2) = 0.3164
        aCoeff(3) = 0.1743
        aCoeff(4) = 0.0561
    ElseIf nSubjs = 9 Then
        aCoeff(1) = 0.5888
        aCoeff(2) = 0.3244
        aCoeff(3) = 0.1976
        aCoeff(4) = 0.0947
        aCoeff(5) = 0
    ElseIf nSubjs = 10 Then
        aCoeff(1) = 0.5739
        aCoeff(2) = 0.3291
        aCoeff(3) = 0.2141
        aCoeff(4) = 0.1224
        aCoeff(5) = 0.0399
    ElseIf nSubjs = 11 Then
        aCoeff(1) = 0.5601
        aCoeff(2) = 0.3315
        aCoeff(3) = 0.226
        aCoeff(4) = 0.1429
        aCoeff(5) = 0.0695
        aCoeff(6) = 0
    ElseIf nSubjs = 12 Then
        aCoeff(1) = 0.5475
        aCoeff(2) = 0.3325
        aCoeff(3) = 0.2347
        aCoeff(4) = 0.1586
        aCoeff(5) = 0.0922
        aCoeff(6) = 0.0303
    ElseIf nSubjs = 13 Then
        aCoeff(1) = 0.5359
        aCoeff(2) = 0.3325
        aCoeff(3) = 0.2412
        aCoeff(4) = 0.1707
        aCoeff(5) = 0.1099
        aCoeff(6) = 0.0539
        aCoeff(7) = 0
    ElseIf nSubjs = 14 Then
        aCoeff(1) = 0.5251
        aCoeff(2) = 0.3318
        aCoeff(3) = 0.246
        aCoeff(4) = 0.1802
        aCoeff(5) = 0.124
        aCoeff(6) = 0.0727
        aCoeff(7) = 0.024
    ElseIf nSubjs = 15 Then
        aCoeff(1) = 0.515
        aCoeff(2) = 0.3306
        aCoeff(3) = 0.2495
        aCoeff(4) = 0.1878
        aCoeff(5) = 0.1353
        aCoeff(6) = 0.088
        aCoeff(7) = 0.0433
        aCoeff(8) = 0
    ElseIf nSubjs = 16 Then
        aCoeff(1) = 0.5056
        aCoeff(2) = 0.329
        aCoeff(3) = 0.2521
        aCoeff(4) = 0.1939
        aCoeff(5) = 0.1447
        aCoeff(6) = 0.1005
        aCoeff(7) = 0.0593
        aCoeff(8) = 0.0196
    ElseIf nSubjs = 17 Then
        aCoeff(1) = 0.4968
        aCoeff(2) = 0.3273
        aCoeff(3) = 0.254
        aCoeff(4) = 0.1988
        aCoeff(5) = 0.1524
        aCoeff(6) = 0.1109
        aCoeff(7) = 0.0725
        aCoeff(8) = 0.0359
        aCoeff(9) = 0
    ElseIf nSubjs = 18 Then
        aCoeff(1) = 0.4886
        aCoeff(2) = 0.3253
        aCoeff(3) = 0.2553
        aCoeff(4) = 0.2027
        aCoeff(5) = 0.1587
        aCoeff(6) = 0.1197
        aCoeff(7) = 0.0837
        aCoeff(8) = 0.0496
        aCoeff(9) = 0.0163
    ElseIf nSubjs = 19 Then
        aCoeff(1) = 0.4808
        aCoeff(2) = 0.3232
        aCoeff(3) = 0.2561
        aCoeff(4) = 0.2059
        aCoeff(5) = 0.1641
        aCoeff(6) = 0.1271
        aCoeff(7) = 0.0932
        aCoeff(8) = 0.0612
        aCoeff(9) = 0.0303
        aCoeff(10) = 0
    ElseIf nSubjs = 20 Then
        aCoeff(1) = 0.4734
        aCoeff(2) = 0.3211
        aCoeff(3) = 0.2565
        aCoeff(4) = 0.2085
        aCoeff(5) = 0.1686
        aCoeff(6) = 0.1334
        aCoeff(7) = 0.1013
        aCoeff(8) = 0.0711
        aCoeff(9) = 0.0422
        aCoeff(10) = 0.014
    ElseIf nSubjs = 21 Then
        aCoeff(1) = 0.4643
        aCoeff(2) = 0.3185
        aCoeff(3) = 0.2578
        aCoeff(4) = 0.2119
        aCoeff(5) = 0.1736
        aCoeff(6) = 0.1399
        aCoeff(7) = 0.1092
        aCoeff(8) = 0.0804
        aCoeff(9) = 0.053
        aCoeff(10) = 0.0263
        aCoeff(11) = 0
    ElseIf nSubjs = 22 Then
        aCoeff(1) = 0.459
        aCoeff(2) = 0.3156
        aCoeff(3) = 0.2571
        aCoeff(4) = 0.2131
        aCoeff(5) = 0.1764
        aCoeff(6) = 0.1443
        aCoeff(7) = 0.115
        aCoeff(8) = 0.0878
        aCoeff(9) = 0.0618
        aCoeff(10) = 0.0368
        aCoeff(11) = 0.0122
    ElseIf nSubjs = 23 Then
        aCoeff(1) = 0.4542
        aCoeff(2) = 0.3126
        aCoeff(3) = 0.2563
        aCoeff(4) = 0.2139
        aCoeff(5) = 0.1787
        aCoeff(6) = 0.148
        aCoeff(7) = 0.1201
        aCoeff(8) = 0.0941
        aCoeff(9) = 0.0696
        aCoeff(10) = 0.0459
        aCoeff(11) = 0.0228
        aCoeff(12) = 0
    ElseIf nSubjs = 24 Then
        aCoeff(1) = 0.4493
        aCoeff(2) = 0.3098
        aCoeff(3) = 0.2554
        aCoeff(4) = 0.2145
        aCoeff(5) = 0.1807
        aCoeff(6) = 0.1512
        aCoeff(7) = 0.1245
        aCoeff(8) = 0.0997
        aCoeff(9) = 0.0764
        aCoeff(10) = 0.0539
        aCoeff(11) = 0.0321
        aCoeff(12) = 0.0107
    ElseIf nSubjs = 25 Then
        aCoeff(1) = 0.445
        aCoeff(2) = 0.3069
        aCoeff(3) = 0.2543
        aCoeff(4) = 0.2148
        aCoeff(5) = 0.1822
        aCoeff(6) = 0.1539
        aCoeff(7) = 0.1283
        aCoeff(8) = 0.1046
        aCoeff(9) = 0.0823
        aCoeff(10) = 0.061
        aCoeff(11) = 0.0403
        aCoeff(12) = 0.02
        aCoeff(13) = 0
    ElseIf nSubjs = 26 Then
        ' Update if necessary.
        
    End If
    
    shapiroWilkTable5 = aCoeff
End Function

Function shapiroWilkTable6(nSubjs As Long) As Double()
' Reference:
'     Shapiro, S., & Wilk, M. (1965).
'     An Analysis of Variance Test for Normality (Complete Samples).
'     Biometrika, 52(3/4), 591-611.
'
' Table 6 (partial).
'
    Dim pDistr(0 To 10, 0 To 1) As Double ' <~~ 0: W value, 1: p value
    
    ' p values.
    pDistr(0, 1) = 0
    pDistr(1, 1) = 0.01
    pDistr(2, 1) = 0.02
    pDistr(3, 1) = 0.05
    pDistr(4, 1) = 0.1
    pDistr(5, 1) = 0.5
    pDistr(6, 1) = 0.9
    pDistr(7, 1) = 0.95
    pDistr(8, 1) = 0.98
    pDistr(9, 1) = 0.99
    pDistr(10, 1) = 1
    
    ' W values.
    pDistr(0, 0) = 0
    pDistr(10, 0) = 1
    
    If nSubjs = 3 Then
        pDistr(1, 0) = 0.753
        pDistr(2, 0) = 0.756
        pDistr(3, 0) = 0.767
        pDistr(4, 0) = 0.789
        pDistr(5, 0) = 0.959
        pDistr(6, 0) = 0.998
        pDistr(7, 0) = 0.999
        pDistr(8, 0) = 1
        pDistr(9, 0) = 1
    ElseIf nSubjs = 4 Then
        pDistr(1, 0) = 0.687
        pDistr(2, 0) = 0.707
        pDistr(3, 0) = 0.748
        pDistr(4, 0) = 0.792
        pDistr(5, 0) = 0.935
        pDistr(6, 0) = 0.987
        pDistr(7, 0) = 0.992
        pDistr(8, 0) = 0.996
        pDistr(9, 0) = 0.997
    ElseIf nSubjs = 5 Then
        pDistr(1, 0) = 0.686
        pDistr(2, 0) = 0.715
        pDistr(3, 0) = 0.762
        pDistr(4, 0) = 0.806
        pDistr(5, 0) = 0.927
        pDistr(6, 0) = 0.979
        pDistr(7, 0) = 0.986
        pDistr(8, 0) = 0.991
        pDistr(9, 0) = 0.993
    ElseIf nSubjs = 6 Then
        pDistr(1, 0) = 0.713
        pDistr(2, 0) = 0.743
        pDistr(3, 0) = 0.788
        pDistr(4, 0) = 0.826
        pDistr(5, 0) = 0.927
        pDistr(6, 0) = 0.974
        pDistr(7, 0) = 0.981
        pDistr(8, 0) = 0.986
        pDistr(9, 0) = 0.989
    ElseIf nSubjs = 7 Then
        pDistr(1, 0) = 0.73
        pDistr(2, 0) = 0.76
        pDistr(3, 0) = 0.803
        pDistr(4, 0) = 0.838
        pDistr(5, 0) = 0.928
        pDistr(6, 0) = 0.972
        pDistr(7, 0) = 0.979
        pDistr(8, 0) = 0.985
        pDistr(9, 0) = 0.988
    ElseIf nSubjs = 8 Then
        pDistr(1, 0) = 0.749
        pDistr(2, 0) = 0.778
        pDistr(3, 0) = 0.818
        pDistr(4, 0) = 0.851
        pDistr(5, 0) = 0.932
        pDistr(6, 0) = 0.972
        pDistr(7, 0) = 0.978
        pDistr(8, 0) = 0.984
        pDistr(9, 0) = 0.987
    ElseIf nSubjs = 9 Then
        pDistr(1, 0) = 0.764
        pDistr(2, 0) = 0.791
        pDistr(3, 0) = 0.829
        pDistr(4, 0) = 0.859
        pDistr(5, 0) = 0.935
        pDistr(6, 0) = 0.972
        pDistr(7, 0) = 0.978
        pDistr(8, 0) = 0.984
        pDistr(9, 0) = 0.986
    ElseIf nSubjs = 10 Then
        pDistr(1, 0) = 0.781
        pDistr(2, 0) = 0.806
        pDistr(3, 0) = 0.842
        pDistr(4, 0) = 0.869
        pDistr(5, 0) = 0.938
        pDistr(6, 0) = 0.972
        pDistr(7, 0) = 0.978
        pDistr(8, 0) = 0.983
        pDistr(9, 0) = 0.986
    ElseIf nSubjs = 11 Then
        pDistr(1, 0) = 0.792
        pDistr(2, 0) = 0.817
        pDistr(3, 0) = 0.85
        pDistr(4, 0) = 0.876
        pDistr(5, 0) = 0.94
        pDistr(6, 0) = 0.973
        pDistr(7, 0) = 0.979
        pDistr(8, 0) = 0.984
        pDistr(9, 0) = 0.986
    ElseIf nSubjs = 12 Then
        pDistr(1, 0) = 0.805
        pDistr(2, 0) = 0.828
        pDistr(3, 0) = 0.859
        pDistr(4, 0) = 0.883
        pDistr(5, 0) = 0.943
        pDistr(6, 0) = 0.973
        pDistr(7, 0) = 0.979
        pDistr(8, 0) = 0.984
        pDistr(9, 0) = 0.986
    ElseIf nSubjs = 13 Then
        pDistr(1, 0) = 0.814
        pDistr(2, 0) = 0.837
        pDistr(3, 0) = 0.866
        pDistr(4, 0) = 0.889
        pDistr(5, 0) = 0.945
        pDistr(6, 0) = 0.974
        pDistr(7, 0) = 0.979
        pDistr(8, 0) = 0.984
        pDistr(9, 0) = 0.986
    ElseIf nSubjs = 14 Then
        pDistr(1, 0) = 0.825
        pDistr(2, 0) = 0.846
        pDistr(3, 0) = 0.874
        pDistr(4, 0) = 0.895
        pDistr(5, 0) = 0.947
        pDistr(6, 0) = 0.975
        pDistr(7, 0) = 0.98
        pDistr(8, 0) = 0.984
        pDistr(9, 0) = 0.986
    ElseIf nSubjs = 15 Then
        pDistr(1, 0) = 0.835
        pDistr(2, 0) = 0.855
        pDistr(3, 0) = 0.881
        pDistr(4, 0) = 0.901
        pDistr(5, 0) = 0.95
        pDistr(6, 0) = 0.975
        pDistr(7, 0) = 0.98
        pDistr(8, 0) = 0.984
        pDistr(9, 0) = 0.987
    ElseIf nSubjs = 16 Then
        pDistr(1, 0) = 0.844
        pDistr(2, 0) = 0.863
        pDistr(3, 0) = 0.887
        pDistr(4, 0) = 0.906
        pDistr(5, 0) = 0.952
        pDistr(6, 0) = 0.976
        pDistr(7, 0) = 0.981
        pDistr(8, 0) = 0.985
        pDistr(9, 0) = 0.987
    ElseIf nSubjs = 17 Then
        pDistr(1, 0) = 0.851
        pDistr(2, 0) = 0.869
        pDistr(3, 0) = 0.892
        pDistr(4, 0) = 0.91
        pDistr(5, 0) = 0.954
        pDistr(6, 0) = 0.977
        pDistr(7, 0) = 0.981
        pDistr(8, 0) = 0.985
        pDistr(9, 0) = 0.987
    ElseIf nSubjs = 18 Then
        pDistr(1, 0) = 0.858
        pDistr(2, 0) = 0.874
        pDistr(3, 0) = 0.897
        pDistr(4, 0) = 0.914
        pDistr(5, 0) = 0.956
        pDistr(6, 0) = 0.978
        pDistr(7, 0) = 0.982
        pDistr(8, 0) = 0.986
        pDistr(9, 0) = 0.988
    ElseIf nSubjs = 19 Then
        pDistr(1, 0) = 0.863
        pDistr(2, 0) = 0.879
        pDistr(3, 0) = 0.901
        pDistr(4, 0) = 0.917
        pDistr(5, 0) = 0.957
        pDistr(6, 0) = 0.978
        pDistr(7, 0) = 0.982
        pDistr(8, 0) = 0.986
        pDistr(9, 0) = 0.988
    ElseIf nSubjs = 20 Then
        pDistr(1, 0) = 0.868
        pDistr(2, 0) = 0.884
        pDistr(3, 0) = 0.905
        pDistr(4, 0) = 0.92
        pDistr(5, 0) = 0.959
        pDistr(6, 0) = 0.979
        pDistr(7, 0) = 0.983
        pDistr(8, 0) = 0.986
        pDistr(9, 0) = 0.988
    ElseIf nSubjs = 21 Then
        pDistr(1, 0) = 0.873
        pDistr(2, 0) = 0.888
        pDistr(3, 0) = 0.908
        pDistr(4, 0) = 0.923
        pDistr(5, 0) = 0.96
        pDistr(6, 0) = 0.98
        pDistr(7, 0) = 0.983
        pDistr(8, 0) = 0.987
        pDistr(9, 0) = 0.989
    ElseIf nSubjs = 22 Then
        pDistr(1, 0) = 0.878
        pDistr(2, 0) = 0.892
        pDistr(3, 0) = 0.911
        pDistr(4, 0) = 0.926
        pDistr(5, 0) = 0.961
        pDistr(6, 0) = 0.98
        pDistr(7, 0) = 0.984
        pDistr(8, 0) = 0.987
        pDistr(9, 0) = 0.989
    ElseIf nSubjs = 23 Then
        pDistr(1, 0) = 0.881
        pDistr(2, 0) = 0.895
        pDistr(3, 0) = 0.914
        pDistr(4, 0) = 0.928
        pDistr(5, 0) = 0.962
        pDistr(6, 0) = 0.981
        pDistr(7, 0) = 0.984
        pDistr(8, 0) = 0.987
        pDistr(9, 0) = 0.989
    ElseIf nSubjs = 24 Then
        pDistr(1, 0) = 0.884
        pDistr(2, 0) = 0.898
        pDistr(3, 0) = 0.916
        pDistr(4, 0) = 0.93
        pDistr(5, 0) = 0.963
        pDistr(6, 0) = 0.981
        pDistr(7, 0) = 0.984
        pDistr(8, 0) = 0.987
        pDistr(9, 0) = 0.989
    ElseIf nSubjs = 25 Then
        pDistr(1, 0) = 0.888
        pDistr(2, 0) = 0.901
        pDistr(3, 0) = 0.918
        pDistr(4, 0) = 0.931
        pDistr(5, 0) = 0.964
        pDistr(6, 0) = 0.981
        pDistr(7, 0) = 0.985
        pDistr(8, 0) = 0.988
        pDistr(9, 0) = 0.989
    ElseIf nSubjs = 26 Then
        ' Update if necessary.
        
    End If
    
    shapiroWilkTable6 = pDistr
End Function
