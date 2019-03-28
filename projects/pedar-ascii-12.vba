' ================================================================================================
' pedar-X ASCII Data Processing
'
' Purpose:
'     Rearrange pedar-X ASCII output to a form which can simulate video playback.
'     View without requiring pedar-X Recorder software.
'
' Requirements:
'     Microsoft® Excel® 2016 MSO (16.0.9029.2106) 32-bit (Microsoft Corp., Redmond, WA, USA).
'
' Usage:
'     Run the appropriate macro on the desired .asc and .fgt files.
'     A spreadsheet of the rearranged data would be generated.
'     Run the playback macro to automatically scroll down spreadsheet, simulating a video.
'
' Limitations:
'     *
'
' Version:
'     2018/08/03
'
' Author:
'     ©2018 Zhuoyuan (Roscoe) Lai, ahnouuyz@gmail.com
' ================================================================================================

' To Do:
'     "Go to % step" button.
'
'     .asc and .fgt pairing validation.
'         Need a way to check if .asc and .fgt files for the same walk are selected.
'         Check that the total time for both are the same.
'
'     Playback speed timing test.
'
'     Graph tracking (difficult!).
'     COP tracking (insane!).


' ================================================================================================
' Globals
' ================================================================================================

' Enable Sleep function.
#If VBA7 Then
    ' For 64 Bit Systems.
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    ' For 32 Bit Systems.
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Toggle conditional formatting type.
Private Const appColorScale As Boolean = True
Private Const appDataBars As Boolean = True

' Block dimensions and positions.
Private Const blockHeight As Long = 19
Private Const blockWidth As Long = 21
Private Const lBlock As Long = 3
Private Const mBlock As Long = lBlock + 7
Private Const rBlock As Long = mBlock + 3

' Force filtering parameters.
Private Const stepWindow As Long = 5 ' Size of scanning window in number of frames (N * 0.02 s)
Private Const stepWidth As Long = stepWindow \ 2

' How did I get 16 N? Knowing the answer beforehand.
' A threshold of 15 N would result in the wrong number of steps.
Private Const stepCrit As Double = 16 ' Critical force threshold in Newtons



Sub createInstructionsAndButtons()
    Dim L2 As Long
    Dim instructionString() As String
    
    Sheets.Add Sheets(1)
    
    instructionString = Split( _
        "Handling ASCII output from pedar-x® recorder" & vbTab & _
        vbNullString & vbTab & _
        "Available exported output from pedar-x® recorder:" & vbTab & _
        "    .asc file:    Raw pressure from the 2 x 99 individual sensors." & vbTab & _
        "    .fgt file:    GFR and COP (x and y coordinates)." & vbTab & _
        "    .mva file:    GRF, peak pressure, and average pressure." & vbTab & _
        "    .lst file:    Matrix of 5 mm x 5 mm cells (it's weird)." & vbTab & _
        "Just focus on the .asc and .fgt files for now." & vbTab & _
        "" & vbTab & _
        "Available macros:" & vbTab & _
        "    Reconstitute .asc and .fgt data to insole matrix configuration." & vbTab & _
        "        (MPP is not available yet)." & vbTab & _
        "    Auto-scroll down rows, simulating video playback." & vbTab & _
        "    (Buttons are still being refined)." & vbTab & _
        "" & vbTab & _
        "Alt+F8 to open macro list." & vbTab & _
        "Alt+F11 to open VBA editor and view source code." & vbTab & _
        "", vbTab)
    
    For L2 = 0 To UBound(instructionString)
        Cells(1 + L2, 1) = instructionString(L2)
    Next L2
    
    callCreateMainButtons
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

Function callCreateNavButtons(Optional FirstPosition = "W3:Z4")
    Const numOfButtons As Long = 5
    
    Dim L2 As Long
    Dim buttonLabels() As String
    Dim macroList() As String
    
    buttonLabels = Split( _
        "," & _
        "Auto-scroll down rows," & _
        "Previous frame," & _
        "Next frame," & _
        "Reposition buttons," & _
        "Delete buttons", ",")
    
    macroList = Split( _
        "," & _
        "callAutoScrollDown," & _
        "prevFrame," & _
        "nextFrame," & _
        "repositionNavButtons," & _
        "deleteButtons,", ",")
    
    For L2 = 1 To numOfButtons
        createButton Range(FirstPosition).Offset((L2 - 1) * 2, 0), buttonLabels(L2), macroList(L2)
    Next L2
End Function



Sub callAutoScrollDown()
    Do
        timeStart = InputBox("Please select start time (in % of walk):", , 0)
        
        If LenB(timeStart) = 0 Then
            MsgBox "No changes made. Exiting script."
            Exit Sub
        ElseIf Not IsNumeric(timeStart) Then
            MsgBox "Please enter a number." '
        ElseIf timeStart < 0 Or timeStart > 100 Then
            MsgBox "Please enter a number between 0 and 100." '
        End If
    Loop While Not IsNumeric(timeStart) Or timeStart < 0 Or timeStart > 100
    
    Do
        timeDelay = InputBox("Please select delay time (in milliseconds):", , 20)
        
        If LenB(timeDelay) = 0 Then
            MsgBox "No changes made. Exiting script."
            Exit Sub
        ElseIf Not IsNumeric(timeDelay) Then
            MsgBox "Please enter a number." '
        ElseIf timeDelay <= 0 Then
            MsgBox "Please enter a positive number." '
        End If
    Loop While Not IsNumeric(timeDelay) Or timeDelay <= 0
    
    autoScrollDown timeStart, timeDelay
End Sub

Function autoScrollDown(timeStart, timeDelay)
' Auto-scroll down sheet by block height, simulating video playback.
'
    Dim L2 As Long
    Dim lastRow As Long
    Dim startRow As Long
    Dim numOfFrames As Long
    
    lastRow = ActiveSheet.UsedRange.Rows.Count - 1
    startRow = (timeStart / 100 * lastRow) - (timeStart / 100 * lastRow) Mod blockHeight
    numOfFrames = (lastRow - startRow) \ blockHeight - 1
    
    Cells(1, 1).Select
    ActiveWindow.SmallScroll Down:=startRow
    For L2 = 1 To numOfFrames
        If L2 Mod (1000 \ timeDelay) = 0 Then DoEvents
        ActiveWindow.SmallScroll Down:=blockHeight
        Sleep timeDelay
    Next L2
End Function

Sub callPrintVideoSheet()
    Do
        insoleType = InputBox( _
            "Please select the insole type:" & vbCr & _
            "  1. Standard version" & vbCr & _
            "  2. Wide version", "Select insole type", 2)
        
        If LenB(insoleType) = 0 Then
            MsgBox "No changes made. Exiting script."
            Exit Sub
        ElseIf insoleType <> 1 And insoleType <> 2 Then
            MsgBox "Please enter a valid number." '
        End If
    Loop While insoleType <> 1 And insoleType <> 2
    
    If insoleType = 2 Then
        printVideoSheet True
    Else
        printVideoSheet False
    End If
End Sub

Function printVideoSheet(wideSole)
' Produce video matrix from .asc (and .fgt) files.
' 3 sec
'
    Dim eSec As Double
    Dim lastRow As Long
    Dim arrAscFgt As Variant
    
    arrAscFgt = returnValidArrs
    If Not IsArray(arrAscFgt) Then Exit Function
    
    eSec = Timer()
    
    printArray videoMatrix(arrAscFgt(1), arrAscFgt(2), wideSole), 0
    
    ' Color.
    lastRow = ActiveSheet.UsedRange.Rows.Count - 1
    
    If appColorScale Then
        absColorScale Range(Cells(1, lBlock + 1), Cells(lastRow, mBlock))
        absColorScale Range(Cells(1, rBlock + 1), Cells(lastRow, rBlock + 7))
    End If
    
    If appDataBars Then
        createDataBarCF Range(Cells(1, lBlock + 1), Cells(lastRow, mBlock))
        createDataBarCF Range(Cells(1, rBlock + 1), Cells(lastRow, rBlock + 7))
    End If
    
    createDataBarCF Cells(1, 1).Resize(lastRow)
    createDataBarCF Cells(1, 3).Resize(lastRow)
    createDataBarCF Cells(1, mBlock + 1).Resize(lastRow, 3)
    createDataBarCF Cells(1, rBlock + 8).Resize(lastRow)
    
'    createDataBarCF Range(Cells(1, 1), Cells(lastRow, 1))
'    createDataBarCF Range(Cells(1, 3), Cells(lastRow, 3))
'    createDataBarCF Range(Cells(1, mBlock + 1), Cells(lastRow, mBlock + 3))
'    createDataBarCF Range(Cells(1, rBlock + 8), Cells(lastRow, rBlock + 8))
    
    resizeCells
    
    callCreateNavButtons
    
    eSec = Timer() - eSec
    Debug.Print "Total time: " & Round(eSec, 5) & " sec"
End Function

Function resizeCells()
    Const timeWidth As Long = 8
    Const forceWidth As Long = 11
    Const insoleWidth As Long = 6
    Const stepCountWidth As Long = 9
    
    Dim L2 As Long
    
    For L2 = 1 To 7
        Columns(lBlock + L2).ColumnWidth = insoleWidth
        Columns(rBlock + L2).ColumnWidth = insoleWidth
    Next L2
    
    Columns(1).ColumnWidth = timeWidth
    Columns(2).ColumnWidth = timeWidth
    Columns(3).ColumnWidth = stepCountWidth
    Columns(mBlock + 1).ColumnWidth = forceWidth
    Columns(mBlock + 2).ColumnWidth = forceWidth
    Columns(mBlock + 3).ColumnWidth = forceWidth
    Columns(rBlock + 8).ColumnWidth = stepCountWidth
    
    Rows.RowHeight = 30 ' About 19 rows visible
    
    With ActiveSheet.UsedRange
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Function

Function videoMatrix(ByRef ascArr, Optional fgtArr = Empty, Optional wideSole = False) As Variant
' The (very) long matrix which will be printed onto a new sheet, for video playback.
' ascArr should be the data table portion of the .asc file.
' fgtArr should be the data table portion of the .fgt file.
'
' Block (19 x 21) Design Layout:
' ========================================================================================
'       01  02  03  04  05  06  07  08  09  10  11  12  13  14  15  16  17  18  19  20  21
'       --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --
' 01:
' 02:
' 03:   Ti          __  99  98  97  96  95  __              __  95  96  97  98  99  __
' 04:   __  __      __  94  93  92  91  90  __              __  90  91  92  93  94  __
' 05:    %   s      89  88  87  86  85  84  83              83  84  85  86  87  88  89
' 06:               82  81  80  79  78  77  76              76  77  78  79  80  81  82
' 07:               75  74  73  72  71  70  69              69  70  71  72  73  74  75
' 08:               68  67  66  65  64  63  62              62  63  64  65  66  67  68
' 09:               61  60  59  58  57  56  55       F      55  56  57  58  59  60  61
' 10:               54  53  52  51  50  49  48  Le  To  Ri  48  49  50  51  52  53  54
' 11:               47  46  45  44  43  42  41  __  __  __  41  42  43  44  45  46  47
' 12:               40  39  38  37  36  35  34   N   N   N  34  35  36  37  38  39  40
' 13:               33  32  31  30  29  28  27              27  28  29  30  31  32  33
' 14:               26  25  24  23  22  21  20              20  21  22  23  24  25  26
' 15:               19  18  17  16  15  14  13              13  14  15  16  17  18  19
' 16:           St  12  11  10  09  08  07  06              06  07  08  09  10  11  12  St
' 17:           __  __  05  04  03  02  01  __              __  01  02  03  04  05  __  __
' 18:
' 19:
' ========================================================================================
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim colInd As Long
    Dim outRows As Long
    Dim scanStart As Long
    Dim blockStart As Long
    Dim maxTime As Double
    Dim lMaxForce As Double
    Dim rMaxForce As Double
    Dim tMaxForce As Double
    Dim lrInsole As Variant
    Dim lrMPPArr As Variant
    Dim lStepCount As Variant
    Dim rStepCount As Variant
    
    colInd = LBound(ascArr, 2) ' Check if array index starts from 0 or 1 (or other)
    
    ' Count steps using force data, and determine maximum force for the walk.
    If Not IsEmpty(fgtArr) Then
        ReDim lForce(LBound(fgtArr) To UBound(fgtArr)) As Double
        ReDim rForce(LBound(fgtArr) To UBound(fgtArr)) As Double
        ReDim tForce(LBound(fgtArr) To UBound(fgtArr)) As Double
        
        For L2 = LBound(fgtArr) To UBound(fgtArr)
            lForce(L2) = fgtArr(L2, colInd + 1)
            rForce(L2) = fgtArr(L2, colInd + 5)
            tForce(L2) = lForce(L2) + rForce(L2)
            
            If lMaxForce < lForce(L2) Then lMaxForce = lForce(L2)
            If rMaxForce < rForce(L2) Then rMaxForce = rForce(L2)
            If tMaxForce < tForce(L2) Then tMaxForce = tForce(L2)
        Next L2
        
        lStepCount = stepCounter(lForce)
        rStepCount = stepCounter(rForce)
    End If
    
    maxTime = ascArr(UBound(ascArr), colInd)
    outRows = (UBound(ascArr) + 1) * blockHeight - (LBound(ascArr) = 0) ' Add 1 if start from 0
    
    ReDim outArr(1 To outRows, 1 To blockWidth) As Variant
    
    scanStart = LBound(ascArr)
    For L2 = scanStart To UBound(ascArr)
        blockStart = 2 + (L2 - scanStart) * blockHeight
        
        ' Left and right insoles, 15 x 7 arrays.
        lrInsole = lrInsReconst(ascArr, L2, wideSole)
        
        For L3 = 1 To 15
            For L4 = 1 To 7
                outArr(L3 + blockStart, L4 + lBlock) = lrInsole(1)(L3, L4)
                outArr(L3 + blockStart, L4 + rBlock) = lrInsole(2)(L3, L4)
            Next L4
        Next L3
        
        ' Time, Force, Step labels.
        outArr(blockStart + 1, 1) = "Time"
        outArr(blockStart + 2, 1) = Round(ascArr(L2, colInd) / maxTime * 100, 2)
        outArr(blockStart + 3, 1) = "%Walk"
        
        outArr(blockStart + 2, 2) = ascArr(L2, colInd)
        outArr(blockStart + 3, 2) = "s"
        
        outArr(blockStart + 7, mBlock + 2) = "GRF"
        outArr(blockStart + 8, mBlock + 1) = "Left"
        outArr(blockStart + 8, mBlock + 2) = "Total"
        outArr(blockStart + 8, mBlock + 3) = "Right"
        
        outArr(blockStart + 10, mBlock + 1) = "N"
        outArr(blockStart + 10, mBlock + 2) = "N"
        outArr(blockStart + 10, mBlock + 3) = "N"
        
        If Not IsEmpty(fgtArr) Then
            If L2 > 3 And L2 < UBound(fgtArr) - 1 Then
                If LenB(lStepCount(L2)) <> 0 Then
                    outArr(blockStart + 14, 3) = "L Step #"
                    outArr(blockStart + 15, 3) = lStepCount(L2)
                End If
                
                If LenB(rStepCount(L2)) <> 0 Then
                    outArr(blockStart + 14, rBlock + 8) = "R Step #"
                    outArr(blockStart + 15, rBlock + 8) = rStepCount(L2)
                End If
            End If
            
            outArr(blockStart + 9, mBlock + 1) = lForce(L2)
            outArr(blockStart + 9, mBlock + 2) = tForce(L2)
            outArr(blockStart + 9, mBlock + 3) = rForce(L2)
        End If
    Next L2
    
    ' MPP.
    blockStart = 2 + (L2 - scanStart) * blockHeight
    
    lrMPPArr = lrMPP(ascArr, wideSole)
    
    For L3 = 1 To 15
        For L4 = 1 To 7
            outArr(L3 + blockStart, L4 + lBlock) = lrMPPArr(1)(L3, L4)
            outArr(L3 + blockStart, L4 + rBlock) = lrMPPArr(2)(L3, L4)
        Next L4
    Next L3
    
    outArr(blockStart + 1, 1) = "MPP"
    
    outArr(blockStart + 7, mBlock + 2) = "MaxForce"
    outArr(blockStart + 8, mBlock + 1) = "Left"
    outArr(blockStart + 8, mBlock + 2) = "Total"
    outArr(blockStart + 8, mBlock + 3) = "Right"
    
    outArr(blockStart + 10, mBlock + 1) = "N"
    outArr(blockStart + 10, mBlock + 2) = "N"
    outArr(blockStart + 10, mBlock + 3) = "N"
    
    If Not IsEmpty(fgtArr) Then
        outArr(blockStart + 9, mBlock + 1) = lMaxForce
        outArr(blockStart + 9, mBlock + 2) = tMaxForce
        outArr(blockStart + 9, mBlock + 3) = rMaxForce
    End If
    
    videoMatrix = outArr
End Function

Function lrMPP(ByRef ascArr, wideSole) As Variant
' Derive the maximum pressure picture data.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim leftMax As Double
    Dim rightMax As Double
    Dim lrMaxArr(1 To 1, 0 To 198) As Double
    
    For L2 = 1 To 99
        leftMax = 0
        rightMax = 0
        
        For L3 = LBound(ascArr) To UBound(ascArr)
            If leftMax < ascArr(L3, LBound(ascArr, 2) + L2) Then
                leftMax = ascArr(L3, LBound(ascArr, 2) + L2)
            End If
            
            If rightMax < ascArr(L3, LBound(ascArr, 2) + 99 + L2) Then
                rightMax = ascArr(L3, LBound(ascArr, 2) + 99 + L2)
            End If
        Next L3
        
        lrMaxArr(1, L2) = leftMax
        lrMaxArr(1, L2 + 99) = rightMax
    Next L2
    
    lrMPP = lrInsReconst(lrMaxArr, 1, wideSole)
End Function

Function lrInsReconst(ByRef ascArr, refRow, wideSole) As Variant
' Rearrange numbers from (1 row of 99) x 2 (1D) to (15 rows of 7) x 2 (2D):
'
' Standard left insole               |   Standard right insole
' ================================   |   ================================
'       01  02  03  04  05  06  07   |         01  02  03  04  05  06  07
'       --  --  --  --  --  --  --   |         --  --  --  --  --  --  --
' 01:   __  __  99  98  97  96  __   |   01:   __  96  97  98  99  __  __
' 02:   __  95  94  93  92  91  90   |   02:   90  91  92  93  94  95  __
' 03:   89  88  87  86  85  84  83   |   03:   83  84  85  86  87  88  89
' 04:   82  81  80  79  78  77  76   |   04:   76  77  78  79  80  81  82
' 05:   75  74  73  72  71  70  69   |   05:   69  70  71  72  73  74  75
' 06:   68  67  66  65  64  63  62   |   06:   62  63  64  65  66  67  68
' 07:   61  60  59  58  57  56  55   |   07:   55  56  57  58  59  60  61
' 08:   54  53  52  51  50  49  48   |   08:   48  49  50  51  52  53  54
' 09:   47  46  45  44  43  42  41   |   09:   41  42  43  44  45  46  47
' 10:   40  39  38  37  36  35  34   |   10:   34  35  36  37  38  39  40
' 11:   33  32  31  30  29  28  27   |   11:   27  28  29  30  31  32  33
' 12:   26  25  24  23  22  21  20   |   12:   20  21  22  23  24  25  26
' 13:   19  18  17  16  15  14  13   |   13:   13  14  15  16  17  18  19
' 14:   12  11  10  09  08  07  06   |   14:   06  07  08  09  10  11  12
' 15:   __  05  04  03  02  01  __   |   15:   __  01  02  03  04  05  __
' ================================   |   ================================
'
' Wide left insole                   |   Wide right insole
' ================================   |   ================================
'       01  02  03  04  05  06  07   |         01  02  03  04  05  06  07
'       --  --  --  --  --  --  --   |         --  --  --  --  --  --  --
' 01:   __  99  98  97  96  95  __   |   01:   __  95  96  97  98  99  __
' 02:   __  94  93  92  91  90  __   |   02:   __  90  91  92  93  94  __
' 03:   89  88  87  86  85  84  83   |   03:   83  84  85  86  87  88  89
' 04:   82  81  80  79  78  77  76   |   04:   76  77  78  79  80  81  82
' 05:   75  74  73  72  71  70  69   |   05:   69  70  71  72  73  74  75
' 06:   68  67  66  65  64  63  62   |   06:   62  63  64  65  66  67  68
' 07:   61  60  59  58  57  56  55   |   07:   55  56  57  58  59  60  61
' 08:   54  53  52  51  50  49  48   |   08:   48  49  50  51  52  53  54
' 09:   47  46  45  44  43  42  41   |   09:   41  42  43  44  45  46  47
' 10:   40  39  38  37  36  35  34   |   10:   34  35  36  37  38  39  40
' 11:   33  32  31  30  29  28  27   |   11:   27  28  29  30  31  32  33
' 12:   26  25  24  23  22  21  20   |   12:   20  21  22  23  24  25  26
' 13:   19  18  17  16  15  14  13   |   13:   13  14  15  16  17  18  19
' 14:   12  11  10  09  08  07  06   |   14:   06  07  08  09  10  11  12
' 15:   __  05  04  03  02  01  __   |   15:   __  01  02  03  04  05  __
' ================================   |   ================================
'
' Sensor numbers increase from posterior/proximal to anterior/distal, medial to lateral.
' But array indices start from the top-left and advances towards the bottom-right.
' Quite the opposite, and rather troublesome.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim lCount As Long
    Dim lArr(1 To 15, 1 To 7) As Variant ' Unassigned cells will be empty, not 0
    Dim rArr(1 To 15, 1 To 7) As Variant ' Unassigned cells will be empty, not 0
    Dim outArr(1 To 2) As Variant
    
    lCount = LBound(ascArr, 2) ' Must have one column before data (e.g. time column)
    
    For L2 = 15 To 3 Step -1
        If L2 <> 15 Then
            For L3 = 1 To 7
                lCount = lCount + 1
                lArr(L2, 8 - L3) = ascArr(refRow, lCount)
                rArr(L2, L3) = ascArr(refRow, lCount + 99)
            Next L3
        ElseIf L2 = 15 Then
            For L3 = 2 To 6
                lCount = lCount + 1
                lArr(L2, 8 - L3) = ascArr(refRow, lCount)
                rArr(L2, L3) = ascArr(refRow, lCount + 99)
            Next L3
        End If
    Next L2
    
    For L2 = 2 To 1 Step -1
        If wideSole Then
            For L3 = 2 To 6
                lCount = lCount + 1
                lArr(L2, 8 - L3) = ascArr(refRow, lCount)
                rArr(L2, L3) = ascArr(refRow, lCount + 99)
            Next L3
        Else
            If L2 = 2 Then
                For L3 = 1 To 6
                    lCount = lCount + 1
                    lArr(L2, 8 - L3) = ascArr(refRow, lCount)
                    rArr(L2, L3) = ascArr(refRow, lCount + 99)
                Next L3
            ElseIf L2 = 1 Then
                For L3 = 2 To 5
                    lCount = lCount + 1
                    lArr(L2, 8 - L3) = ascArr(refRow, lCount)
                    rArr(L2, L3) = ascArr(refRow, lCount + 99)
                Next L3
            End If
        End If
    Next L2
    
    outArr(1) = lArr
    outArr(2) = rArr
    lrInsReconst = outArr
End Function

' ================================================================================================
' Common Functions/Subroutines
' ================================================================================================

Function returnValidArrs() As Variant
' Returns [ascArr1, (fgtArr1 or Empty)].
'
    Dim L2 As Long
    Dim iCount As Long
    Dim ascCount As Long
    Dim fgtCount As Long
    Dim iVar As Variant
    Dim ascArr1 As Variant
    Dim fgtArr1 As Variant
    Dim outArr(1 To 2) As Variant
    
    MsgBox _
        "Please select 2 files:" & vbCr & _
        "  1.  One .asc file" & vbCr & _
        "  2.  The cognate .fgt file" & vbCr & vbCr & _
        "Please do not select more than one .asc file." & vbCr & vbCr & _
        "Please do not select more than two files.", , "Instructions"
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        If .Show = -1 Then
            For Each iVar In .SelectedItems
                ' Check number of files.
                iCount = iCount + 1
                If iCount > 2 Then
                    MsgBox "I told you not to select more than two files!"
                    Exit Function
                End If
                
                ' Get valid file paths.
                If Right$(iVar, 4) = ".asc" Then
                    ascCount = ascCount + 1
                    If ascCount > 1 Then
                        MsgBox "I told you not to select more than one .asc file!"
                        Exit Function
                    End If
                    
                    ascArr1 = readAscFgtFile(iVar)
                ElseIf Right$(iVar, 4) = ".fgt" Then
                    fgtCount = fgtCount + 1
                    fgtArr1 = readAscFgtFile(iVar)
                End If
            Next
        Else
            MsgBox "No files selected. Exiting script."
            Exit Function
        End If
    End With
    
    If ascCount = 1 Then
        outArr(1) = ascArr1
        
        If fgtCount = 1 Then
            outArr(2) = fgtArr1
        Else
            outArr(2) = Empty
        End If
    Else
        MsgBox "No .asc files detected!"
        Exit Function
    End If
    
    returnValidArrs = outArr
End Function

Function readAscFgtFile(filePath) As Variant
    Dim L2 As Long
    Dim dataArr() As Variant
    
    Workbooks.Open filePath, , True
    
    ' Trim to data table values only (no headers).
    Do
        L2 = L2 + 1
    Loop While Left$(Cells(L2, 1).Value2, 4) <> "time"
    Do
        L2 = L2 + 1
    Loop While LenB(Cells(L2, 1).Value2) = 0
'    dataArr = Range(Cells(L2, 1), ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell)).Value2
    dataArr = ActiveSheet.UsedRange.Value2
    ActiveWorkbook.Close False
    
    readAscFgtFile = dataArr
End Function

Function stepCounter(ByRef forceArr) As Variant
' Receives an array of forces [index 1 to N].
' Returns an array of integers [index 4 to N-2].
'
    Dim L2 As Long
    Dim sCount As Long
    Dim isStep As Boolean
    
    ReDim outArr(LBound(forceArr) + 1 + stepWidth To UBound(forceArr) - stepWidth) As Variant
    
    For L2 = LBound(forceArr) + 1 + stepWidth To UBound(forceArr) - stepWidth
        If isHeelStrike(forceArr, L2) Then
            sCount = sCount + 1
            isStep = True
        ElseIf isToeOff(forceArr, L2) Then
            isStep = False
        End If
        
        If isStep Then outArr(L2) = sCount
    Next L2
    
    stepCounter = outArr
End Function

Function isHeelStrike(ByRef forceArr, forceInd) As Boolean
' Determine if heel contact occured at the current frame.
'
    If Not isStance(forceArr, forceInd - 1) And isStance(forceArr, forceInd) Then
        isHeelStrike = True
    End If
End Function

Function isToeOff(ByRef forceArr, forceInd) As Boolean
' Determine if toe off occured at the current frame.
'
    If isStance(forceArr, forceInd - 1) And Not isStance(forceArr, forceInd) Then
        isToeOff = True
    End If
End Function

Function isStance(ByRef forceArr, forceInd) As Boolean
' Determine if the foot is on the ground.
'
    Dim L2 As Long
    Dim localForce As Double
    
    For L2 = forceInd - stepWidth To forceInd + stepWidth
        If localForce < forceArr(L2) Then localForce = forceArr(L2)
    Next L2
    
    If localForce > stepCrit Then isStance = True
End Function

Function absColorScale(iRng As Range)
' Color scale based on absolute pressure values.
'
' The smallest reported difference in pressure values is 2.5 kPa.
' Index an array such that each index corresponds to an increment of 2.5 kPa.
'   1   ~~> 2.5 kPa
'   2   ~~> 5.0 kPa
'   3   ~~> 7.5 kPa
'   n   ~~> (2.5 * n) kPa
'   80  ~~> 200.0 kPa
'   81  ~~> 202.5 kPa
'   240 ~~> 600.0 kPa
' Break into 5 groups:
'   [01-20]  ~~> [2.5-50 kPa] blue to cyan
'   [21-40]  ~~> [52.5-100 kPa] cyan to green
'   [41-60]  ~~> [102.5-150 kPa] green to yellow
'   [61-80]  ~~> [152.5-200 kPa] yellow to red
'   [81-240] ~~> [202.5-600 kPa] red to rink
'
' It's not an even distribution, but I really want 200 kPa to be colored red.
'
' Limit array to index 240 (600 kPa).
' Any higher value would be colored pink.
'
    Dim L2 As Long
    Dim iCell As Range
    Dim bcgyrpArr(1 To 240) As Long
    
    ' Define the overall color array.
    For L2 = 1 To 20
        bcgyrpArr(L2) = RGB(0, CLng(255 * L2 / 20), 255) ' Blue to cyan 01-20
        bcgyrpArr(20 + L2) = RGB(0, 255, CLng(255 * (20 - L2) / 20)) ' Cyan to green 21-40
        bcgyrpArr(40 + L2) = RGB(CLng(255 * L2 / 20), 255, 0) ' Green to yellow 41-60
        bcgyrpArr(60 + L2) = RGB(255, CLng(255 * (20 - L2) / 20), 0) ' Yellow to red 61-80
        bcgyrpArr(80 + L2) = RGB(255, 0, CLng(255 * L2 / 160)) ' Red to pink 81-100
    Next L2
    
    For L2 = 21 To 160
        bcgyrpArr(80 + L2) = RGB(255, 0, CLng(255 * L2 / 160)) ' Red to pink 101-240
    Next L2
    
    For Each iCell In iRng
        With iCell
            If LenB(.Value2) <> 0 Then
                If .Value2 > 0 Then
                    If .Value2 <= 600 Then
                        .Interior.Color = bcgyrpArr(.Value2 / 2.5)
                    Else
                        .Interior.Color = 16711935 ' Pink
                    End If
                End If
            End If
        End With
    Next
End Function

Function createDataBarCF(iRng As Range)
' Assuming no negative values.
'
    Dim cfDataBar As Databar
    Set cfDataBar = iRng.FormatConditions.AddDatabar
    
    With cfDataBar
        .BarColor.Color = 5920255 ' Red
        .BarFillType = xlDataBarFillGradient
        .BarBorder.Type = xlDataBarBorderSolid
        .BarBorder.Color.Color = 5920255 ' Red
        
'        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
'        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
'        .Direction = xlContext
'        .NegativeBarFormat.ColorType = xlDataBarColor
'        .NegativeBarFormat.BorderColorType = xlDataBarColor
'        .AxisPosition = xlDataBarAxisAutomatic
'        .AxisColor.Color = 0
'        .NegativeBarFormat.Color.Color = 255
'        .NegativeBarFormat.BorderColor.Color = 255
    End With
End Function

Function printArray(ByRef inputArr, Optional bookOrSheet As Long = 0)
' Prints a 2D array on a new workbook or worksheet.
'
    If bookOrSheet = 0 Then Workbooks.Add
    If bookOrSheet = 1 Then Sheets.Add , ActiveSheet
    Cells(1, 1).Resize(UBound(inputArr), UBound(inputArr, 2)).Value2 = inputArr
'    Range(Cells(1, 1), Cells(1 + UBound(inputArr), 1 + UBound(inputArr, 2))).Value2 = inputArr
End Function

Function callCreateMainButtons(Optional FirstPosition = "L4:O5")
    Const numOfButtons As Long = 3
    
    Dim L2 As Long
    Dim buttonLabels() As String
    Dim macroList() As String
    
    buttonLabels = Split( _
        "," & _
        "Input .asc and .fgt files," & _
        "Reposition buttons," & _
        "Delete buttons", ",")
    
    macroList = Split( _
        "," & _
        "callPrintVideoSheet," & _
        "repositionMainButtons," & _
        "deleteButtons,", ",")
    
    For L2 = 1 To numOfButtons
        createButton Range(FirstPosition).Offset((L2 - 1) * 2, 0), buttonLabels(L2), macroList(L2)
    Next L2
End Function

Sub createMainButtons()
    callCreateMainButtons
End Sub

Sub createNavButtons()
    callCreateNavButtons
End Sub

Function repositionMainButtons()
    ActiveSheet.Buttons.Delete
    callCreateMainButtons ActiveCell.Row, ActiveCell.Column
End Function

Function repositionNavButtons()
    ActiveSheet.Buttons.Delete
    callCreateNavButtons ActiveCell.Row, ActiveCell.Column
End Function

Function prevFrame()
    ActiveWindow.SmallScroll Down:=-blockHeight
    ActiveSheet.Buttons.Delete
    ActiveCell.Offset(-blockHeight, 0).Select
    callCreateNavButtons ActiveCell.Row, ActiveCell.Column
End Function

Function nextFrame()
    ActiveWindow.SmallScroll Down:=blockHeight
    ActiveSheet.Buttons.Delete
    ActiveCell.Offset(blockHeight, 0).Select
    callCreateNavButtons ActiveCell.Row, ActiveCell.Column
End Function

Sub deleteButtons()
    ActiveSheet.Buttons.Delete
End Sub

' ================================================================================================
' Unused functions
' - Reading text files turns out to be slower
' - Might have to do with large number of columns (?)
' ================================================================================================

Function convertFile(filePath) As String
' Convert .asc or .fgt files to tab-delimited text format.
' Returns the new file path.
'
    Dim fileExt As String
    Dim fileNewPath As String
    
    fileExt = Right(filePath, 3)
    fileNewPath = Left(filePath, Len(filePath) - 4) & fileExt & ".txt"
    
    Workbooks.Open filePath, , True
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileNewPath, xlText
    ActiveWorkbook.Close False
    Application.DisplayAlerts = True
    
    convertFile = fileNewPath
End Function

Function readTextFile(filePath) As String()
' Reads each line of text file into a 1D array of strings.
'
    Dim fileNum As Long
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
        readTextFile = Split(Input$(LOF(fileNum), #fileNum), vbNewLine)
    Close #fileNum
End Function

Function linesToTable(ByRef inArr) As Variant
' Converts a 1D array of strings with tab-delimited values to 2D table form.
'
    Dim L2 As Long
    Dim L3 As Long
    ReDim outArr(0 To UBound(inArr), 0 To UBound(Split(inArr(0), vbTab))) As Variant
    
    For L2 = 0 To UBound(inArr)
        For L3 = 0 To UBound(Split(inArr(0), vbTab))
            If LenB(inArr(L2)) <> 0 Then
                outArr(L2, L3) = Split(inArr(L2), vbTab)(L3)
            End If
        Next L3
    Next L2
    
    linesToTable = outArr
End Function

Function numToLetter(num As Long) As String
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

Function filterForce(ByRef forceArr) As Variant
' Filter out (set to 0) force values during swing phase.
' Use rules to determine if a foot is on the ground.
' WARNING: This will edit data!
' With a working step count function available, the utility of this is uncertain.
'
    Dim L2 As Long
    Dim isStep As Boolean
    
    ReDim outArr(LBound(forceArr) + 4 To UBound(forceArr) - 2) As Variant
    
    For L2 = LBound(forceArr) + 4 To UBound(forceArr) - 2
        If isHeelStrike(forceArr, L2) Then
            isStep = True
        ElseIf isToeOff(forceArr, L2) Then
            isStep = False
        End If
        
        If isStep Then
            outArr(L2) = forceArr(L2)
        Else
            outArr(L2) = 0
        End If
    Next L2
    
    filterForce = outArr
End Function
