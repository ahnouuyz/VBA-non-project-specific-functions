' Ongoing:
' - Rename Subs and Functions, reorganize code blocks.
' - Minimize source code.
'
' Future plans:
' - Function to highlight potential outliers.
'

Private PrintReplicates As Boolean
Private PrintMean95CI As Boolean
Private PlotChart As Boolean

' ================================================================================================
' Phase 01: Create PlateMap from raw data file.
' ================================================================================================

Sub A01_RawDataAndPlateMap()
    Dim RawFilePath As String
    Dim RawStringArr() As String
    Dim RawArrArr() As Variant
    Dim RawSheet2D() As Variant
    Dim PlateArr(0) As Variant
    
    MsgBox "Please select the qPCR Ct data (.txt) file."
    RawFilePath = GetFilePath
    
    ' Check 1: If no file is selected, exit.
    If LenB(RawFilePath) = 0 Then Exit Sub
    
    ' Check 2: If not text file, try to read using Excel.
    If Right(RawFilePath, 4) <> ".txt" Then
        Dim wb As Workbook
        Set wb = Workbooks.Open(RawFilePath)
        RawSheet2D = ReindexArr2D(ActiveSheet.UsedRange.Value2)
        wb.Close False
    Else
        RawStringArr = readTextFile(RawFilePath) ' Read text file into a String()
        RawArrArr = TextArrToArrArr(RawStringArr, vbTab) ' Convert String() into an array of arrays
        RawSheet2D = ArrArrToArr2D(RawArrArr) ' Convert array of arrays into a 2D array
    End If
    
    ' Construct PlateMap.
    ' Store it as the 1st entry in an array of arrays (because... reasons in later part of code).
    PlateArr(0) = PlateMap384_96(RawSheet2D)
    
    ' Print raw data.
    Workbooks.Add
    PrintArr2D RawSheet2D, , , "Raw"
    
    ' Print plate map.
    DisplayPlateMap PlateArr, 5
End Sub

Function DisplayPlateMap(PlateArr, Optional OutRow = 5)
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim ColorList() As Long
    Dim GenesList() As Variant
    
    GenesList = GetConditionsGenesList(PlateArr, 2) ' Get list of genes.
    ColorList = GetRainbowColorList(UBound(GenesList)) ' Get list of colors.
    
    PrintReplicates = True
    
    Sheets.Add , Sheets(Sheets.Count)
    PrintArr2D PlateArr(0), OutRow, , "PlateMap"
    createButtons2
    createButtons3
    
    ' Decorate PlateMap.
    Application.DisplayAlerts = False
    For L2 = 1 + OutRow To UBound(PlateArr(0)) + OutRow Step 3
        Cells(L2, 1).Resize(3).Merge
        For L3 = 2 To UBound(PlateArr(0), 2) + 1
            Cells(L2, L3).Resize(3).BorderAround xlContinuous
            For L4 = 1 To UBound(GenesList)
                With Cells(L2 + 1, L3)
                    If .Value2 = GenesList(L4) Then .Interior.color = ColorList(L4)
                End With
            Next L4
        Next L3
    Next L2
    Application.DisplayAlerts = True
    
    Cells(OutRow + 1, 2).Select
    ActiveWindow.FreezePanes = True
    
    With ActiveSheet.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Select
    End With
    ActiveWindow.Zoom = True
    
    ' Deselect.
    Cells.Select
    Application.CutCopyMode = False
    Cells(1, 1).Select
End Function

Function PlateMap384_96(RawSheet2D) As Variant
' 384 well: 16 x 24 ~~> 48 x 24 array.
' 96 well: 8 x 12 ~~> 24 x 12 array.
' Each well corresponds to a triple row of (condition, gene, Ct value).
'
    Dim L2 As Long
    Dim r1 As Long
    Dim c1 As Long
    Dim rN As Long
    Dim cN As Long
    Dim RawCtCol As Long
    Dim Is384 As Boolean
    
    ' Find the column with raw Ct values.
    For r1 = 0 To UBound(RawSheet2D)
        If RawSheet2D(r1, 0) Like "*Well*" Then Exit For
    Next r1
    
    For c1 = 0 To UBound(RawSheet2D, 2)
        If Len(RawSheet2D(r1, c1)) = 2 Then ' "C_" or "C?" or "Ct"
            If Left(RawSheet2D(r1, c1), 1) = "C" Then
                RawCtCol = c1
                Exit For
            End If
        End If
    Next c1
    If RawCtCol = 0 Then MsgBox "Unable to find raw Ct column."
    
    ' Check if plate is 384-well or 96-well.
    ' Assume 96-well until evidence to the contrary is found.
    For L2 = 0 To UBound(RawSheet2D)
        If LenB(RawSheet2D(L2, 0)) Then
            If IsNumeric(RawSheet2D(L2, 0)) Then
                If RawSheet2D(L2, 0) > 96 Then
                    Is384 = True
                    Exit For
                End If
            Else
                If IsNumeric(Right(RawSheet2D(L2, 0), Len(RawSheet2D(L2, 0)) - 1)) Then
                    If UCase(Left(RawSheet2D(L2, 0), 1)) > "H" Then
                        Is384 = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next L2
    
    rN = 24 - 24 * (Is384 = True)
    cN = 12 - 12 * (Is384 = True)
    
    ReDim PlateArr(0 To rN, 0 To cN) As Variant
    
    ' Numbers along the margins.
    For L2 = 1 To rN Step 3
        PlateArr(L2, 0) = L2 \ 3 + 1
    Next L2
    
    For L2 = 1 To cN
        PlateArr(0, L2) = L2
    Next L2
    
    ' Arrange data based on well number.
    For L2 = 0 To UBound(RawSheet2D)
        If LenB(RawSheet2D(L2, 0)) <> 0 Then
            If IsNumeric(RawSheet2D(L2, 0)) Then
                r1 = ((RawSheet2D(L2, 0) - 1) \ cN) * 3 + 1
                c1 = (RawSheet2D(L2, 0) - 1) Mod cN + 1
            ElseIf IsNumeric(Right(RawSheet2D(L2, 0), Len(RawSheet2D(L2, 0)) - 1)) Then
                r1 = (StrToNum(Left(RawSheet2D(L2, 0), 1)) - 1) * 3 + 1
                c1 = CLng(Right(RawSheet2D(L2, 0), Len(RawSheet2D(L2, 0)) - 1))
            Else
                r1 = 0
                c1 = 0
            End If
            
            If r1 <> 0 And c1 <> 0 Then
                PlateArr(r1, c1) = RawSheet2D(L2, 1) ' Condition
                PlateArr(r1 + 1, c1) = RawSheet2D(L2, 2) ' Gene
                PlateArr(r1 + 2, c1) = RawSheet2D(L2, RawCtCol) ' Raw Ct value
            End If
        End If
    Next L2
    
    PlateMap384_96 = PlateArr
End Function

Function MakeLiteral(val) As Variant
    If LenB(val) Then ' Leave blanks alone
        If Not IsNumeric(val) Then ' Do not make numbers literal
            If Left(val, 1) <> "=" Then ' Do not make equations literal
                If Left(val, 1) <> "'" Then ' Do not add more than one apostrophe
                    MakeLiteral = "'" & val ' Append an apostrophe to the left
                    Exit Function
                End If
            End If
        End If
    End If
    MakeLiteral = val ' Return the original value, unmodified
End Function

Function MakeLiteralArr2D(Arr2D) As Variant()
    Dim L2 As Long
    Dim L3 As Long
    
    ReDim NewArr(LBound(Arr2D) To UBound(Arr2D), LBound(Arr2D, 2) To UBound(Arr2D, 2)) As Variant
    
    For L2 = LBound(Arr2D) To UBound(Arr2D)
        For L3 = LBound(Arr2D, 2) To UBound(Arr2D, 2)
            NewArr(L2, L3) = MakeLiteral(Arr2D(L2, L3))
        Next L3
    Next L2
    
    MakeLiteralArr2D = NewArr
End Function

Function GetConditionsGenesList(ArrArr, RowIndex) As Variant()
' 1 ~~> Conditions
' 2 ~~> Genes
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim TitlesArr() As String
    ReDim ListArr(0) As Variant
    
    TitlesArr = Split(",Conditions,Genes", ",")
    ListArr(0) = TitlesArr(RowIndex)
    
    For L4 = LBound(ArrArr) To UBound(ArrArr)
        For L2 = 1 To UBound(ArrArr(L4)) Step 3
            For L3 = 1 To UBound(ArrArr(L4), 2)
                ' Must have a value for both condition and gene.
                If LenB(ArrArr(L4)(L2, L3)) <> 0 And LenB(ArrArr(L4)(L2 + 1, L3)) <> 0 Then
                    If Not InList(ListArr, ArrArr(L4)(L2 - 1 + RowIndex, L3)) Then
                        ReDim Preserve ListArr(0 To UBound(ListArr) + 1) As Variant
                        ListArr(UBound(ListArr)) = ArrArr(L4)(L2 - 1 + RowIndex, L3)
                    End If
                End If
            Next L3
        Next L2
    Next L4
    
    GetConditionsGenesList = ListArr
End Function

' ================================================================================================
' Phase 02: Create calculations sheet.
' ================================================================================================

Sub A02_CalculationTables()
' Part 2: Run after selecting which wells to analyze.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim SubConditionsString As String
    Dim SubGenesString As String
    Dim SubArr() As Variant
    Dim SubConditionsList() As Variant
    Dim SubGenesList() As Variant
    Dim SortedArr() As Variant
    Dim MainTable() As Variant
    Dim ForPrism() As Variant
    
    ' Make rectangular selections on the plate map.
    ReDim SubArr(1 To Selection.Areas.Count) As Variant
    
    For L2 = 1 To Selection.Areas.Count
        SubArr(L2) = Selection.Areas(L2).Value2
        If Not PassArrayCheck(SubArr(L2)) Then Exit Sub
    Next L2
    
    SubConditionsList = GetConditionsGenesList(SubArr, 1)
    SubGenesList = GetConditionsGenesList(SubArr, 2)
    
    For L2 = LBound(SubConditionsList) + 1 To UBound(SubConditionsList)
        SubConditionsString = SubConditionsString & SubConditionsList(L2) & ","
    Next L2
    For L2 = LBound(SubGenesList) + 1 To UBound(SubGenesList)
        SubGenesString = SubGenesString & SubGenesList(L2) & ","
    Next L2
    
    SortedArr = GetSortedArr(SubArr, SubGenesList, SubConditionsList)
    MaxRep = UBound(SortedArr, 2) - 1
    MainTable = ProcessCtData(SortedArr, MaxRep)
    
    Sheets.Add , Sheets(Sheets.Count)
    PrintArr2D MainTable, , , NameSheet(2)
    
    ' Drop-down lists.
    CreateDropdownList Cells(1, 2), SubGenesString
    CreateDropdownList Cells(2, 2), SubConditionsString
    Cells(1, 4).Resize(2).Value2 = "=IF(ISBLANK(RC[-2]),"" <<< Please select from the drop-down list"","""")"
    
    Columns(3).Hidden = True
    Cells(1, 6 + MaxRep).Resize(, 10).EntireColumn.Hidden = True
    DecorateSheet1 SubConditionsList, SubGenesList, MaxRep
    Range("A:B").NumberFormat = "@" ' Text format
    
    ' Tables for PRISM.
    ForPrism = PivotForPRISM(MainTable, SubConditionsList, SubGenesList, MaxRep)
    
    If PrintReplicates Then
        Sheets.Add , Sheets(Sheets.Count)
        PrintArr2D ForPrism(1), , , NameSheet(3)
        DecorateSheet2 MaxRep
        If PlotChart Then PlotClusterColumnsMeanSD
    End If
    
    If PrintMean95CI Then
        Sheets.Add , Sheets(Sheets.Count)
        PrintArr2D ForPrism(2), , , NameSheet(4)
        DecorateSheet2
        If PlotChart Then PlotClusterColumnsMean95CI
    End If
    
    ' Select the latest Workings sheet.
    For L2 = Sheets.Count To 1 Step -1
        If Left(Sheets(L2).Name, 8) = "Workings" Then
            Sheets(L2).Select
            Cells(1, 2).Select
            Exit For
        End If
    Next L2
End Sub

Function PassArrayCheck(SelectArr) As Boolean
' Check if a selection is valid.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim r3 As Long
    Dim rN As Long
    Dim NotEmpty As Boolean
    
    ' Check 0: Selection is an array.
    If Not IsArray(SelectArr) Then
        MsgBox "Error: Please select at least one rectangular range of cells.", , "Invalid Selection"
        Exit Function
    End If
    
    r1 = LBound(SelectArr) ' Should be a condition
    r2 = r1 + 1 ' Should be a gene
    r3 = r2 + 1 ' Should be a number or "Undetermined"
    rN = UBound(SelectArr) - r1 + 1 ' Should be a number or "Undetermined"
    
    ' Check 1: Number of rows is a multple of 3.
    If rN Mod 3 <> 0 Then
        MsgBox "Error: Number of rows selected should be a multiple of 3. Please try again.", , "Invalid Selection"
        Exit Function
    End If
    
    ' Check 2: Selection is not frame-shifted.
    ' Only checks one (the 1st) well.
    If Not (IsNumeric(SelectArr(r3, 1)) Or SelectArr(r3, 1) = "Undetermined") Then
        MsgBox "Error: Selection might be frame-shifted. Please try again.", , "Invalid Selection"
        Exit Function
    End If
    
    ' Check 4: Selection is not completely empty.
    For L2 = LBound(SelectArr) To UBound(SelectArr)
        For L3 = LBound(SelectArr, 2) To UBound(SelectArr, 2)
            If LenB(SelectArr(L2, L3)) Then
                NotEmpty = True
                Exit For
            End If
        Next L3
        If NotEmpty Then Exit For
    Next L2
    
    If Not NotEmpty Then
        MsgBox "Error: Selection is empty. Please try again.", , "Invalid Selection"
        Exit Function
    End If
    
    PassArrayCheck = True
End Function

Function GetSortedArr(SelectArrArr, SubGenesList, SubConditionsList)
' This function is very important.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim L5 As Long
    Dim L7 As Long
    Dim MaxRep As Long
    Dim RepCount As Long
    
    ReDim SortedArr(0) As Variant
    
    ' Find all replicates for a gene-condition combination.
    For L2 = 1 To UBound(SubGenesList)
        For L3 = 1 To UBound(SubConditionsList)
            ReDim SortedRow(0) As Variant
            RepCount = 0
            
            For L7 = LBound(SelectArrArr) To UBound(SelectArrArr)
                For L5 = 1 To UBound(SelectArrArr(L7), 2)
                    For L4 = 1 To UBound(SelectArrArr(L7)) Step 3 ' L4 = first row of well (Condition)
                        If SelectArrArr(L7)(L4 + 1, L5) = SubGenesList(L2) Then
                            If SelectArrArr(L7)(L4, L5) = SubConditionsList(L3) Then
                                RepCount = RepCount + 1

                                ReDim Preserve SortedRow(0 To 1 + RepCount) As Variant
                                SortedRow(1 + RepCount) = SelectArrArr(L7)(L4 + 2, L5)
                            End If
                        End If
                    Next L4
                Next L5
            Next L7
            
            If RepCount Then
                If MaxRep < RepCount Then MaxRep = RepCount
                SortedRow(0) = SubConditionsList(L3)
                SortedRow(1) = SubGenesList(L2)
                
                ReDim Preserve SortedArr(0 To UBound(SortedArr) + 1) As Variant
                SortedArr(UBound(SortedArr)) = SortedRow
            End If
        Next L3
    Next L2
    
    ' Header row.
    ReDim SortedRow(0 To 1 + MaxRep) As Variant
    SortedRow(0) = "Condition"
    SortedRow(1) = "Gene"
    For L2 = 1 To MaxRep
        SortedRow(L2 + 1) = "Ct " & L2
    Next L2
    SortedArr(0) = SortedRow
    
    GetSortedArr = ArrArrToArr2D(SortedArr)
End Function

Function CreateDropdownList(CellRange, ListFormula)
    With CellRange
        .Interior.color = RGB(255, 235, 179)
        With .Validation
            .Delete
            .Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:=ListFormula
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    End With
End Function

Function NameSheet(SheetNumber) As String
' Allows multiple analyses from the same PlateMap without naming issues.
'
    Dim L2 As Long
    Dim MaxNum As Long
    Dim SheetNames As Variant
    
    SheetNames = Array(, , "Workings", "Replicates", "Mean95CI")
    
    For L2 = 1 To Sheets.Count
        If SheetNames(SheetNumber) = Left(Sheets(L2).Name, Len(Sheets(L2).Name) - 2) Then
            If MaxNum < Right(Sheets(L2).Name, 1) Then
                MaxNum = Right(Sheets(L2).Name, 1)
            End If
        End If
    Next L2
    
    NameSheet = SheetNames(SheetNumber) & "_" & MaxNum + 1
End Function

Function ProcessCtData(SortedArr, MaxRep) As Variant
' Create formulae on the Workings sheet.
'
    Const alpha As Double = 0.05 ' Significance level
    
    Dim L2 As Long
    Dim L3 As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim c3 As Long
    Dim c4 As Long
    Dim cN As Long
    Dim LookupCol As String
    Dim LookupArr As String
    Dim IndexMatchStr1 As String
    Dim IndexMatchStr2 As String
    Dim Row_RawCt As String
    Dim Row_ddCt As String
    Dim NotBlankCond As String
    Dim IncreaseCond As String
    Dim DecreaseCond As String
    
    c1 = 3 ' Raw Ct 1
    c2 = c1 + MaxRep ' CV
    c3 = c2 + 12 ' RQ 1
    c4 = c3 + 2 + MaxRep ' RQ Mean
    cN = c4 + 3 + 1
    
    Dim StatsArr() As Variant
    ReDim StatsArr(0 To UBound(SortedArr) + 3, 0 To cN) As Variant
    
    ' Header rows.
    StatsArr(0, 0) = "Endogenous Control"
    StatsArr(1, 0) = "Reference Sample"
    
    StatsArr(3, 0) = "Condition"
    StatsArr(3, 1) = "Gene"
    StatsArr(3, 2) = "ConditionGene" ' Hidden
    For L2 = 1 To MaxRep
        StatsArr(3, c1 - 1 + L2) = "Raw Ct " & L2
    Next L2
    StatsArr(3, c2) = "Mean"
    StatsArr(3, c2 + 1) = "CV (%)"
    StatsArr(3, c2 + 2) = "SD" ' Hidden
    StatsArr(3, c2 + 3) = "Count" ' Hidden
    StatsArr(3, c2 + 4) = "=R1C2" & "&"" Mean""" ' Hidden
    StatsArr(3, c2 + 5) = "=R1C2" & "&"" SD""" ' Hidden
    StatsArr(3, c2 + 6) = "=R1C2" & "&"" Count""" ' Hidden
    StatsArr(3, c2 + 7) = ChrW(916) & "Ct" ' Hidden
    StatsArr(3, c2 + 8) = "Pooled SD" ' Hidden
    StatsArr(3, c2 + 9) = ChrW(916) & "Ct (ref)" ' Hidden
    StatsArr(3, c2 + 10) = ChrW(916) & ChrW(916) & "Ct" ' Hidden
    StatsArr(3, c2 + 11) = "95CI" ' Hidden
    For L2 = 1 To MaxRep
        StatsArr(3, c3 - 1 + L2) = "RQ " & L2
    Next L2
    StatsArr(3, c3 + MaxRep) = "Mean"
    StatsArr(3, c3 + 1 + MaxRep) = "SD"
    StatsArr(3, c4) = "2^(-" & ChrW(916) & ChrW(916) & "Ct)"
    StatsArr(3, c4 + 1) = "RQ Max"
    StatsArr(3, c4 + 2) = "RQ Min"
    
    LookupCol = "R5C3:R" & UBound(SortedArr) + 4 & "C3"
    LookupArr = "R5C3:R" & UBound(SortedArr) + 4 & "C" & 12 + MaxRep
    IndexMatchStr1 = "=INDEX(" & LookupArr & ",MATCH(RC1&"";;""&R1C2," & LookupCol & ",0),"
    IndexMatchStr2 = "=INDEX(" & LookupArr & ",MATCH(R2C2&"";;""&RC2," & LookupCol & ",0),"
    Row_RawCt = "RC[-" & 12 + MaxRep & "]"
    NotBlankCond = "IF(AND(LEN(RC[-2])>0,LEN(RC[-1])>0)"
    IncreaseCond = "IF(AND(RC[-2]>1,RC[-1]>1),""Increase"""
    DecreaseCond = "IF(AND(RC[-2]<1,RC[-1]<1),""Decrease"",""No change"")"
    
    ' Copy raw data.
    For L2 = 1 To UBound(SortedArr)
        StatsArr(L2 + 3, 0) = SortedArr(L2, 0)
        StatsArr(L2 + 3, 1) = SortedArr(L2, 1)
        
        For L3 = 2 To UBound(SortedArr, 2)
            StatsArr(L2 + 3, L3 + 1) = SortedArr(L2, L3)
        Next L3
    Next L2
    
    ' Add calculation formulae.
    For L2 = 4 To UBound(StatsArr)
        StatsArr(L2, 2) = "=RC[-2]&"";;""&RC[-1]" ' Condition;;Gene
        StatsArr(L2, c2) = "=IFERROR(AVERAGE(RC[-" & MaxRep & "]:RC[-1]),"""")"  ' Ct Mean (target)
        StatsArr(L2, c2 + 1) = "=IFERROR(RC[1]/RC[-1]*100,"""")" ' CV
        
        ' Hidden.
        StatsArr(L2, c2 + 2) = "=IFERROR(STDEV(RC[-" & 2 + MaxRep & "]:RC[-3]),"""")" ' Ct SD (target)
        StatsArr(L2, c2 + 3) = "=COUNT(RC[-" & 3 + MaxRep & "]:RC[-4])" ' Count (target)
        StatsArr(L2, c2 + 4) = IndexMatchStr1 & 2 + MaxRep & ")" ' Mean (endo)
        StatsArr(L2, c2 + 5) = IndexMatchStr1 & 4 + MaxRep & ")" ' SD (endo)
        StatsArr(L2, c2 + 6) = IndexMatchStr1 & 5 + MaxRep & ")" ' Count (endo
        StatsArr(L2, c2 + 7) = "=IFERROR(RC[-7]-RC[-3],"""")" ' dCt Mean
        StatsArr(L2, c2 + 8) = "=IFERROR(SQRT(RC[-6]^2/RC[-5]+RC[-3]^2/RC[-2]),"""")" ' Pooled SD
        StatsArr(L2, c2 + 9) = IndexMatchStr2 & 9 + MaxRep & ")" ' dCt (ref)
        StatsArr(L2, c2 + 10) = "=IFERROR(RC[-3]-RC[-1],"""")" ' ddCt
        StatsArr(L2, c2 + 11) = "=IFERROR(RC[-3]*TINV(" & alpha & ",RC[-8]+RC[-5]-2),"""")" ' 95CI
        
        For L3 = 1 To MaxRep
            Row_ddCt = Row_RawCt & "-RC[-" & 7 + L3 & "]-RC[-" & 2 + L3 & "]"
            StatsArr(L2, c3 - 1 + L3) = "=IFERROR(IF(ISBLANK(" & Row_RawCt & "),"""",2^-(" & Row_ddCt & ")),"""")" ' RQ
        Next L3
        
        StatsArr(L2, c3 + MaxRep) = "=IFERROR(AVERAGE(RC[-" & MaxRep & "]:RC[-1]),"""")"  ' RQ Mean
        StatsArr(L2, c3 + 1 + MaxRep) = "=IFERROR(STDEV(RC[-" & 1 + MaxRep & "]:RC[-2]),"""")" ' RQ SD
        StatsArr(L2, c4) = "=IFERROR(2^-RC[-" & 4 + MaxRep & "],"""")" ' 2^(-ddCt)
        StatsArr(L2, c4 + 1) = "=IFERROR(2^(-RC[-" & 5 + MaxRep & "]+RC[-" & 4 + MaxRep & "]),"""")" ' RQ Max
        StatsArr(L2, c4 + 2) = "=IFERROR(2^(-RC[-" & 6 + MaxRep & "]-RC[-" & 5 + MaxRep & "]),"""")" ' RQ Min
        StatsArr(L2, c4 + 3) = "=IFERROR(" & NotBlankCond & "," & IncreaseCond & "," & DecreaseCond & "),""""),"""")" ' Extra
    Next L2
    
    ProcessCtData = StatsArr
End Function

Function GetParticularColumn(Arr2D, RowStr, ColStr)
    Dim L2 As Long
    Dim L3 As Long
    
    For L2 = 1 To UBound(Arr2D)
        If Arr2D(L2, LBound(Arr2D, 2)) Like RowStr Then
            For L3 = LBound(Arr2D, 2) + 1 To UBound(Arr2D, 2)
                If Arr2D(L2, L3) Like ColStr Then
                    GetParticularColumn = L3
                    Exit Function
                End If
            Next L3
        End If
    Next L2
End Function

Function PivotForPRISM(Arr2D, ConditionsArr, GenesArr, MaxRep) As Variant()
' OutputArr(1) ~~> Mean+SD
' OutputArr(2) ~~> Mean+95CI
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim L5 As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim WorkSheetName As String
    Dim GrandOutputArr(1 To 2) As Variant
    
    ReDim OutputArr1(0 To UBound(GenesArr), 0 To UBound(ConditionsArr) * MaxRep) As String
    ReDim OutputArr2(0 To UBound(GenesArr), 0 To UBound(ConditionsArr) * 3) As String
    
    ' Header rows.
    OutputArr1(0, 0) = "Copy to PRISM (Replicates)"
    OutputArr2(0, 0) = "Copy to PRISM (Mean+95CI)"
    For L2 = 1 To UBound(ConditionsArr)
        OutputArr1(0, 1 + (L2 - 1) * MaxRep) = ConditionsArr(L2)
        For L3 = 2 To MaxRep
            OutputArr1(0, L3 + (L2 - 1) * MaxRep) = "RQ " & L3
        Next L3
        
        OutputArr2(0, 1 + (L2 - 1) * 3) = ConditionsArr(L2)
        OutputArr2(0, 2 + (L2 - 1) * 3) = "Upper limit"
        OutputArr2(0, 3 + (L2 - 1) * 3) = "Lower limit"
    Next L2
    
    ' Get latest Workings sheet name.
    For L2 = Sheets.Count To 1 Step -1
        If Sheets(L2).Name Like "Workings*" Then
            WorkSheetName = Sheets(L2).Name
            Exit For
        End If
    Next L2
    
    ' Get column numbers.
    c1 = GetParticularColumn(Arr2D, "*Condition*", "RQ 1") ' 1 column before "RQ 1"
    c2 = GetParticularColumn(Arr2D, "*Condition*", "2^(-*") ' 1 column before "2^(-ddCt)"
    
    ' Populate table.
    For L2 = 1 To UBound(GenesArr)
        OutputArr1(L2, 0) = GenesArr(L2)
        OutputArr2(L2, 0) = GenesArr(L2)
        
        For L3 = 1 To UBound(ConditionsArr)
            For L4 = 1 To UBound(Arr2D)
                If Arr2D(L4, 0) = ConditionsArr(L3) And Arr2D(L4, 1) = GenesArr(L2) Then
                    For L5 = 1 To MaxRep
                        OutputArr1(L2, L5 + (L3 - 1) * MaxRep) = "=" & WorkSheetName & "!R" & L4 + 1 & "C" & c1 + L5
                    Next L5
                    
                    For L5 = 1 To 3
                        OutputArr2(L2, L5 + (L3 - 1) * 3) = "=" & WorkSheetName & "!R" & L4 + 1 & "C" & c2 + L5
                    Next L5
                End If
            Next L4
        Next L3
    Next L2
    
    GrandOutputArr(1) = OutputArr1
    GrandOutputArr(2) = OutputArr2
    
    PivotForPRISM = GrandOutputArr
End Function

Function PlotClusterColumnsMeanSD()
    Dim L2 As Long
    Dim L3 As Long
    Dim MaxRep As Long
    Dim nGenes As Long
    Dim nConditions As Long
    Dim SheetArr() As Variant
    Dim MeanSDArr() As String
    
    SheetArr = ReindexArr2D(ActiveSheet.UsedRange.Value2)
    
    MaxRep = CLng(Right(SheetArr(0, UBound(SheetArr, 2)), Len(SheetArr(0, UBound(SheetArr, 2))) - 3))
    nGenes = UBound(SheetArr)
    nConditions = UBound(SheetArr, 2) \ MaxRep
    
    ReDim MeanSDArr(0 To nGenes, 0 To 2 * nConditions) As String
    
    MeanSDArr(0, 0) = "PlotData"
    
    For L2 = 1 To nConditions
        MeanSDArr(0, L2) = "'" & SheetArr(0, 1 + (L2 - 1) * MaxRep)
        MeanSDArr(0, nConditions + L2) = "SD(" & SheetArr(0, 1 + (L2 - 1) * MaxRep) & ")"
    Next L2
    
    For L2 = 1 To nGenes
        MeanSDArr(L2, 0) = SheetArr(L2, 0)
        
        For L3 = 1 To nConditions
            MeanSDArr(L2, L3) = "=AVERAGE(R[0]C" & 2 + (L3 - 1) * MaxRep & ":R[0]C" & 1 + L3 * MaxRep & ")"
            MeanSDArr(L2, nConditions + L3) = "=STDEV(R[0]C" & 2 + (L3 - 1) * MaxRep & ":R[0]C" & 1 + L3 * MaxRep & ")"
        Next L3
    Next L2
    
    PrintArr2D MeanSDArr, 1, UBound(SheetArr, 2) + 3
    
    
    
    ReDim ErrArr(1 To 2, 1 To nConditions) As String
    
    For L2 = 1 To nConditions
        ErrArr(1, L2) = Cells(2, UBound(SheetArr, 2) + 3 + nConditions + L2).Resize(nGenes).Address
        ErrArr(2, L2) = Cells(2, UBound(SheetArr, 2) + 3 + nConditions + L2).Resize(nGenes).Address
    Next L2
    
    PlotClusterColumnsChart Range("A" & UBound(SheetArr) + 3 & ":R33"), Cells(1, UBound(SheetArr, 2) + 3).Resize(1 + nGenes, 1 + nConditions), nConditions, ErrArr
End Function

Function PlotClusterColumnsMean95CI()
    Dim L2 As Long
    Dim L3 As Long
    Dim nGenes As Long
    Dim nConditions As Long
    Dim SheetArr() As Variant
    Dim Mean95CIArr() As String
    
    SheetArr = ReindexArr2D(ActiveSheet.UsedRange.Value2)
    
    nGenes = UBound(SheetArr)
    nConditions = UBound(SheetArr, 2) \ 3
    
    ReDim Mean95CIArr(0 To nGenes, 0 To 3 * nConditions) As String
    
    Mean95CIArr(0, 0) = "PlotData"
    
    For L2 = 1 To nConditions
        Mean95CIArr(0, L2) = "'" & SheetArr(0, 1 + (L2 - 1) * 3)
        Mean95CIArr(0, nConditions + L2) = "upper(" & SheetArr(0, 1 + (L2 - 1) * 3) & ")"
        Mean95CIArr(0, 2 * nConditions + L2) = "lower(" & SheetArr(0, 1 + (L2 - 1) * 3) & ")"
    Next L2
    
    For L2 = 1 To nGenes
        Mean95CIArr(L2, 0) = SheetArr(L2, 0)
        
        For L3 = 1 To nConditions
            Mean95CIArr(L2, L3) = "=0+R[0]C" & 2 + (L3 - 1) * 3 ' 2^(-ddCt)
            Mean95CIArr(L2, nConditions + L3) = "=R[0]C" & 3 + (L3 - 1) * 3 & "-R[0]C[-" & nConditions & "]" ' Upper error
            Mean95CIArr(L2, 2 * nConditions + L3) = "=R[0]C[-" & 2 * nConditions & "]-R[0]C" & 4 + (L3 - 1) * 3 ' Lower error
        Next L3
    Next L2
    
    PrintArr2D Mean95CIArr, 1, UBound(SheetArr, 2) + 3
    
    
    
    ReDim ErrArr(1 To 2, 1 To nConditions) As String
    
    For L2 = 1 To nConditions
        ErrArr(1, L2) = Cells(2, UBound(SheetArr, 2) + 3 + nConditions + L2).Resize(nGenes).Address
        ErrArr(2, L2) = Cells(2, UBound(SheetArr, 2) + 3 + 2 * nConditions + L2).Resize(nGenes).Address
    Next L2
    
    PlotClusterColumnsChart Range("A" & UBound(SheetArr) + 3 & ":R33"), Cells(1, UBound(SheetArr, 2) + 3).Resize(1 + nGenes, 1 + nConditions), nConditions, ErrArr
End Function

Function PlotClusterColumnsChart(OutRng, SourceRng, nSeries, ErrAddresses)
    Dim L2 As Long
    
    With OutRng
        ActiveSheet.Shapes.AddChart(xlColumnClustered, .Left, .Top, .Width, .Height).Select
    End With
    
    With ActiveChart
        .SetSourceData SourceRng
        .ApplyLayout 1
        .ChartTitle.Delete
        With .Axes(xlValue)
            .MinimumScale = 0
            .HasTitle = True
            .AxisTitle.Text = "Relative mRNA expression"
            .TickLabels.NumberFormat = "0.0"
        End With
        
        For L2 = 1 To nSeries
            With .SeriesCollection(L2)
                .HasErrorBars = True
                .ErrorBar xlY, xlBoth, xlCustom, Range(ErrAddresses(1, L2)), Range(ErrAddresses(2, L2))
'                .ErrorBar xlY, xlErrorBarIncludePlusValues, xlCustom, Range(ErrAddresses(1, L2))
            End With
        Next L2
    End With
End Function

Function DecorateSheet1(ConditionsArr, GenesArr, MaxRep)
    Const r1 As Long = 4 ' Header row
    
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim cN As Long
    
    L2 = 1
    Do Until LenB(Cells(r1, L2)) = 0
        If Not IsError(Cells(r1, L2).Value2) Then
            If Cells(r1, L2).Value2 = "Gene" Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
                
                L3 = r1 + 1
                L4 = r1 + 1
                Do Until LenB(Cells(L3, L2)) = 0
                    Do Until Cells(L4, L2).Value2 <> Cells(L3, L2).Value2
                        L4 = L4 + 1
                    Loop

                    ColorScale2 Range(Cells(L3, L2 + 2), Cells(L4 - 1, L2 + 2 + MaxRep)), "Green"
                    L3 = L4
                Loop
            ElseIf Cells(r1, L2).Value2 = "ConditionGene" Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
                Cells(r1 + 1, L2 + 1).Select
                ActiveWindow.FreezePanes = True
            ElseIf Cells(r1, L2).Value2 = "Raw Ct " & MaxRep Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
            ElseIf Cells(r1, L2).Value2 = "CV (%)" Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
                ColorScale2 Columns(L2), "Yellow" ' CV (%)
                ColorScale2 Columns(L2 + 1), "Yellow" ' SD (target)
                ColorScale2 Columns(L2 + 4), "Yellow" ' SD (endo)
                Range("A1:C1").Offset(0, L2 + 2).EntireColumn.Interior.color = RGB(197, 217, 241) ' Light blue
            ElseIf Right(Cells(r1, L2).Value2, 5) = "Count" Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
            ElseIf Cells(r1, L2).Value2 = "Pooled SD" Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
                ColorScale2 Columns(L2), "Yellow" ' Pooled SD
            ElseIf Cells(r1, L2).Value2 = "95CI" Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
                ColorScale2 Columns(L2), "Yellow" ' 95CI,
                ColorScale3 Range("A1").Resize(1, MaxRep + 1).Offset(0, L2).EntireColumn ' RQs
            ElseIf Cells(r1, L2).Value2 = "RQ " & MaxRep Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
                ColorScale2 Columns(L2 + 2), "Yellow" ' SD
            ElseIf Left(Cells(r1, L2).Value2, 4) = "2^(-" Then
                Columns(L2 - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
            ElseIf Cells(r1, L2).Value2 = "RQ Min" Then
                Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
            End If
        End If
        
        L2 = L2 + 1
    Loop
    cN = L2 - 1
    
    L2 = r1
    Do Until LenB(Cells(L2, 2).Value2) = 0 And LenB(Cells(L2 + 1, 2).Value2) = 0
        If Cells(L2, 2).Value2 <> Cells(L2 + 1, 2).Value2 Then
            Rows(L2).Borders(xlEdgeBottom).LineStyle = xlContinuous
        End If
        
        L2 = L2 + 1
    Loop
    
    With Rows(r1 + 1 & ":" & L2 - 1)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$B$2=$A" & 1 + 4
        .FormatConditions(.FormatConditions.Count).Interior.color = RGB(197, 241, 241)
        
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$B$1=$B" & 1 + 4
        .FormatConditions(.FormatConditions.Count).Interior.color = RGB(197, 217, 241)
    End With
    
    Columns(1).AutoFit
    Rows("1:2").RowHeight = 20
    ActiveSheet.UsedRange.NumberFormat = "0.00"
End Function

Function DecorateSheet2(Optional MaxRep = 3)
    Dim L2 As Long
    Dim rN As Long
    
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    L2 = 1
    Do While LenB(Cells(1, L2).Value2)
        If Cells(1, L2).Value2 = "Lower limit" Or Cells(1, L2).Value2 Like "* " & MaxRep Then
            Columns(L2).Borders(xlEdgeRight).LineStyle = xlContinuous
        End If
        
        L2 = L2 + 1
    Loop
    
    rN = 1
    Do While LenB(Cells(rN + 1, 1).Value2)
        rN = rN + 1
    Loop
    
    Columns(1).AutoFit
    ColorScale3 Rows("2:" & rN)
    Rows("2:" & rN).NumberFormat = "0.00"
    Rows(rN).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Function

Function ColorScale2(Rng, color)
    Dim colorIndex As Long
    
    If UCase(color) = "RED" Then
        colorIndex = 7039480
    ElseIf UCase(color) = "GREEN" Then
        colorIndex = 8109667
    ElseIf UCase(color) = "YELLOW" Then
        colorIndex = 65535
    End If
    
    With Rng
        .FormatConditions.AddColorScale ColorScaleType:=2
        With .FormatConditions(1)
            .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .ColorScaleCriteria(1).FormatColor.color = 16776444 ' White
            .ColorScaleCriteria(2).Type = xlConditionValueHighestValue
            .ColorScaleCriteria(2).FormatColor.color = colorIndex
        End With
    End With
End Function

Function ColorScale3(Rng)
    With Rng
        .FormatConditions.AddColorScale ColorScaleType:=3
        With .FormatConditions(1)
            .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .ColorScaleCriteria(1).FormatColor.color = 13011546 ' Blue
            .ColorScaleCriteria(2).Type = xlConditionValueNumber
            .ColorScaleCriteria(2).value = 1 ' <~~ RQ = 1 implies no change
            .ColorScaleCriteria(2).FormatColor.color = 16776444 ' White
            .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
            .ColorScaleCriteria(3).FormatColor.color = 7039480 ' Red
        End With
    End With
End Function

' ====================
' Meta
' ====================

Sub createButtons1()
    CreateButton Range("F5:H6"), "Select qPCR data (.txt) file", "A01_RawDataAndPlateMap"
    BorderAroundButtons Range("F5:H6")
End Sub

Sub createButtons2()
    CreateButton Range("B2:X3"), "Analyze selected cells", "A02_CalculationTables"
    BorderAroundButtons Range("B2:X3")
End Sub

Sub createButtons3()
    Dim bLabel1 As String
    Dim bLabel2 As String
    Dim bLabel3 As String
    
    If PrintReplicates Then
        bLabel1 = "Print replicates: ON"
    Else
        bLabel1 = "Print replicates: OFF"
    End If
    
    If PrintMean95CI Then
        bLabel2 = "Print mean+95CI: ON"
    Else
        bLabel2 = "Print mean+95CI: OFF"
    End If
    
    If PlotChart Then
        bLabel3 = "Plot cluster columns: ON"
    Else
        bLabel3 = "Plot cluster columns: OFF"
    End If
    
    DeleteButtons
    createButtons2
    CreateButton Range("AA2:AD3"), bLabel1, "ToggleOption1"
    CreateButton Range("AA4:AD5"), bLabel2, "ToggleOption2"
    CreateButton Range("AA6:AD7"), bLabel3, "ToggleOption3"
    BorderAroundButtons Range("AA2:AD7")
End Sub

Function ToggleOption1()
    PrintReplicates = Not PrintReplicates
    createButtons3
End Function

Function ToggleOption2()
    PrintMean95CI = Not PrintMean95CI
    createButtons3
End Function

Function ToggleOption3()
    PlotChart = Not PlotChart
    createButtons3
End Function

Sub DeleteButtons()
    ActiveSheet.Buttons.Delete
End Sub

' ====================
' General functions
' ====================

Function CreateButton(PositionRng, ButtonText, Optional OnAction = vbNullString)
    With PositionRng
        ActiveSheet.Buttons.Add(.Left, .Top, .Width, .Height).Text = ButtonText
    End With
    
    With ActiveSheet.Shapes(ActiveSheet.Shapes.Count)
        .Placement = xlMoveAndSize
        .OnAction = OnAction
    End With
End Function

Function BorderAroundButtons(ButtonsRng)
    With ButtonsRng
        With .Offset(-1, -1).Resize(.Rows.Count + 2, .Columns.Count + 2)
            .Interior.color = RGB(149, 279, 215)
            .BorderAround xlContinuous
        End With
    End With
End Function

Function InList(list, value) As Boolean
    Dim V2 As Variant
    
    For Each V2 In list
        If value = V2 Then
            InList = True
            Exit Function
        End If
    Next V2
    InList = False
End Function

Function PrintArr2D(Arr2D, Optional r1 = 1, Optional c1 = 1, Optional SheetName = vbNullString)
    Dim rN As Long
    Dim cN As Long
    
    rN = UBound(Arr2D) - LBound(Arr2D) + 1
    cN = UBound(Arr2D, 2) - LBound(Arr2D, 2) + 1
    Cells(r1, c1).Resize(rN, cN).Value2 = MakeLiteralArr2D(Arr2D)
    
    If LenB(SheetName) Then ActiveSheet.Name = SheetName
End Function

Function NumToStr(num) As String
    Dim tempNum As Long
    Dim letterIndex As Long
    Dim result As String
    
    tempNum = num
    Do
        letterIndex = ((tempNum - 1) Mod 26)
        result = Chr(letterIndex + 65) & result
        tempNum = (tempNum - letterIndex) \ 26
    Loop While tempNum > 0
    
    NumToStr = result
End Function

Function StrToNum(str) As Long
    Dim L2 As Long
    Dim result As Long
    Dim tempStr As String
    
    tempStr = UCase(str)
    For L2 = 0 To Len(str) - 1
        result = result + (Asc(Right(tempStr, 1)) - 64) * (26 ^ L2)
        tempStr = Left(tempStr, Len(tempStr) - 1)
    Next L2
    
    StrToNum = result
End Function

Function GetFilePath() As String
    With Application
        If .OperatingSystem Like "*Win*" Then
            With .FileDialog(msoFileDialogFilePicker)
                If .Show <> -1 Then MsgBox "No file selected. Exiting script.": Exit Function
                GetFilePath = .SelectedItems(1)
            End With
        ElseIf .OperatingSystem Like "*Mac*" Then
            If val(.Version) > 14 Then
                Dim MacFilePath As String
                On Error Resume Next
                MacFilePath = MacScript("choose file as string")
                If LenB(MacFilePath) = 0 Then MsgBox "No file selected. Exiting script.": Exit Function
                GetFilePath = MacFilePath
            End If
        End If
    End With
End Function

Function readTextFile(filePath) As String()
' Read specified text file line-by-line into a 1D array of strings.
'
    Dim fileNum As Long
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
        readTextFile = Split(Input$(LOF(fileNum), #fileNum), vbNewLine) ' MUST be vbNewLine
    Close #fileNum
End Function

Function TextArrToArrArr(TextArr, Optional Delimiter = vbTab) As Variant()
' Convert an array of delimited strings into an array of arrays.
'
    Dim L2 As Long
    
    ReDim ArrArr(LBound(TextArr) To UBound(TextArr)) As Variant
    
    For L2 = LBound(ArrArr) To UBound(ArrArr)
        ArrArr(L2) = Split(TextArr(L2), Delimiter)
    Next L2
    
    TextArrToArrArr = ArrArr
End Function

Function ReindexArr2D(Arr2D) As Variant()
' Convert a 2D array to base 0 indices for both dimensions.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim rN As Long
    Dim cN As Long
    
    rN = UBound(Arr2D) - LBound(Arr2D) ' Number of rows - 1
    cN = UBound(Arr2D, 2) - LBound(Arr2D, 2) ' Number of columns - 1
    
    ReDim ZeroBaseArr2D(0 To rN, 0 To cN) As Variant
    
    For L2 = 0 To rN
        For L3 = 0 To cN
            ZeroBaseArr2D(L2, L3) = Arr2D(LBound(Arr2D) + L2, LBound(Arr2D, 2) + L3)
        Next L3
    Next L2
    
    ReindexArr2D = ZeroBaseArr2D
End Function

Function ArrArrToArr2D(ArrArr) As Variant()
' Convert an array of arrays to a 2D array.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim c1 As Long
    Dim cN As Long
    
    c1 = 2147483647
    For L2 = LBound(ArrArr) To UBound(ArrArr)
        If c1 > LBound(ArrArr(L2)) Then c1 = LBound(ArrArr(L2))
        If cN < UBound(ArrArr(L2)) Then cN = UBound(ArrArr(L2))
    Next L2
    
    ReDim Arr2D(LBound(ArrArr) To UBound(ArrArr), c1 To cN) As Variant
    
    For L2 = LBound(ArrArr) To UBound(ArrArr)
        For L3 = c1 To UBound(ArrArr(L2))
            Arr2D(L2, L3) = ArrArr(L2)(L3)
        Next L3
    Next L2
    
    ArrArrToArr2D = Arr2D
End Function

Function GetRainbowColorList(nColors) As Long()
    Dim L2 As Long
    Dim RainbowArr(0 To 512) As Long
    
    ReDim OutputArr(1 To nColors) As Long
    
    RainbowArr(0) = RGB(255, 126, 127)
    For L2 = 1 To 128
        RainbowArr(L2) = RGB(255, 127 + L2, 127)
        RainbowArr(L2 + 128) = RGB(255 - L2, 255, 127)
        RainbowArr(L2 + 256) = RGB(127, 255, 127 + L2)
        RainbowArr(L2 + 384) = RGB(127, 255 - L2, 255)
    Next L2
    
    For L2 = 1 To nColors
        OutputArr(L2) = RainbowArr((L2 - 1) * 512 \ (nColors - 1))
    Next L2
    
    GetRainbowColorList = OutputArr
End Function
