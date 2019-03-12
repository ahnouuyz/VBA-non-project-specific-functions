Sub SheetToPPT_Multi9x9()
    SheetToPPT 9, 9
End Sub

Function SheetToPPT(rN, cN)
    Dim L2 As Long
    Dim nMax As Long
    Dim iRng As Range
    Dim PPA As Object
    Dim pptPres As Object
    
    OpenPPT PPA, pptPres
    If pptPres Is Nothing Then
        MsgBox "PPT presentation not found. Ending script."
        Exit Function
    End If
    
    nMax = Application.Max(Columns(cN + 1))
    If nMax = 0 Then nMax = 1
    
    For L2 = 1 To nMax
        Set iRng = Cells(1, 1).Resize(rN, cN).Offset((L2 - 1) * (rN + 1), 0)
        
        CopyRangeToPPT iRng, PPA, pptPres
    Next L2
    
'    pptPres.Save
'    PPA.Quit
End Function

Sub SelectionToPPT_Once()
    Dim PPA As Object
    Dim pptPres As Object
    
    OpenPPT PPA, pptPres
    If pptPres Is Nothing Then
        MsgBox "PPT presentation not found. Ending script."
        Exit Sub
    End If
    
    CopyRangeToPPT Selection, PPA, pptPres
    
'    pptPres.Save
'    PPA.Quit
End Sub

Function OpenPPT(PPA, pptPres)
    Dim FilePath As String
    
    Set PPA = CreateObject("PowerPoint.Application")
    
    MsgBox "Please select the template PowerPoint file."
    FilePath = getFilePath
    If LenB(FilePath) = 0 Then Exit Function
    
    PPA.Presentations.Open FilePath
    Set pptPres = PPA.ActivePresentation
End Function

Function CopyRangeToPPT(iRng, PPA, pptPres)
    Dim nBox As Long
    Dim BoxName As String
    
    ' Template Shape numbers.
    ' 1 - Middle box.
    ' 2 - Left box.
    ' 3 - Template grid.
    ' 4 - Right box.
    ' 5 - Left text box.
    ' 6 - Top text box.
    ' 7 - New table (after pasting).
    BoxName = "NDRC: GUSTO (Diurnal Cortisol) - BOX"
    
    With pptPres
        nBox = .Slides.Count
        BoxName = BoxName & " " & nBox
        
        With .Slides(nBox)
            .Duplicate
            .Select
            
            ' Set text box values.
            .Shapes(5).TextFrame.TextRange.Text = BoxName
            .Shapes(6).TextFrame.TextRange.Text = BoxName
            
'            ' New method. Copy values directly into the existing grid.
'            With .Shapes(3).Table
'                With .Cell(1, 1).Shape
'                    .Select
'                    iRng.Copy
'                    PPA.CommandBars.ExecuteMso "PasteExcelTableDestinationTableStyle"
'
'                    While LenB(.TextFrame.TextRange.Text) < 2
'                        DoEvents
'                    Wend
'                End With
'
'                For r1 = 1 To .Rows.Count
'                    For c1 = 1 To .Columns.Count
'                        With .Cell(r1, c1)
'                            If LenB(.Shape.TextFrame.TextRange.Text) = 2 Then ' <~~ Don't really know what's going on
'                                    .Borders.Item(5).Weight = 1 ' DiagonalUp
'                                    .Borders.Item(6).Weight = 1 ' DiagonalDown
'                            End If
'                        End With
'                    Next c1
'                Next r1
'            End With
            
            ' Old method. Paste a new table, resize and replace old one.
            ' Set font size to 9 if using this method.
            ' Old method is faster.
            iRng.Copy
            PPA.CommandBars.ExecuteMso "PasteSourceFormatting"
            
            ' Wait for shape to be indexed.
            While .Shapes.Count < 7
                DoEvents
            Wend
            
            ' Resize pasted table to match template grid exactly.
            .Shapes(7).Left = .Shapes(3).Left
            .Shapes(7).Top = .Shapes(3).Top
            .Shapes(7).Height = .Shapes(3).Height
            .Shapes(7).Width = .Shapes(3).Width
            
            .Shapes(3).Delete
        End With
    End With
    
    Application.CutCopyMode = False
End Function



Sub SelectionTo9x9()
    If Not IsArray(Selection.Value2) Then
        MsgBox "Please select an array."
        Exit Sub
    End If
    
    With Selection
        SelectionToRxC .Value2, .Rows.Count, .Columns.Count, 9, 9
    End With
End Sub

Function SelectionToRxC(selectArr, rS, cS, rN, cN)
' Rearranges selection to (rN by cN) 2D array(s).
' If more than 1 line selected, will assume rows to be separate entries.
' If multiple rows, arrays will be printed top-to-bottom, separated by an empty row.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim nN As Long
    Dim outRow As Long
    Dim Arr1DArr() As Variant
    Dim ArrRxC() As Variant
    Dim newWB As Workbook
    
    nN = rN * cN
    
    ' Automatically resize selection.
    If rS <> 1 And cS <> 1 And cS <> nN Then
        MsgBox "Resizing Selection from " & rS & "x" & cS & " to " & rS & "x" & nN & "."
        Selection.Resize(rS, nN).Select
        SelectionToRxC Selection.Value2, rS, nN, rN, cN
        Exit Function
    End If
    
    Set newWB = Workbooks.Add
    
    If rS = 1 Or cS = 1 Then
        ' Single line (row or column).
        ArrRxC = Arr1DToArrRxC(selectArr, rN, cN)
        PrintArr2D_Boxes ArrRxC, 1, 1
    ElseIf cS = nN Then
        ' Multi line case(s), assume ALL arranged in rows.
        Arr1DArr = Arr2DToArrArr(selectArr)
        
        For L2 = LBound(Arr1DArr) To UBound(Arr1DArr)
            ArrRxC = Arr1DToArrRxC(Arr1DArr(L2), rN, cN)
            outRow = 1 + (L2 - LBound(Arr1DArr)) * (rN + 1)
            
            PrintArr2D_Boxes ArrRxC, outRow, 1
            Cells(outRow, cN + 1).Value2 = L2
        Next L2
    End If
    
    If Not newWB Is Nothing Then
        With newWB.ActiveSheet.Columns(cN + 1)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
End Function

Function Arr1DToArrRxC(Arr1D, rN, cN) As Variant()
    Dim L2 As Long
    Dim r1 As Long
    Dim c1 As Long
    Dim S2 As String
    Dim V2 As Variant
    Dim ArrN() As Variant
    Dim ArrRxC() As Variant
    
    ReDim ArrN(1 To rN * cN)
    ReDim ArrRxC(1 To rN, 1 To cN)
    
    ' Re-index the input array to start from 1.
    ' Truncate if too long, and give a report.
    For Each V2 In Arr1D
        L2 = L2 + 1
        
        If L2 <= UBound(ArrN) Then
            ArrN(L2) = V2
        Else
            S2 = S2 & "~~> Index: " & L2 & ", Value: " & V2 & Chr(10)
        End If
    Next V2
    
    ' Report truncated values, if any.
    If LenB(S2) Then
        MsgBox "Truncations detected:" & Chr(10) & S2
        Debug.Print "Truncations detected:" & Chr(10) & S2
    End If
    
    ' Procedure to convert 1D to 2D.
    For L2 = 1 To UBound(ArrN)
        r1 = (L2 - 1) \ cN + 1
        c1 = (L2 - 1) Mod cN + 1
        
        ArrRxC(r1, c1) = ArrN(L2)
    Next L2
    
    Arr1DToArrRxC = ArrRxC
End Function

Function PrintArr2D_Boxes(Arr2D, Optional r1 = 1, Optional c1 = 1)
' Assumption: Arr2D is the appropriate size.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim rN As Long
    Dim cN As Long
    Dim rHeight1 As Double
    Dim rHeight2 As Double
    
    rN = UBound(Arr2D, 1) - LBound(Arr2D, 1) + 1
    cN = UBound(Arr2D, 2) - LBound(Arr2D, 2) + 1
    
    rHeight1 = 48
    With Cells(r1, c1).Resize(rN, cN)
        .Value2 = Arr2D
        .Font.Size = 9 ' 8 for new method, 9 for old method
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Borders.Weight = 3
        
        ' Reduce font size by 1 if contents overflow by too much.
        .Rows.RowHeight = rHeight1
        .Rows.AutoFit
        For L2 = 1 To rN
            If rHeight2 < .Rows(L2).Height Then rHeight2 = .Rows(L2).Height
        Next L2
        
        If rHeight2 - 1 > rHeight1 Then
            .Font.Size = .Font.Size - 1
        End If
        .Rows.RowHeight = rHeight1
    End With
    
    ' Cross out empty boxes.
    For L2 = r1 To r1 + rN - 1
        For L3 = c1 To c1 + cN - 1
            If LenB(Cells(L2, L3)) = 0 Then
                Cells(L2, L3).Borders(xlDiagonalDown).Weight = 3
                Cells(L2, L3).Borders(xlDiagonalUp).Weight = 3
            End If
        Next L3
    Next L2
End Function

' ================================================================================================
' Common Functions
' ================================================================================================

Function getFilePath() As String
    With Application
        If .OperatingSystem Like "*Win*" Then
            With .FileDialog(msoFileDialogFilePicker)
                If .Show <> -1 Then MsgBox "No file selected. Exiting script.": Exit Function
                getFilePath = .SelectedItems(1)
            End With
        ElseIf .OperatingSystem Like "*Mac*" Then
            If Val(.Version) > 14 Then
                Dim MacFilePath As String
                
                On Error Resume Next
                MacFilePath = MacScript("choose file as string")
                If LenB(MacFilePath) = 0 Then MsgBox "No file selected. Exiting script.": Exit Function
                getFilePath = MacFilePath
            End If
        End If
    End With
End Function

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



Sub Expand_Raw_Labels()
    Dim L2 As Long
    Dim L3 As Long
    Dim nID As Long
    Dim urArr1()
    Dim urArr2()
    Dim idArr
    
    urArr1 = ActiveSheet.UsedRange.Value2
    urArr2 = ActiveSheet.UsedRange.Value2
    
    For L2 = 2 To UBound(urArr1)
        nID = 0
        idArr = Split(urArr1(L2, 2), ";;")
        
        urArr2(L2, 2) = idArr(0)
        For L3 = 1 To UBound(idArr)
            urArr2(L2, 2) = urArr2(L2, 2) & Chr(10) & idArr(L3)
        Next L3
        
        For L3 = 3 To UBound(urArr1, 2)
            Do While LenB(urArr1(L2, L3))
                If Len(urArr1(L2, L3)) = 2 Then
                    urArr2(L2, L3) = idArr(nID) & Chr(10) & decodeBidigit(urArr1(L2, L3))
                ElseIf Len(urArr1(L2, L3)) = 1 Then
                    urArr2(L2, L3) = vbNullString
                End If
                
                L3 = L3 + 1
            Loop
            
            If Len(urArr1(L2, L3 - 1)) <> 0 And Len(urArr1(L2, L3)) = 0 Then
                nID = nID + 1
            End If
        Next L3
    Next L2
    
    Sheets.Add , ActiveSheet
    
    PrintArr2D urArr2
    copyFormat
    
    With Cells
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Size = 8
    End With
    
    Cells(2, 3).Select
    ActiveWindow.FreezePanes = True
    Columns("A:B").AutoFit
End Sub

Function copyFormat()
    Sheets(ActiveSheet.Previous.Name).UsedRange.Copy
    ActiveSheet.UsedRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Function

Function decodeBidigit(num, Optional timePoint = "7YO") As String
    If Len(num) = 2 Then
        Dim digit_1 As Long
        Dim digit_2 As Long
        Dim dict_1
        Dim dict_2
        
        dict_1 = Array(, "WD1", "WD2", "WD3", "WE1", "WE2")
        dict_2 = Array(, "Awake", "Awake + 15min", "Awake + 30min", "Awake + 1hour", "10am - 12nn", "3pm - 5pm", "7pm - 9pm")
        
        digit_1 = Left(num, 1)
        digit_2 = Right(num, 1)
        
        decodeBidigit = timePoint & " " & dict_1(digit_1) & Chr(10) & dict_2(digit_2)
    End If
End Function

Function PrintArr2D(Arr2D, Optional r1 = 1, Optional c1 = 1)
    Dim rN As Long
    Dim cN As Long
    
    rN = UBound(Arr2D) - LBound(Arr2D) + 1
    cN = UBound(Arr2D, 2) - LBound(Arr2D, 2) + 1
    Cells(r1, c1).Resize(rN, cN).Value2 = Arr2D
End Function
