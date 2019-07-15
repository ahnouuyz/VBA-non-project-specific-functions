' ================================================================================================
' Unclassified
' ================================================================================================

Function inList(value, list) As Boolean
' Check if the given value exists in the given list.
'
    Dim V2 As Variant
    
    For Each V2 In list
        If value = V2 Then
            inList = True
            Exit Function
        End If
    Next V2
End Function

Function toString(num) As String
' Convert the given number to letters.
' Mainly to reference columns in Excel.
'
    Dim tNum As Long
    Dim rLetter As Long
    Dim result As String
    
    tNum = num
    Do
        rLetter = ((tNum - 1) Mod 26)
        result = Chr(rLetter + 65) & result
        tNum = (tNum - rLetter) \ 26
    Loop While tNum > 0
    
    toString = result
End Function

Function toLong(str) As Long
' Convert the given letters to a number.
'
    Dim L2 As Long
    Dim result As Long
    Dim tStr As String
    
    tStr = UCase(str)
    For L2 = 0 To Len(str) - 1
        result = result + (Asc(Right(tStr, 1)) - 64) * (26 ^ L2)
        tStr = Left(tStr, Len(tStr) - 1)
    Next L2
    
    toLong = result
End Function

Function Wait(waitSecs)
    Dim Start As Double
    
    Start = Timer()
    While Timer() < Start + waitSecs
        DoEvents
    Wend
End Function

Sub removeBlankRows()
    Dim L2 As Long
    Dim S2 As String
    Dim brRng As Range
    Dim urArr()
    Dim drArr
    
    ActiveSheet.Copy , ActiveSheet
    
    urArr = ActiveSheet.UsedRange.Value2
    
    For L2 = 1 To UBound(urArr)
        If Application.CountA(Rows(L2)) = 0 Then
            If Not brRng Is Nothing Then
                Set brRng = Union(brRng, Rows(L2))
            Else
                Set brRng = Rows(L2)
            End If
        End If
    Next L2
    
    drArr = Split(brRng.Address(0), ",")
    For L2 = 1 To UBound(drArr) + 1
        S2 = S2 & "  " & L2 & ". " & drArr(L2 - 1) & Chr(10)
    Next L2
    
    brRng.Delete
    MsgBox L2 - 1 & " rows were deleted:" & Chr(10) & S2
    Debug.Print L2 - 1 & " rows were deleted:" & Chr(10) & S2
End Sub

' ================================================================================================
' Generate a list of all files within the selected folder
' ================================================================================================

Sub filesInFolder()
' Create a list of all files in all subfolders of the selected folder.
'
    Dim FSO As Object
    Dim folderPath As String
    Dim reportArrArr() As Variant
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    MsgBox "Please select a root folder.", , "List all files within folder"
    folderPath = getFolderPath
    If LenB(folderPath) = 0 Then Exit Sub
    
    reportArrArr = updateReportArrArr
    folderScan FSO.getFolder(folderPath), reportArrArr
    reportArrArr = ArrArrToArr2D(reportArrArr)
    
    Workbooks.Add
    PrintArr2D reportArrArr
    
    customDecoration1
End Sub

Function customDecoration1()
' This seems out of place...
'
    Rows(1).Interior.Color = 6299648 ' Dark blue
    Rows(1).Font.Color = 16777215 ' White
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    Columns.AutoFit
End Function

Function getFolderPath() As String
    With Application
        If .OperatingSystem Like "*Win*" Then
            With .FileDialog(msoFileDialogFolderPicker)
                If .Show <> -1 Then MsgBox "No folder selected. Exiting script.": Exit Function
                getFolderPath = .SelectedItems(1)
            End With
        ElseIf .OperatingSystem Like "*Mac*" Then
            If val(.Version) > 14 Then
                Dim MacFolderPath As String
                
                On Error Resume Next
                MacFolderPath = MacScript("choose folder as string") ' <~~ NOT TESTED YET (!)
                If LenB(MacFolderPath) = 0 Then MsgBox "No folder selected. Exiting script.": Exit Function
                getFolderPath = MacFolderPath
                
                ' FSO methods would probably not work on a Mac in any case...
            End If
        End If
    End With
End Function

Function folderScan(currentFolder, Optional reportArrArr = Empty)
' Loop through all files in all subfolders of the current folder.
'
    Dim FSO As Object
    Dim fsoFolder As Object
    Dim fsoFile As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Recursive calls into subfolders, if available.
    For Each fsoFolder In currentFolder.SubFolders
        folderScan FSO.getFolder(fsoFolder), reportArrArr
    Next fsoFolder
    
    For Each fsoFile In currentFolder.Files
        ' Other instructions for file processing may be written in here.
        
        reportArrArr = updateReportArrArr(reportArrArr, fsoFile)
    Next fsoFile
End Function

Function updateReportArrArr(Optional reportArrArr = Empty, Optional iFile = Empty) As Variant()
' Update/Create an array of arrays to record file metadata.
'
    Dim newEntry(1 To 3) As String
    
    If Not (IsEmpty(reportArrArr) Or IsEmpty(iFile)) Then
        ReDim Preserve reportArrArr(0 To UBound(reportArrArr) + 1)
        
        newEntry(1) = UBound(reportArrArr)
        newEntry(2) = iFile.Name
        newEntry(3) = iFile.Path
        reportArrArr(UBound(reportArrArr)) = newEntry
        
        updateReportArrArr = reportArrArr
    ElseIf IsEmpty(reportArrArr) Then
        Dim newArrArr() As Variant
        ReDim newArrArr(0)
        
        newEntry(1) = "No."
        newEntry(2) = "File name"
        newEntry(3) = "File path"
        newArrArr(0) = newEntry
        
        If IsEmpty(iFile) Then
            updateReportArrArr = newArrArr
        Else
            updateReportArrArr = updateReportArrArr(newArrArr, iFile)
        End If
    End If
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

' ================================================================================================
' Read data from text files.
' ================================================================================================

Sub testFileReader()
    Dim filePath As String
    Dim TextArr() As String
    Dim ArrArr() As Variant
    Dim Arr2D() As Variant
    
    MsgBox "Please select file."
    filePath = getFilePath
    If LenB(filePath) = 0 Then Exit Sub
    
    TextArr = readTextFile(filePath)
    ArrArr = TextArrToArrArr(TextArr)
    Arr2D = ArrArrToArr2D(ArrArr)
    
    Workbooks.Add
    PrintArr2D Arr2D
End Sub

Function getFilePath() As String
    With Application
        If .OperatingSystem Like "*Win*" Then
            With .FileDialog(msoFileDialogFilePicker)
                If .Show <> -1 Then MsgBox "No file selected. Exiting script.": Exit Function
                getFilePath = .SelectedItems(1)
            End With
        ElseIf .OperatingSystem Like "*Mac*" Then
            If val(.Version) > 14 Then
                Dim MacFilePath As String
                
                On Error Resume Next
                MacFilePath = MacScript("choose file as string")
                If LenB(MacFilePath) = 0 Then MsgBox "No file selected. Exiting script.": Exit Function
                getFilePath = MacFilePath
            End If
        End If
    End With
End Function

Function readTextFile(filePath) As String()
' Store data from text file (line-by-line) in an array of strings.
' Delimiter has to be vbNewLine or Chr(10).
' Using vbCr results in strange behavior. Not sure why...
'
    Dim fileNum As Long
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
        readTextFile = Split(Input$(LOF(fileNum), #fileNum), Chr(10))
    Close #fileNum
End Function

Function TextArrToArrArr(TextArr, Optional Delimiter = vbTab) As Variant()
' Convert an array of delimited strings into an array of arrays.
'
    Dim L2 As Long
    Dim ArrArr() As Variant
    
    ReDim ArrArr(LBound(TextArr) To UBound(TextArr))
    
    For L2 = LBound(ArrArr) To UBound(ArrArr)
        If LenB(TextArr(L2)) Then
            ArrArr(L2) = Split(TextArr(L2), Delimiter)
        End If
    Next L2
    
    TextArrToArrArr = ArrArr
End Function

' ================================================================================================
' QuickSort (for array of values and array of arrays)
' ================================================================================================

Function swapValues(value1, value2)
    Dim V2 As Variant
    
    V2 = value1
    value1 = value2
    value2 = V2
End Function

Function quickSortArr(Arr, Optional r1 = -1, Optional rN = -1)
' Sort values of an array in ascending lexicographic order.
'
    Dim tr1 As Long
    Dim trN As Long
    Dim pivotValue As Variant
    
    If Not IsArray(Arr) Then
        MsgBox "Not an array, cannot be sorted."
        Exit Function
    End If
    
    If r1 = -1 Then r1 = LBound(Arr)
    If rN = -1 Then rN = UBound(Arr)
    tr1 = r1
    trN = rN
    
    ' Choose value in the middle as pivot.
    pivotValue = Arr((r1 + rN) \ 2)
    
    Do While tr1 <= trN
        ' Earliest value larger than pivot is in the wrong space.
        Do While Arr(tr1) < pivotValue And tr1 < rN
            tr1 = tr1 + 1
        Loop
        
        ' Latest value smaller than pivot is in the wrong space.
        Do While Arr(trN) > pivotValue And trN > r1
            trN = trN - 1
        Loop
        
        ' Swap positions.
        If tr1 <= trN Then
            swapValues Arr(tr1), Arr(trN)
            tr1 = tr1 + 1
            trN = trN - 1
        End If
    Loop
    
    If r1 < trN Then quickSortArr Arr, r1, trN
    If tr1 < rN Then quickSortArr Arr, tr1, rN
End Function

Function quickSortArrArr(ArrArr, cKey, Optional r1 = -1, Optional rN = -1)
' Sort an array of arrays in ascending lexicographic order w.r.t. a key column.
'
    Dim tr1 As Long
    Dim trN As Long
    Dim pivotValue As Variant
    
    If Not IsArray(ArrArr) Then
        MsgBox "Not an array, cannot be sorted."
        Exit Function
    End If
    
    If r1 = -1 Then r1 = LBound(ArrArr)
    If rN = -1 Then rN = UBound(ArrArr)
    tr1 = r1
    trN = rN
    
    ' Choose value in the middle as pivot.
    pivotValue = ArrArr((r1 + rN) \ 2)(cKey)
    
    Do While tr1 <= trN
        ' Earliest value larger than pivot is in the wrong space.
        Do While ArrArr(tr1)(cKey) < pivotValue And tr1 < rN
            tr1 = tr1 + 1
        Loop
        
        ' Latest value smaller than pivot is in the wrong space.
        Do While ArrArr(trN)(cKey) > pivotValue And trN > r1
            trN = trN - 1
        Loop
        
        ' Swap positions.
        If tr1 <= trN Then
            swapValues ArrArr(tr1), ArrArr(trN)
            tr1 = tr1 + 1
            trN = trN - 1
        End If
    Loop
    
    If r1 < trN Then quickSortArrArr ArrArr, cKey, r1, trN
    If tr1 < rN Then quickSortArrArr ArrArr, cKey, tr1, rN
End Function

' ================================================================================================
' Indirect interaction with Excel
' ================================================================================================

Function makeLiteral(val) As Variant
    If LenB(val) Then ' Leave blanks alone
        If Not IsNumeric(val) Then ' Do not make numbers literal
            If Left(val, 1) <> "=" Then ' Do not make equations literal
                If Left(val, 1) <> "'" Then ' Do not add more than one apostrophe
                    makeLiteral = "'" & val ' Append an apostrophe to the left
                    Exit Function
                End If
            End If
        End If
    End If
    makeLiteral = val ' Return the original value, unmodified
End Function

Function makeLiteralArr2D(Arr2D) As Variant()
    Dim L2 As Long
    Dim L3 As Long
    
    ReDim NewArr(LBound(Arr2D) To UBound(Arr2D), LBound(Arr2D, 2) To UBound(Arr2D, 2)) As Variant
    
    For L2 = LBound(Arr2D) To UBound(Arr2D)
        For L3 = LBound(Arr2D, 2) To UBound(Arr2D, 2)
            NewArr(L2, L3) = makeLiteral(Arr2D(L2, L3))
        Next L3
    Next L2
    
    makeLiteralArr2D = NewArr
End Function

' ================================================================================================
' Direct interaction with Excel
' ================================================================================================

Function testCreateButton()
    Workbooks.Add
    
    createButton Range("D5:G10"), "Test" & Chr(10) & "Button"
End Function
