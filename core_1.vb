' ================================================================================================
' Unclassified
' ================================================================================================

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
