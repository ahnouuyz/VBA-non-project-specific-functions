' ================================================================================================
' Generate a list of terminal folders with file counts and folder sizes
' ================================================================================================

' Change to True if hierarchy includes the DICOM folder (e.g.):
'   ..\6YO\DICOM\010-12345\YYYYMMDD\XXXXXXXX\*
Private Const withDICOM As Boolean = False



Sub Create_Inventory()
' Create a list of all terminal folders in the selected folder.
' Count number of files in each terminal folder as well as size of terminal folder.
'
    Dim t0 As Double
    Dim folderPath As String
    Dim reportArrArr() As Variant
    Dim FSO As Object
    
    MsgBox _
        "The folder hierarchy should be as follows (e.g.):" & Chr(10) & _
        "    ..\6YO\010-12345\YYYYMMDD\XXXXXXXX\*" & Chr(10) & _
        "where the asterisk (*) stands for individual files." & Chr(10) & Chr(10) & _
        "Please select the folder equivalent to 6YO in the example or higher." & Chr(10) & _
        "The last 4 levels of folders will be reported." & Chr(10) & _
        "The process may take a few minutes (depending on number of items)." & Chr(10) & Chr(10) & _
        "If the following hierarchy is used instead:." & Chr(10) & _
        "    ..\6YO\DICOM\010-12345\YYYYMMDD\XXXXXXXX\*" & Chr(10) & _
        "Press Alt+F11 to open VBA editor and change the value for withDICOM to True." & Chr(10) & _
        "", , "Select a folder containing MRI data"
    folderPath = getFolderPath
    If LenB(folderPath) = 0 Then Exit Sub
    
    t0 = Timer()
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    reportArrArr = updateReportArrArr
    folderScan FSO.GetFolder(folderPath), reportArrArr

    Workbooks.Add
    PrintArr2D ArrArrToArr2D(reportArrArr)
    customDecoration1

    MsgBox "Time taken: " & Round(Timer() - t0, 5) & " secs."
End Sub

Function customDecoration1()
    Rows(1).Interior.Color = 6299648 ' Dark blue
    Rows(1).Font.Color = 16777215 ' White
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    ActiveSheet.UsedRange.HorizontalAlignment = xlCenter
    Columns.AutoFit
End Function

Function folderScan(currentFolder, Optional reportArrArr = Empty)
' Loop through all files in all subfolders of the current folder.
'
    Dim L2 As Long
    Dim pathArr() As String
    Dim infoArr(1 To 6) As Variant
    Dim fsoFolder As Object
    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    For Each fsoFolder In currentFolder.SubFolders
        folderScan FSO.GetFolder(fsoFolder), reportArrArr
    Next fsoFolder
    
    ' If current folder is a terminal folder, look at the files inside.
    If currentFolder.SubFolders.Count = 0 Then
        ' Extract the names of the last 4 levels of folders.
        ' Is there a more elegant way to do this?
        pathArr = Split(currentFolder.Path, "\")
        
        If withDICOM Then
            For L2 = UBound(pathArr) - 2 To UBound(pathArr)
                infoArr(L2 - UBound(pathArr) + 4) = pathArr(L2)
            Next L2
            infoArr(1) = pathArr(UBound(pathArr) - 4)
        Else
            For L2 = UBound(pathArr) - 3 To UBound(pathArr)
                infoArr(L2 - UBound(pathArr) + 4) = pathArr(L2)
            Next L2
        End If
        
        infoArr(5) = currentFolder.Files.Count ' Number of files in folder
        infoArr(6) = currentFolder.Size ' Total size of folder
        
        ' Input entry into the report array.
        reportArrArr = updateReportArrArr(reportArrArr, infoArr)
    End If
End Function

Function updateReportArrArr(Optional reportArrArr = Empty, Optional infoArr = Empty) As Variant()
' Update/Create an array of arrays to record folder metadata.
'
    Dim subjID As String
    Dim newEntry(1 To 12) As String
    
    If Not (IsEmpty(reportArrArr) Or IsEmpty(infoArr)) Then
        ReDim Preserve reportArrArr(0 To UBound(reportArrArr) + 1)
        
        ' Derive the subject's ID from Folder(1) name.
        ' Rules may require further refinement.
        If Len(infoArr(2)) = 9 Then
            subjID = infoArr(2)
        ElseIf Len(infoArr(2)) > 9 Then
            If Mid(infoArr(2), 10, 1) = "-" And Len(infoArr(2)) = 13 Then
                subjID = infoArr(2)
            Else
                subjID = Left(infoArr(2), 9)
            End If
        End If
        
        newEntry(1) = infoArr(1) ' Time point (4.5YO or 6YO)
        newEntry(3) = subjID ' ID
        newEntry(4) = "=IF(R[-1]C[-1]=RC[-1],IF(R[-1]C[2]=RC[2],R[-1]C,R[-1]C+1),1)" ' Disk No. within ID
        newEntry(6) = infoArr(2) ' Folder(1) name
        newEntry(7) = "=IF(R[-1]C[-1]=RC[-1],IF(R[-1]C[1]=RC[1],R[-1]C,R[-1]C+1),1)" ' Folder(2) count within Folder(1) - Should be 1 for everything
        newEntry(8) = infoArr(3) ' Folder(2) name
        newEntry(9) = "=IF(R[-1]C[-3]=RC[-3],IF(R[-1]C[1]=RC[1],R[-1]C,R[-1]C+1),1)" ' Folder(3) count within Folder(1)
        newEntry(10) = infoArr(4) ' Folder(3) name
        newEntry(11) = infoArr(5) ' File count
        newEntry(12) = infoArr(6) ' Folder size (bytes)
        
        If UBound(reportArrArr) <> 1 Then
            newEntry(2) = "=IF(R[-1]C[1]=RC[1],R[-1]C,R[-1]C+1)" ' ID count
            newEntry(5) = "=IF(R[-1]C[1]=RC[1],R[-1]C,R[-1]C+1)" ' Folder(1) count
        Else
            newEntry(2) = 1
            newEntry(5) = 1
        End If
        
        reportArrArr(UBound(reportArrArr)) = newEntry
        
        updateReportArrArr = reportArrArr
    ElseIf IsEmpty(reportArrArr) Then
        Dim newArrArr() As Variant
        ReDim newArrArr(0)
        
        newEntry(1) = "Time Point"
        newEntry(2) = "ID No."
        newEntry(3) = "ID"
        newEntry(4) = "Disk No."
        newEntry(5) = "Folder(1) No."
        newEntry(6) = "Folder(1) Name [ID]"
        newEntry(7) = "Folder(2) No."
        newEntry(8) = "Folder(2) Name [Date]"
        newEntry(9) = "Folder(3) No."
        newEntry(10) = "Folder(3) Name"
        newEntry(11) = "File Count"
        newEntry(12) = "Total Size (bytes)"
        
        newArrArr(0) = newEntry
        
        If Not IsEmpty(infoArr) Then
            updateReportArrArr = updateReportArrArr(newArrArr, infoArr)
        Else
            updateReportArrArr = newArrArr
        End If
    End If
End Function

Sub createButton_1()
    createButton Range("D3:G4"), "Generate inventory", "Create_Inventory"
End Sub

' ================================================================================================
' Contingency
'
' Possible issue 1:
' - File hierarchies:
'   (1) ..\6YO\010-12345\YYYYMMDD\XXXXXXXX\*
'   (2) ..\6YO\010-12345\DICOM\YYYYMMDD\XXXXXXXX\*
' - Currently, (1) is being used.
' - If (2) is required, the following procedure would sort it out.
'
' Possible issue 2:
' - If DICOMDIR is required:
'   - Will have to resolve Issue 1 above.
'   - May have to read the CDs again (NooOOoOOO...!!)
' ================================================================================================

Sub Add_DICOM_Folder()
'
    Dim fSeen As Long
    Dim fMoved As Long
    Dim t0 As Double
    Dim folderPath As String
    Dim reportArrArr() As Variant
    Dim reportEntry(1 To 4) As Variant
    Dim FSO As Object
    
    MsgBox _
        "File hierarchies:" & Chr(10) & _
        "(1) ..\6YO\010-12345\YYYYMMDD\XXXXXXXX\*" & Chr(10) & _
        "(2) ..\6YO\010-12345\DICOM\YYYYMMDD\XXXXXXXX\*" & Chr(10) & Chr(10) & _
        "Currently, (1) is being used." & Chr(10) & _
        "If (2) is required, this procedure would sort it out." & Chr(10) & Chr(10) & _
        "The selected folder must be the 4.5YO or 6YO folder." & Chr(10) & _
        "(or equivalent if they have been renamed)" & Chr(10) & _
        "", , "Generate Inventory"
    folderPath = getFolderPath
    If LenB(folderPath) = 0 Then Exit Sub
    
    t0 = Timer()
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ReDim reportArrArr(0)
    reportEntry(1) = "No."
    reportEntry(2) = "Folder name (ID)"
    reportEntry(3) = "Moved?"
    reportEntry(4) = "Issues?"
    reportArrArr(0) = reportEntry
    
    folderScan_1 FSO.GetFolder(folderPath), fSeen, fMoved, reportArrArr
    
    Workbooks.Add
    PrintArr2D ArrArrToArr2D(reportArrArr)
    customDecoration1
    
    MsgBox _
        "Number of folders seen: " & fSeen & Chr(10) & _
        "Number of folders moved: " & fMoved & Chr(10) & _
        "Time taken: " & Round(Timer() - t0, 5) & " secs."
End Sub

Function folderScan_1(currentFolder, fSeen, fMoved, reportArrArr)
' Loop through all files in all subfolders of the current folder.
'
    Dim nFolders As Long
    Dim reportEntry(1 To 4) As Variant
    Dim fsoFolder_1 As Object
    Dim fsoFolder_2 As Object
    Dim FSO As Object
    Dim skipFolder As Boolean
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    nFolders = currentFolder.SubFolders.Count
    For Each fsoFolder_1 In currentFolder.SubFolders
        skipFolder = False
        
        ' Check for correct folder level.
        ' Each subfolder would start with 0XX-XXXXX, check for this pattern.
        If Left(fsoFolder_1.Name, 1) <> "0" Or Mid(fsoFolder_1.Name, 4, 1) <> "-" Then
            Debug.Print "Anomalous folder name: " & fsoFolder_1.Name
            updateReportArrArr_1 reportArrArr, fsoFolder_1.Name, "No", "Anomalous folder name"
            skipFolder = True
        End If
        
        ' Check that there is only 1 subfolder within.
        If fsoFolder_1.SubFolders.Count > 1 Then
            Debug.Print "Found more than 1 date in a disk: " & fsoFolder_1.Name
            updateReportArrArr_1 reportArrArr, fsoFolder_1.Name, "No", "More than 1 subfolder (date) in folder"
            skipFolder = True
        End If
        
        If Not skipFolder Then
            ' There should be only 1 folder in here (YYYYMMDD).
            For Each fsoFolder_2 In fsoFolder_1.SubFolders
                If fsoFolder_2.Name <> "DICOM" Then ' Do not touch any folders named DICOM
                    If Len(fsoFolder_2.Name) = 8 Then ' Only move folder that is likely YYYYMMDD
                        fsoFolder_1.SubFolders.Add "DICOM"
                        FSO.MoveFolder _
                            fsoFolder_2.Path, _
                            fsoFolder_1.Path & "\DICOM\" & fsoFolder_2.Name
                        
                        updateReportArrArr_1 reportArrArr, fsoFolder_1.Name, "Yes", "All good"
                        fMoved = fMoved + 1
                        Exit For ' Overkill
                    End If
                Else
                    updateReportArrArr_1 reportArrArr, fsoFolder_1.Name, "No", "Already DICOM"
                End If
            Next fsoFolder_2
        End If
        
        ' Safety mechanism to prevent infinite looping.
        fSeen = fSeen + 1
        If fSeen > nFolders Then Exit For
    Next fsoFolder_1
End Function

Function updateReportArrArr_1(reportArrArr, folderName, fMoved, fIssues)
    Dim reportEntry(1 To 4) As Variant
    
    ReDim Preserve reportArrArr(0 To UBound(reportArrArr) + 1)
    
    reportEntry(1) = UBound(reportArrArr)
    reportEntry(2) = folderName
    reportEntry(3) = fMoved
    reportEntry(4) = fIssues
    
    reportArrArr(UBound(reportArrArr)) = reportEntry
End Function

' ================================================================================================
' Support module (common functions)
' ================================================================================================

Function getFolderPath() As String
    With Application
        If .OperatingSystem Like "*Win*" Then
            With .FileDialog(msoFileDialogFolderPicker)
                If .Show <> -1 Then MsgBox "No folder selected. Exiting script.": Exit Function
                getFolderPath = .SelectedItems(1)
            End With
'        ElseIf .OperatingSystem Like "*Mac*" Then
'            If Val(.Version) > 14 Then
'                Dim MacFolderPath As String
'
'                On Error Resume Next
'                MacFolderPath = MacScript("choose folder as string") ' <~~ NOT TESTED YET (!)
'                If LenB(MacFolderPath) = 0 Then MsgBox "No folder selected. Exiting script.": Exit Function
'                getFolderPath = MacFolderPath
'
'                ' FSO methods would probably not work on a Mac in any case...
'            End If
        End If
    End With
End Function

Function ArrArrToArr2D(ArrArr) As Variant()
' Convert an array of arrays to a 2D array (retain original base).
'
    Dim L2 As Long
    Dim L3 As Long
    Dim c1 As Long
    Dim cN As Long
    
    c1 = 2147483647
    For L2 = LBound(ArrArr) To UBound(ArrArr)
        If Not IsEmpty(ArrArr(L2)) Then
            If c1 > LBound(ArrArr(L2)) Then c1 = LBound(ArrArr(L2))
            If cN < UBound(ArrArr(L2)) Then cN = UBound(ArrArr(L2))
        End If
    Next L2
    
    ReDim Arr2D(LBound(ArrArr) To UBound(ArrArr), c1 To cN) As Variant
    
    For L2 = LBound(ArrArr) To UBound(ArrArr)
        If Not IsEmpty(ArrArr(L2)) Then
            For L3 = LBound(ArrArr(L2)) To UBound(ArrArr(L2))
                Arr2D(L2, L3) = ArrArr(L2)(L3)
            Next L3
        End If
    Next L2
    
    ArrArrToArr2D = Arr2D
End Function

Function PrintArr2D(Arr2D, Optional r1 = 1, Optional c1 = 1)
    Dim rN As Long
    Dim cN As Long
    
    rN = UBound(Arr2D) - LBound(Arr2D) + 1
    cN = UBound(Arr2D, 2) - LBound(Arr2D, 2) + 1
    Cells(r1, c1).Resize(rN, cN).Value2 = Arr2D
End Function

Function createButton(PositionRng, ButtonText, Optional OnAction = "doNothing")
    With PositionRng
        ActiveSheet.Buttons.Add(.Left, .Top, .Width, .Height).Text = ButtonText
    End With
    
    With ActiveSheet
        With .Shapes(.Shapes.Count)
            .Placement = xlMoveAndSize
            .OnAction = OnAction
        End With
    End With
End Function

Function doNothing()
    
End Function

Sub deleteButtons()
    ActiveSheet.Buttons.Delete
End Sub
