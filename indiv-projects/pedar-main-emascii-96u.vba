' ================================================================================================
' Purpose:
' Compile Pedar data (more precisely, Novel Emascii data in multiple files).
'
' The Problem(s):
' Due to the chosen methodology, (Emascii) data for each step is stored in separate files.
' Compiling all the data is expected to be tedious, and prone to errors.
' Some variables of interest may not be calculated by Novel's software.
'
' Requirements:
'     Software involved:
'     1. pedar-X® Recorder SD Card version 25.3.6 (Novel GmbH, Munich, Germany).
'     2. novel-win® (32-bit) version 24.3.43 (Novel GmbH, Munich, Germany).
'     3. Microsoft® Excel® 2016 MSO (16.0.9029.2106) 32-bit (Microsoft Corp., Redmond, WA, USA).
'     4. [Optional] IBM® SPSS® Statistics Version 25 2017 (IBM Corp., Armonk, NY, USA).
'
'     Prerequisites/Preparation:
'     Pedar data for whole walking trials are recorded in .sol files.
'
'     Using the pedar-X® Recorder software:
'         Isolate individual steps from the walking trials.
'         Save as new .sol files.
'         This is done manually and is rather tedious.
'     Adhere to a file naming convention that will assign a unique name to each step.
'     The current system is: "S**M*T*R**.sol" or "S**M*T*R**x.sol", where * are numbers.
'     The string "S**M*T*R**" shall be referred to as StepID.
'
'     Using the novel-win® software, process each .sol file using the desired mask protocol.
'     This is also a tedious process.
'     Seven Emascii files should be produced for each footstep.
'     The file names sould be either "group" or StepID (identical to the name of the .sol file).
'     *** Store these seven files in a folder (with NOTHING ELSE inside).
'     *** Each footstep should have a folder.
'     *** The folder name should also be StepID.
'     Ensure that at least group.sta and StepID.lin are present.
'         Consider creating a "Lite" version by removing the other 5 files.
'
'     Place all folders in one master folder.
'     How the folders are organized within the master folder is inconsequential.
'     They may be grouped by subject, or by model, or not at all.
'         It is recommended that they not be grouped at all if speed is desired.
'         Otherwise, grouping by model offers flexibility in choosing which shoes to compare.
'
' Usage:
' A "do everything" function is now available.
' Once activated, the user will be prompted to select the folder containing all the data.
' That will be all the input required, for now.
' Barring any errors, the output should be generated in around 30 seconds.
'
' Limitations:
' Program currently unable to handle varying numbers of shoes,
'     (which should not be happening in any case).
' Parts of the source code may need modification if a project contains:
'     different file/folder naming systems,
'     etc.
' Shapiro-Wilk test limited to sample sizes of 25 or less.
' Mauchly's test currently not available.
' Sphericity estimates for 2-way ANOVA interaction term are currently not available.
' Operation in Mac OS may not be supported at the moment.
'
' Version:
' 2018/08/06
'
' Author:
' ©2018 Zhuoyuan (Roscoe) Lai, zhuoyuan.lai@connect.qut.edu.au, ahnouuyz@gmail.com
' ================================================================================================

' To Do:
'     [DONE!] Split color scale subroutine into separate functions.
'     [DONE!] Create a function to handle F tables.
'         Elegance will be prioritized over efficiency (calculations are fast).
'     [DONE!] SPSS integration!
'         But it's rather messy now...
'
'     Significance labels for bar charts.
'
'     Rewrite bar charts function.
'     Arrange group color functions together (?)
'     (RE)Group statistics functions together.
'     Tidy up SPSS handling.


' ================================================================================================
' Global Constants
'     Information that cannot be derived elsewhere must be hard-coded here.
'     In time, an option may be available to obtain user input for some of them.
' ================================================================================================

Option Explicit

Private Const scanFreq As Long = 50 ' Scan frequency in Hertz
Private Const scanPeriod As Long = 1000 / scanFreq ' Scan period in milliseconds
Private Const maxLngVal As Long = 2147483647
Private Const maxDblVal As Double = 1.79769313486231E+308
Private Const typeOneErr As Double = 0.05
Private Const errDiv0 As String = "#DIV/0!"

' Names of shoe models 1, 2, 3, 4, 5, and 6, respectively.
' First comma means index 0 of the array will be empty.
Private Const globalShoeNames As String = _
    "," & _
    "Ascent Horizon," & _
    "Brooks Addiction," & _
    "New Balance 857," & _
    "Nike Free RN," & _
    "Ascent Sustain," & _
    "Asics Melbourne"

' New variables derived from force-time data.
Private Const ftCalcVars As String = _
    "Maximum loading rate," & _
    "First force peak," & _
    "Force dip," & _
    "Second force peak," & _
    "Instant of maximum loading rate," & _
    "Instant of first force peak," & _
    "Instant of force dip," & _
    "Instant of second force peak"
' Units
Private Const ftCalcUnit As String = _
    "[N/s]," & _
    "[N]," & _
    "[N]," & _
    "[N]," & _
    "[%ROP]," & _
    "[%ROP]," & _
    "[%ROP]," & _
    "[%ROP]"

' ================================================================================================
' Declare Global Variables
' ================================================================================================

Private staRunOnce As Long
Private numOfShoes As Long
Private numOfSubjs As Long
Private numOfSteps As Long
Private numOfComps As Long
Private numOfMasks As Long ' Not including whole foot
Private varsArr() As String
Private unitArr() As String
Private maskArr() As String
Private subjArr() As Long
Private shoeNumArr() As Long
Private shoeNameArr() As String
Private pairCompArr1() As Long
Private pairCompArr2() As Long
Private globalVarsColl As New Collection
Private globalUnitColl As New Collection
Private globalMaskColl As New Collection

Function resetGlobalVariables()
    staRunOnce = 0
    numOfShoes = 0
    numOfSubjs = 0
    numOfSteps = 0
    numOfComps = 0
    numOfMasks = 0
End Function

' ================================================================================================
' Meta Subroutine
' ================================================================================================

Sub Z00_createMenu()
    Dim L2 As Long
    Dim instructionString() As String
    
    Sheets.Add Sheets(1)
    
    instructionString = Split( _
        "Novel Emascii Data Processing" & vbTab & _
        "" & vbTab & _
        "Raw pedar-x data are stored on .sol files." & vbTab & _
        "Emascii data files are obtained by processing .sol files in novel-win." & vbTab & _
        "  Also referred to as multimask evaluation." & vbTab & _
        "  For novel-win® (32-bit) version 24.3.43:" & vbTab & _
        "    Up to 7 output files will be produced per .sol file." & vbTab & _
        "" & vbTab & _
        "For the methodology where each .sol file contains only a single step:" & vbTab & _
        "  There would be many files generated (usually in the hundereds)." & vbTab & _
        "  Compiling the data from these separate file may be tedious." & vbTab & _
        "  This is why you are here." & vbTab & _
        "" & vbTab & _
        "Congratulations on completing the multimask evaluations!" & vbTab & _
        "    If you have used novel-win® (32-bit) version 24.3.43," & vbTab & _
        "        you should have up to 7 data files per footstep." & vbTab & _
        "Please ensure that there is a folder for each footstep." & vbTab & _
        "Please ensure that the folders contains only the 7 files." & vbTab & _
        "It would be good to have the folder names keep the convention," & vbTab & _
        "    but just make sure no files were renamed or overwritten." & vbTab & _
        "Only 2 of the 7 files will be used, so a delete function is included." & vbTab & _
        "" & vbTab & _
        "Alt+F8 to view list of macros." & vbTab & _
        "Alt+F11 to open VBA editor and view source code.", vbTab)
    
    For L2 = 0 To UBound(instructionString)
        Cells(1 + L2, 1) = instructionString(L2)
    Next L2
    
    callCreateButtons
    
    For L2 = 1 To Sheets.Count
        If Sheets(L2).Name = "Main" Then Exit Sub
    Next L2
    ActiveSheet.Name = "Main"
End Sub

Function callCreateButtons(Optional rowRef = 4, Optional colRef = 12)
'
    Const numOfButtons As Long = 4
    Const bW As Double = 190 ' About 4 default columns long
    Const bH As Double = 59.5 ' About 4 default rows high
    Const bS As Double = bH + 0.5
    
    Dim L2 As Long
    Dim xPos As Long
    Dim yPos As Long
    Dim buttonLabels() As String
    Dim macroList() As String
    
    xPos = Cells(rowRef, colRef).Left + 1.5
    yPos = Cells(rowRef, colRef).Top + 1
    
    buttonLabels = Split( _
        "," & _
        "Complile raw data," & _
        "Input SPSS data," & _
        "Scan folder for files," & _
        "Delete files", ",")
    
    macroList = Split( _
        "," & _
        "B01a_requestDoEverything," & _
        "F01_inputSpssResults," & _
        "B01c_requestFindFiles," & _
        "Z01_deleteFiles,", ",")
    
    For L2 = 1 To numOfButtons
        ActiveSheet.Buttons.Add(xPos, yPos + (L2 - 1) * bS, bW, bH).Text = buttonLabels(L2)
        ActiveSheet.Shapes(L2).OnAction = macroList(L2)
    Next L2
End Function

' ================================================================================================
' Common Functions
' ================================================================================================

Sub B00_mainOperator()
    Dim opVar As Variant
    Do
        opVar = InputBox( _
            "What would you like to do next?" & vbCr & vbCr & _
            "1 ~~> Extract and analyze data" & vbCr & _
            "2 ~~> Input SPSS output data" & vbCr & _
            "3 ~~> List all files within a selected folder" & vbCr & vbCr & _
            "Please enter the corresponding number:", "Select procedure", 1)
        
        If LenB(opVar) = 0 Then MsgBox "No changes made. Exiting script.": Exit Sub
        If opVar <> 1 And opVar <> 2 And opVar <> 3 Then
            MsgBox "Please enter a number from the list.", , "Invalid entry"
        End If
    Loop While opVar <> 1 And opVar <> 2 And opVar <> 3
    
    If opVar = 1 Then B01a_requestDoEverything
    If opVar = 2 Then F01_inputSpssResults
    If opVar = 3 Then B01c_requestFindFiles
End Sub

Function B01a_requestDoEverything()
    MsgBox _
        "Extract .sta and .lin data from files in the selected folder." & vbCr & _
        "This includes all subfolders within the selected folder." & vbCr & _
        "" & vbCr & _
        "3 files will be created:" & vbCr & _
        "    1. Output (.xlsx) file." & vbCr & _
        "    2. A list of files processed." & vbCr & _
        "    3. Data for SPSS analysis (.xlsx) file." & vbCr & vbCr & _
        "The first two will not be saved." & vbCr & _
        "The last one will be automatically saved and closed." & vbCr & vbCr & _
        "Please select a folder.", , "Extract and analyze data"
    Call B02_selectFolder(1)
End Function

Function B01c_requestFindFiles()
    MsgBox _
        "Scan for all files within the selected folder." & vbCr & _
        "This includes all subfolders within the selected folder." & vbCr & _
        "A list of files found will be generated." & vbCr & vbCr & _
        "Please select a folder.", , "Scan folder for files"
    Call B02_selectFolder(3)
End Function

Function B02_selectFolder(iCase As Long)
    Dim L2 As Long
    Dim fldPath As String
    
    ' Obtain the address of the desired folder.
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then MsgBox "No folder selected. Exiting script.": Exit Function
        fldPath = .SelectedItems(1)
    End With
    
    ' This will be the user's last chance to abort the procedure.
    L2 = MsgBox( _
        "The folder with the address: " & vbCr & vbCr & _
        "'" & fldPath & "'" & vbCr & vbCr & _
        "has been selected." & vbCr & vbCr & _
        "This will be the last chance to abort. Proceed?", vbOKCancel, "Confirm folder")
    
    If L2 = vbOK Then
        B03_mainSwitchBoard iCase, fldPath
    Else
        MsgBox "No changes made. Exiting script."
    End If
End Function

Function B03_mainSwitchBoard(iCase As Long, fldPath As String)
    Dim nFile As Long
    Dim FSO As Object
    Dim eSec As Double
    Dim eSec2 As Double
    Dim destPath As String
    Dim ftArr As Variant
    Dim mastArr As Variant
    
    ReDim repArr(0) As String
    ReDim linArr(0) As String
    ReDim staArr(0) As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    eSec = Timer()
    eSec2 = Timer()
    Application.ScreenUpdating = False
    
    repArr(0) = _
        "No." & vbTab & _
        "StepID" & vbTab & _
        "File path"
    
    linArr(0) = _
        "StepID" & vbTab & _
        "Time" & vbTab & _
        "Force" & vbTab & _
        "Area" & vbTab & _
        "Peak P" & vbTab & _
        "Mean P"
    
    staArr(0) = _
        "Variable" & vbTab & _
        "Units" & vbTab & _
        "Mask" & vbTab & _
        "Model" & vbTab & _
        "Subject" & vbTab & _
        "StepID" & vbTab & _
        "Value"
    
    If iCase = 1 Then
        resetGlobalVariables
        destPath = FSO.GetParentFolderName(fldPath)
        
        ' Collect .lin and .sta data.
        B04_folderScan FSO.getFolder(fldPath), nFile, repArr, linArr, staArr, iCase
        
        Debug.Print ".lin and .sta data collection: " & Round(Timer - eSec2, 5) & " sec"
        eSec2 = Timer
        
        ' Process force-time data.
        ftArr = C02_genFinalFtArr(C01_genRawFtArr(linesToTable(linArr)))
        C03_addDataToSta ftArr, staArr
        
        ' Sort force-time array and master table.
        QuickSort ftArr, 1 + LBound(ftArr), UBound(ftArr), 2, 2
        QuickSort staArr, 1 + LBound(staArr), UBound(staArr), 1
        
        ' Generate master table.
        mastArr = linesToTable(staArr)
        
        ' Set global variables.
        setGlobalVariables ftArr ' Set numOfShoes, numOfSubjs, numOfSteps, numOfComps
        setVarsUnitMaskArrs mastArr, 0, 1, 2 ' Set varsList, unitList, maskList
        
        ' Output available information.
        ftArr = arrOfArrTo2dArr(ftArr)
        outputFtArr ftArr
        C04_ftGraphs ftArr
        
'        ' Save .xlsx input for SPSS.
'        D01_spssTableCreate mastArr, destPath
'
'        ' Save .txt file of master table.
'        printArray mastArr, , 0, 1, 0, "None", "None"
'        ActiveWorkbook.SaveAs destPath & "\mastArr " & Format(Now, "yyyymmdd_hhmm"), xlText
'        ActiveWorkbook.Close False
'
'        Debug.Print "Save files: " & Round(Timer - eSec2, 5) & " sec"
'        eSec2 = Timer
        
        E01_masterTableProcess mastArr, Empty
        
        Debug.Print "Entire results table process: " & Round(Timer - eSec2, 5) & " sec"
        
    ElseIf iCase = 3 Then
        B04_folderScan FSO.getFolder(fldPath), nFile, repArr, Empty, Empty, iCase
    End If
    
    Application.ScreenUpdating = True
    eSec = Timer() - eSec
    
    ReDim Preserve repArr(0 To UBound(repArr) + 2) As String
    repArr(UBound(repArr) - 1) = vbTab & vbTab
    repArr(UBound(repArr) - 0) = vbTab & vbTab & "Total time elasped: " & Round(eSec, 4) & " sec"
    
    printArray linesToTable(repArr), , 0, 1, 0, "All", "None"
    Cells(nFile + 3, 3).Select
    
    Debug.Print "Total time elapsed: " & Round(eSec, 4) & " sec"
    Debug.Print "========================================"
End Function

Function B04_folderScan(fsoFld, ByRef nFile, ByRef repArr, ByRef linA, ByRef staA, iCase)
' Iteratively search for subfolders within the selected folder.
' Once the terminal folder is reached, scan files within that folder.
'
    Dim FSO As Object
    Dim iObj As Object
    Dim passStepID As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With fsoFld
        ' Access subdirectories.
        For Each iObj In .SubFolders
            B04_folderScan FSO.getFolder(iObj), nFile, repArr, linA, staA, iCase
        Next iObj
        
        ' Get StepID from .lin files (with same file name as .sol files) during first-pass.
        For Each iObj In .Files
            With iObj
                If Right$(.Name, 4) = ".lin" Then
                    If Left$(.Name, 5) <> "group" Then ' <~~ Disregard group.lin files
                        passStepID = Left$(.Name, Len(.Name) - 4)
                        Exit For
                    End If
                End If
            End With
        Next iObj
        
        ' Get data during second-pass.
        For Each iObj In .Files
            With iObj
                If iCase = 1 Then
                    If .Name = passStepID & ".lin" Then
                        B05b_extractLin .Path, passStepID, linA
                        B05a_fileLog .Path, nFile, repArr, passStepID
                    ElseIf .Name = "group.sta" Then
                        B05c_extractSta .Path, passStepID, staA
                        B05a_fileLog .Path, nFile, repArr, passStepID
                    End If
                ElseIf iCase = 3 Then
                    B05a_fileLog .Path, nFile, repArr, passStepID
                End If
            End With
        Next iObj
    End With
End Function

Function B05a_fileLog(filePath, nFile, repArr, passStepID)
    nFile = nFile + 1
    ReDim Preserve repArr(0 To nFile) As String
    repArr(nFile) = nFile & vbTab & passStepID & vbTab & filePath
End Function

Function B05b_extractLin(filePath, passStepID, ByRef linArr)
' Extract data from StepID.lin (not group.lin) file.
' Force, area, peak pressure, and mean pressure for the right whole foot only.
'
    Const markerString As String = "right foot"
    Const rowsBelowMarker As Long = 6
    
    Dim L2 As Long
    Dim sourceArr() As String
    
    sourceArr = readTextFile(filePath)
    
    ' Verify file (check for possible illegal changes).
    If sourceArr(0) <> "multimask evaluation" Then
        Debug.Print "Error: " & passStepID & ".lin file may be contaminated."
        Exit Function
    End If
    
    ' Find the first table for the right foot and add to linArr.
    For L2 = LBound(sourceArr) To UBound(sourceArr)
        If sourceArr(L2) = markerString Then ' <~~ 1st "right foot"
            L2 = L2 + rowsBelowMarker ' <~~ Move to 1st row of data in table
            
            ' Add table of data to linArr, line-by-line.
            Do Until LenB(sourceArr(L2)) = 0
                ReDim Preserve linArr(0 To UBound(linArr) + 1) As String
                
                linArr(UBound(linArr)) = passStepID & vbTab & sourceArr(L2)
                L2 = L2 + 1
            Loop
            Exit For
        End If
    Next L2
End Function

Function B05c_extractSta(filePath, passStepID, ByRef staArr)
' Extract any available valid data from a single group.sta file.
' Right foot only.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim rowArr() As String
    Dim stepIdMod As String
    Dim stepIdSub As String
    Dim currentMask As String
    Dim sourceArr() As String
    
    sourceArr = readTextFile(filePath)
    
    ' Verify file (check for possible illegal changes).
    If sourceArr(0) <> "groupmask evaluation" Then
        Debug.Print "Error: Step " & passStepID & " group.sta file may be contaminated."
        Exit Function
    End If
    
    If staRunOnce = 0 Then
        setGlobalCollections sourceArr
        staRunOnce = 1 ' <~~ Only run for the first .sta file
    End If
    
    stepIdMod = Mid$(passStepID, 5, 1) ' Shoe number
    stepIdSub = Mid$(passStepID, 2, 2) ' Subject number
    
    For L2 = LBound(sourceArr) To UBound(sourceArr)
        ' Get mask name for upcoming tables.
        If sourceArr(L2) = "left feet" Then
            currentMask = globalMaskColl(sourceArr(L2 - 2)) ' <~~ Mask number is 2 rows above
        End If
        
        ' Extract data for right feet.
        If sourceArr(L2) = "right feet" Then
            L2 = L2 + 5 ' <~~ Table starts 5 rows down
            
            Do Until LenB(sourceArr(L2)) = 0
                rowArr = Split(sourceArr(L2), vbTab)
                If UBound(rowArr) > 1 Then
                    If LenB(rowArr(1)) <> 0 Then ' <~~ Check if data is available
                        If rowArr(1) = rowArr(2) Then ' <~~ Check if min = max
                            ReDim Preserve staArr(0 To UBound(staArr) + 1) As String
                            
                            staArr(UBound(staArr)) = _
                                globalVarsColl(rowArr(0)) & vbTab & _
                                globalUnitColl(rowArr(0)) & vbTab & _
                                currentMask & vbTab & _
                                stepIdMod & vbTab & _
                                stepIdSub & vbTab & _
                                passStepID & vbTab & _
                                rowArr(1) ' <~~ If single step, then mean = min = max
                        Else
                            Debug.Print "Error: Step " & passStepID & " may not be singular."
                            Exit Function
                        End If
                    End If
                End If
                L2 = L2 + 1
            Loop
        End If
    Next L2
End Function


' ================================================================================================
' Functions for Managing Global Variables
' ================================================================================================

Function setGlobalVariables(Optional ftArr = Empty, Optional mastArr = Empty)
' The force-time data array has to be sorted (by StepID) first.
' Count number of shoes, subjects, steps, and pairwise comparisons.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim iSubj As Long
    Dim iShoe As Long
    Dim iStep As Long
    Dim dupArr() As String
    
    ' The 1st entry (index 0) of these arrays will be empty.
    ReDim subjArr(0) As Long
    ReDim shoeNumArr(0) As Long
    ReDim shoeNameArr(0) As String
    ReDim stepIdArr(0) As String
    ReDim stepCountArr(0) As Long
    
    If Not IsEmpty(ftArr) Then
        For L2 = 1 To UBound(ftArr)
            iSubj = ftArr(L2)(0)
            iShoe = ftArr(L2)(1)
            
            If notInArr(subjArr, iSubj) Then pushIntoLngArr subjArr, iSubj
            If notInArr(shoeNumArr, iShoe) Then pushIntoLngArr shoeNumArr, iShoe
            If notInArr(stepIdArr, ftArr(L2)(2)) Then
                pushIntoStrArr stepIdArr, ftArr(L2)(2)
            Else
                MsgBox "DUPLICATE StepID detected: " & ftArr(L2)(2) & " at index " & L2
            End If
            
            ' If same shoe and subject cf. previous row, increase step count by 1.
            If iShoe = ftArr(L2 - 1)(1) Then ' Same shoe
                If iSubj = ftArr(L2 - 1)(0) Then ' Same subject (redundant if properly sorted)
                    iStep = iStep + 1
                End If
            Else
                ' Different shoe implies different condition since data has been sorted.
                ' Record step count after first iteration (which would return 0).
                If L2 <> 1 Then pushIntoLngArr stepCountArr, iStep
                If numOfSteps < iStep Then numOfSteps = iStep ' Set numOfSteps to max. step count
                iStep = 1 ' Re-set step counter
            End If
        Next L2
    ElseIf Not IsEmpty(mastArr) Then
        For L2 = 1 To UBound(mastArr)
            iSubj = mastArr(L2, 4)
            iShoe = mastArr(L2, 3)
            
            If notInArr(subjArr, iSubj) Then pushIntoLngArr subjArr, iSubj
            If notInArr(shoeNumArr, iShoe) Then pushIntoLngArr shoeNumArr, iShoe
            
            If L2 < 385 Then ' WORKAROUND
                If notInArr(stepIdArr, mastArr(L2, 5)) Then
                    pushIntoStrArr stepIdArr, mastArr(L2, 5)
                Else
                    MsgBox "DUPLICATE StepID detected: " & mastArr(L2, 5) & " at index " & L2
                End If
            End If
            
            ' If same shoe and subject cf. previous row, increase step count by 1.
            If iSubj = mastArr(L2 - 1, 4) Then ' Same subject
                If iShoe = mastArr(L2 - 1, 3) Then ' Same shoe
                    iStep = iStep + 1
                End If
            Else
                ' Different shoe implies different condition since data has been sorted.
                ' Record step count after first iteration (which would return 0).
                If L2 <> 1 Then
                    stepCountArr(0) = iStep ' WORKAROUND
                    pushIntoLngArr stepCountArr, iStep
                    
                    If isUniformArr(stepCountArr) Then
                        ' Do nothing.
                    Else
                        MsgBox "Variant number of steps: " & iStep & " at index " & L2
                    End If
                End If
                
                If numOfSteps < iStep Then numOfSteps = iStep ' Set numOfSteps to max. step count
                iStep = 1 ' Re-set step counter
            End If
        Next L2
    Else
        Err.Raise 5
    End If
    
    numOfSubjs = UBound(subjArr)
    numOfShoes = UBound(shoeNumArr)
    
    setPairComps shoeNumArr
    
'    ' Check for different numbers of steps.
'    stepCountArr(0) = iStep ' Workaround ~~> set the first value equal to the last one
'    If isUniformArr(stepCountArr) Then
'        Debug.Print "  There are " & numOfSteps & " steps for all subject-shoe conditions."
'    Else
'        Debug.Print "Variation in number of steps detected:"
'        Debug.Print "  Max. = " & numOfSteps & " steps."
'        For L2 = LBound(stepCountArr) To UBound(stepCountArr)
'            Debug.Print "  Condition " & L2 & ": " & stepCountArr(L2) & " steps"
'        Next L2
'    End If
End Function

Function setPairComps(ByRef shoeNumArr)
    Dim L2 As Long
    Dim L3 As Long
    ReDim pairCompArr1(0) As Long
    ReDim pairCompArr2(0) As Long
    
    For L2 = 1 To numOfShoes
        pushIntoStrArr shoeNameArr, Split(globalShoeNames, ",")(shoeNumArr(L2))
        For L3 = L2 + 1 To numOfShoes
            numOfComps = numOfComps + 1
            pushIntoLngArr pairCompArr1, L2
            pushIntoLngArr pairCompArr2, L3
        Next L3
    Next L2
End Function

Function notInArr(ByRef refArr, newVal) As Boolean
    Dim arrVal As Variant
    For Each arrVal In refArr
        If newVal = arrVal Then
            notInArr = False
            Exit Function
        End If
    Next arrVal
    notInArr = True
End Function

Function pushIntoLngArr(ByRef refArr, newVal)
    ReDim Preserve refArr(0 To UBound(refArr) + 1) As Long
    refArr(UBound(refArr)) = newVal
End Function

Function pushIntoStrArr(ByRef refArr, newVal)
    ReDim Preserve refArr(0 To UBound(refArr) + 1) As String
    refArr(UBound(refArr)) = newVal
End Function

Function isUniformArr(ByRef inArr) As Boolean
    For Each arrVal In inArr
        If arrVal <> inArr(LBound(inArr)) Then
            isUniformArr = False
            Exit Function
        End If
    Next arrVal
    isUniformArr = True
End Function

Function setVarsUnitMaskArrs(ByRef mastArr, varCol, unitCol, maskCol)
' The master table has to be sorted first.
' Update arrays for variable, units, and masks.
'
    Dim L2 As Long
    Dim iVarUnit As String
    
    ' The 1st entry (index 0) of these arrays will be empty.
    ReDim varsArr(0) As String
    ReDim unitArr(0) As String
    ReDim maskArr(0) As String
    ReDim varsUnitArr(0) As String
    
    For L2 = 1 To UBound(mastArr)
        iVarUnit = mastArr(L2, varCol) & mastArr(L2, unitCol) ' Concatenate variable and unit
        
        If notInArr(varsUnitArr, iVarUnit) Then
            pushIntoStrArr varsUnitArr, iVarUnit
            pushIntoStrArr varsArr, mastArr(L2, varCol)
            pushIntoStrArr unitArr, mastArr(L2, unitCol)
        End If
        
        If notInArr(maskArr, mastArr(L2, maskCol)) Then
            pushIntoStrArr maskArr, mastArr(L2, maskCol)
        End If
    Next L2
    
    numOfMasks = UBound(maskArr) - 1 ' Excluding whole foot
End Function

Function setGlobalCollections(ByRef firstStaArr)
' Define collections used when extracting .sta data.
' Called only once (can be disabled) for the first .sta file.
'
    Dim L2 As Long
    Dim rowArr() As String
    Dim maskName As String
    
    Set globalVarsColl = Nothing
    Set globalUnitColl = Nothing
    Set globalMaskColl = Nothing
    
    For L2 = 0 To UBound(firstStaArr)
        ' Store collection of variables, as defined by Novel.
        If firstStaArr(L2) = "Declaration of variables" Then
            L2 = L2 + 2
            Do Until LenB(firstStaArr(L2)) = 0
                rowArr = Split(firstStaArr(L2), vbTab)
                globalVarsColl.Add Item:=rowArr(1), Key:=Left$(rowArr(0), 3) ' V01, V02, ..., V18
                globalUnitColl.Add Item:=rowArr(2), Key:=Left$(rowArr(0), 3) ' V01, V02, ..., V18
                L2 = L2 + 1
            Loop
        End If
        
        ' Store collection of masks, as defined by Novel.
        If firstStaArr(L2) = "Declaration of masks" Then
            L2 = L2 + 2
            Do Until LenB(firstStaArr(L2)) = 0
                rowArr = Split(firstStaArr(L2), vbTab)
                maskName = Right$(rowArr(1), Len(rowArr(1)) - 2) ' Remove colon and space (": ")
                globalMaskColl.Add Item:=maskName, Key:="Mask " & Right$(rowArr(0), 2)
                L2 = L2 + 1
            Loop
        End If
    Next L2
    
    globalMaskColl.Add Item:="Whole foot", Key:="Total"
End Function

' ================================================================================================
' Functions for Processing Force-time Data
' ================================================================================================

Function C01_genRawFtArr(ByRef linArr) As Variant
' Arrange force-time data horizontally (1 step per row).
' Output an array of arrays (allows for sorting later).
'
    Const stepStartTime As Long = scanPeriod ' Step begins at first scan period (20 ms)
    
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim startCol As Long
    Dim numOfRows As Long
    Dim frameCount As Long
    Dim maxStepTime As Long
    Dim ftArrHead As String
    
    ' Determine the number of rows required and find the longest row.
    For L2 = LBound(linArr) To UBound(linArr)
        If linArr(L2, 1) = stepStartTime Then
            If maxStepTime < frameCount Then maxStepTime = frameCount ' Find the longest row
            
            frameCount = 0 ' Re-set for the coming step
            numOfRows = numOfRows + 1
        End If
        
        If linArr(L2, 2) > 0 Then frameCount = frameCount + 1 ' Only count non-zero force values
    Next L2
    If maxStepTime < frameCount Then maxStepTime = frameCount ' If last row is the longest
    
    ' Determine the number of columns required.
    ftArrHead = _
        "Subject," & _
        "Model," & _
        "StepID," & _
        ftCalcVars & _
        ",Time (ms)"
    startCol = UBound(Split(ftArrHead, ",")) + 1 ' 1 column after Time column (0 ms)
    For L2 = 0 To maxStepTime + 1 ' +1 after longest due to adding a 0 at the end
        ftArrHead = ftArrHead & ("," & L2 * scanPeriod)
    Next L2
    
    ' Create new array and input raw data.
    ReDim ftArr(0 To numOfRows) As Variant
    ftArr(0) = Split(ftArrHead, ",")
    
    For L2 = 1 To UBound(linArr)
        If linArr(L2, 2) <> 0 Then ' Start when force > 0
            ReDim ftRow(0 To UBound(Split(ftArrHead, ","))) As Variant
            
            ftRow(0) = Mid$(linArr(L2, 0), 2, 2) ' Subject
            ftRow(1) = Mid$(linArr(L2, 0), 5, 1) ' Model
            ftRow(2) = linArr(L2, 0) ' StepID
            
            ' Transpose force data onto a row.
            L4 = startCol + 1 ' Start under the 20 ms column
            Do
                ftRow(L4) = CDbl(linArr(L2, 2)) ' Force data, convert to Double
                L4 = L4 + 1 ' 1 step right
                L2 = L2 + 1 ' 1 step down
            Loop While linArr(L2, 2) <> 0
            
            ' Restore 0 to the beginning and end.
            ftRow(startCol) = 0
            ftRow(L4) = 0
            
            ' Add row to array.
            L3 = L3 + 1 ' Next row
            ftArr(L3) = ftRow
        End If
    Next L2
    
    C01_genRawFtArr = ftArr
End Function

Function C02_genFinalFtArr(ByRef ftArr) As Variant
' Calculate force-time derived variables:
'     1. Instantaneous maximum loading rate
'     2. First force peak
'     3. Force dip
'     4. Second force peak
'     5-8. The instances of these four events
'
    Dim L2 As Long
    Dim L3 As Long
    Dim startCol As Long
    Dim frameCount As Long
    Dim currentDiff As Double
    
    Do Until IsNumeric(ftArr(0)(startCol))
        startCol = startCol + 1
    Loop
    
    ' Calculations for derived variables.
    For L2 = 1 + LBound(ftArr) To UBound(ftArr)
        frameCount = 0 ' <~~ Re-set
        ftArr(L2)(5) = maxDblVal ' <~~ Re-set second force peak to a large number
        
        ' Find the largest difference between two successive cells, and when it occurs.
        For L3 = startCol To UBound(ftArr(L2))
            If ftArr(L2)(L3) > 0 Then frameCount = frameCount + 1
            currentDiff = ftArr(L2)(L3) - ftArr(L2)(L3 - 1)
            If ftArr(L2)(3) < currentDiff Then
                ftArr(L2)(3) = currentDiff ' Max loading rate (precursor)
                ftArr(L2)(7) = L3 - startCol ' Instant of max loading rate
            End If
        Next L3
        
        ' Find the local maxima (amplitudes and times).
        For L3 = startCol To startCol + frameCount ' <~~ Start to end
            If L3 <= startCol + (frameCount \ 2) Then ' <~~ Before mid-point
                If L3 >= startCol Then ' <~~ After start
                    If ftArr(L2)(4) < ftArr(L2)(L3) Then
                        ftArr(L2)(4) = ftArr(L2)(L3) ' First force peak
                        ftArr(L2)(8) = L3 - (startCol + 1) ' Instant of first force peak
                    End If
                End If
            ElseIf L3 > startCol + (frameCount \ 2) Then ' <~~ After mid-point
                If L3 <= startCol + frameCount Then ' <~~ Before end
                    If ftArr(L2)(6) < ftArr(L2)(L3) Then
                        ftArr(L2)(6) = ftArr(L2)(L3) ' Second force peak
                        ftArr(L2)(10) = L3 - (startCol + 1) ' Instant of second force peak
                    End If
                End If
            End If
        Next L3
        
        ' Find the local minima (amplitude and time).
        For L3 = startCol To startCol + frameCount ' <~~ Start to end
            If L3 >= ftArr(L2)(8) + (startCol + 1) Then ' <~~ After first force peak
                If L3 <= ftArr(L2)(10) + (startCol + 1) Then ' <~~ Before second force peak
                    If ftArr(L2)(5) > ftArr(L2)(L3) Then
                        ftArr(L2)(5) = ftArr(L2)(L3) ' Force dip
                        ftArr(L2)(9) = L3 - (startCol + 1) ' Instant of force dip
                    End If
                End If
            End If
        Next L3
        
        ftArr(L2)(3) = ftArr(L2)(3) * scanFreq ' Max. loading rate in N/s
        ftArr(L2)(7) = ftArr(L2)(7) / frameCount * 100 ' Max. loading rate time
        ftArr(L2)(8) = ftArr(L2)(8) / frameCount * 100 ' First force peak time
        ftArr(L2)(9) = ftArr(L2)(9) / frameCount * 100 ' Force dip time
        ftArr(L2)(10) = ftArr(L2)(10) / frameCount * 100 ' Second force peak time
    Next L2
    
    C02_genFinalFtArr = ftArr
End Function

Function C03_addDataToSta(ByRef ftArr, ByRef staArr)
' Self-explanatory: Add data from ftArr to staArr.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim newVarsArr() As String
    Dim newUnitArr() As String
    
    newVarsArr = Split(ftCalcVars, ",")
    newUnitArr = Split(ftCalcUnit, ",")
    
    ' Add new data to staArr.
    For L2 = 1 To UBound(ftArr)
        For L3 = LBound(newVarsArr) To UBound(newVarsArr)
            ReDim Preserve staArr(0 To UBound(staArr) + 1) As String
            
            staArr(UBound(staArr)) = _
                newVarsArr(L3) & vbTab & _
                newUnitArr(L3) & vbTab & _
                "Whole foot" & vbTab & _
                ftArr(L2)(1) & vbTab & _
                ftArr(L2)(0) & vbTab & _
                ftArr(L2)(2) & vbTab & _
                ftArr(L2)(3 + L3)
        Next L3
    Next L2
End Function

Function outputFtArr(ByRef ftArr)
'
    Dim L2 As Long
    Dim r1 As Long
    Dim rN As Long
    Dim startCol As Long
    Dim altSwitch As Long
    Dim subModCount As Long
    Dim numOfNewVars As Long
    
    ReDim subModArr(1 To 1) As Long
    
    numOfNewVars = UBound(Split(ftCalcVars, ",")) + 1
    
    ' Count number of rows per block (number of steps per subject).
    ' This allows support for different numbers of steps.
    For L2 = 1 To UBound(ftArr)
        subModCount = subModCount + 1
        
        If L2 = UBound(ftArr) Then
            subModArr(UBound(subModArr)) = subModCount
        ElseIf ftArr(L2, 0) <> ftArr(L2 + 1, 0) Then
            subModArr(UBound(subModArr)) = subModCount
            ReDim Preserve subModArr(1 To UBound(subModArr) + 1) As Long
            
            subModCount = 0
        End If
    Next L2
    
    printArray ftArr, "0,1", 0, 1, 0, "None", "Ft"
    rightBorders "3,7,11," & UBound(ftArr, 2) + 1
    Columns.ColumnWidth = 4
    Columns("A:C").AutoFit
    ActiveWindow.Zoom = 64
    
    ' Implement "watermelon" color scheme.
    Do Until IsNumeric(ftArr(0, startCol))
        startCol = startCol + 1
    Loop
    
    r1 = 2
    For L2 = LBound(subModArr) To UBound(subModArr)
        altSwitch = L2 Mod 2
        
        rN = r1 + subModArr(L2) - 1
        colorForceTime Range(Cells(r1, startCol), Cells(rN, UBound(ftArr, 2) + 1)), altSwitch
        colorForceTime Range(Cells(r1, 4), Cells(rN, 4)), altSwitch ' Column 4 is 1st data column
        colorForceTime Range(Cells(r1, 5), Cells(rN, 7)), altSwitch ' Columns 5-7 are forces
        colorForceTime Range(Cells(r1, 8), Cells(rN, 11)), altSwitch ' Columns 8-11 are times
        
        r1 = r1 + subModArr(L2)
    Next L2
End Function

Function colorForceTime(iRng As Range, ftAlt)
' Cell shading program for force-time data ("watermelon") sheet.
'
    ' Set darkest color shade for most extreme values, 0 (darkest) to 255 (white).
    Const minR As Long = 64
    Const minG As Long = 64
    Const minB As Long = 96
    
    Dim L2 As Long
    Dim indG As Long
    Dim indB As Long
    Dim indR As Long
    Dim rngMin As Long
    Dim rngMed As Long
    Dim rngMax As Long
    Dim iCell As Range
    
    rngMin = maxLngVal
    rngMed = fnArrQuartile(fnMultiTo1DArr(iRng.Value2), 2)
    For Each iCell In iRng
        With iCell
            If LenB(.Value2) <> 0 And IsNumeric(.Value2) Then
                If .Value2 > 0 Then
                    If rngMax < .Value2 Then rngMax = .Value2
                    If rngMin > .Value2 Then rngMin = .Value2
                End If
            End If
        End With
    Next iCell
    
    ReDim colorArr(0 To 1, rngMin To rngMax) As Long
    For L2 = rngMin To rngMax
        If L2 < rngMed Then
            indG = (rngMed - L2) / (rngMed - rngMin) * (255 - minG)
            indB = (rngMed - L2) / (rngMed - rngMin) * (255 - minB)
            colorArr(0, L2) = RGB(255 - indG, 255, 255 - indG)
            colorArr(1, L2) = RGB(255 - indB, 255 - indB, 255)
        Else
            indR = (L2 - rngMed) / (rngMax - rngMed) * (255 - minR)
            colorArr(0, L2) = RGB(255, 255 - indR, 255 - indR)
            colorArr(1, L2) = RGB(255, 255 - indR, 255 - indR)
        End If
    Next L2
    
    For Each iCell In iRng
        With iCell
            If LenB(.Value2) <> 0 And IsNumeric(.Value2) Then
                If .Value2 > 0 Then
                    .Interior.Color = colorArr(ftAlt, CLng(.Value2))
                End If
            End If
        End With
    Next iCell
End Function

Function C04_ftGraphs(ByRef ftArr, Optional rRow = 2, Optional rCol = 3)
' Plot the force-time graphs for each condition.
'
    Const firstDataRow As Long = 2
    Const styleSwitch As Long = 240 ' 240-print, 248-presentation
    Const chartHeight As Long = 225
    Const chartWidth As Long = 500
    Const yAxisMax As Long = 1200
    
    Dim L2 As Long
    Dim L3 As Long
    Dim r1 As Long
    Dim rN As Long
    Dim subModCount As Long
    Dim lettC1 As String
    Dim lettCN As String
    Dim headRow As String
    
    ReDim subModArr(1 To 1) As Long
    ReDim xPos(1 To numOfShoes) As Long
    ReDim yPos(1 To numOfSubjs) As Long
    ReDim dataRng(1 To numOfSubjs, 1 To numOfShoes) As String
    
    ' Find the first and last time columns for the Excel sheet.
    For L2 = LBound(ftArr) To UBound(ftArr)
        If ftArr(0, L2) = 0 Then
            lettC1 = numToLetter(L2 + 1)
            Exit For
        End If
    Next L2
    lettCN = numToLetter(UBound(ftArr) + 1)
    headRow = "," & lettC1 & "1:" & lettCN & "1"
    
    ' Count number of rows in each block.
    ' Unable to handle non-uniform number of shoes yet.
    For L2 = 1 To UBound(ftArr)
        subModCount = subModCount + 1
        
        If L2 = UBound(ftArr) Then
            subModArr(UBound(subModArr)) = subModCount
        ElseIf ftArr(L2, 1) <> ftArr(L2 + 1, 1) Then  ' Next row is a different shoe
            subModArr(UBound(subModArr)) = subModCount
            ReDim Preserve subModArr(1 To UBound(subModArr) + 1) As Long
            subModCount = 0 ' Reset counter
        End If
    Next L2
    
    ' Determine coordinates.
    r1 = firstDataRow
    For L2 = 1 To numOfSubjs
        For L3 = 1 To numOfShoes
            rN = r1 + subModArr(L2) - 1
            dataRng(L2, L3) = lettC1 & r1 & ":" & lettCN & rN
            r1 = r1 + subModArr(L2) ' Next block
        Next L3
        
        yPos(L2) = Cells(rRow, rCol).Top + (L2 - 1) * chartHeight
        If L2 <= numOfShoes Then xPos(L2) = Cells(rRow, rCol).Left + (L2 - 1) * chartWidth
    Next L2
    
    ' Generate charts.
    Sheets.Add , ActiveSheet
    
    For L2 = 1 To numOfSubjs
        For L3 = 1 To numOfShoes
            ActiveSheet.Shapes.AddChart2( _
                styleSwitch, _
                xlXYScatterSmoothNoMarkers, _
                xPos(L3), _
                yPos(L2), _
                chartWidth, _
                chartHeight).Select
            
            With ActiveChart
                .SetSourceData Sheets(ActiveSheet.Previous.Name).Range(dataRng(L2, L3) & headRow)
                .ChartTitle.Text = "S" & subjArr(L2) & " - " & shoeNameArr(L3)
                With .Axes(xlValue)
                    .MinimumScale = 0
                    .MaximumScale = yAxisMax ' Standardize scales
                    .HasTitle = True
                    .AxisTitle.Text = "GRF (N)"
                End With
            End With
        Next L3
    Next L2
    
    ActiveWindow.Zoom = 50
    ActiveSheet.Name = "FtGraphs"
End Function

Function D01_spssTableCreate(ByRef mastArr, destPath)
' Generate spreadsheet for SPSS analysis.
'
    Const lastIdenCol As Long = 3
    
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim numOfRows As Long
    Dim sourceRow As Long
    Dim destCol As Long
    Dim hdrStr As String
    Dim rightBorderCols As String
    Dim pairwiseArray As Variant
    Dim spssArr As Variant
    
    ' Sort shoes into columns.
    pairwiseArray = genArrByShoes(genArrRaw(mastArr))
    
    ' Count number of rows and columns for the new array.
    numOfRows = UBound(pairwiseArray)
    hdrStr = "Variable,Units,Mask,Subject"
    
    For L2 = 1 To numOfShoes
        hdrStr = hdrStr & (",M" & shoeNumArr(L2))
    Next L2
    
    For L2 = 1 To numOfMasks
        For L3 = 1 To numOfShoes
            hdrStr = hdrStr & (",Mask" & L2 & "_" & shoeNumArr(L3))
        Next L3
    Next L2
    
    spssArr = createNewArray(numOfRows, hdrStr) ' Create new array
    
    ' Copy first (4 + numOfShoes) columns.
    For L2 = 1 To numOfRows
        For L3 = 0 To lastIdenCol + numOfShoes
            spssArr(L2, L3) = pairwiseArray(L2, L3)
        Next L3
    Next L2
    
    ' Copy data for two-way repeated measures ANOVA.
    For L2 = 1 To numOfRows
        If spssArr(L2, 2) = maskArr(1) Then ' <~~ Mask 1 (lexicographic order)
            For L3 = 1 To numOfMasks
                For L4 = 1 To numOfShoes
                    sourceRow = L2 + (L3 - 1) * numOfSubjs
                    destCol = L3 * numOfShoes + lastIdenCol + L4
                    spssArr(L2, destCol) = spssArr(sourceRow, lastIdenCol + L4)
                Next L4
            Next L3
        End If
    Next L2
    
    For L2 = 1 + lastIdenCol To 1 + UBound(Split(hdrStr, ",")) Step numOfShoes
        rightBorderCols = rightBorderCols & (L2 & ",")
    Next L2
    rightBorderCols = Left(rightBorderCols, Len(rightBorderCols) - 1)
    
    printArray spssArr, "0,2", 0, 1, 0, "A:D", "SPSS Input"
    rightBorders rightBorderCols
    ActiveWindow.Zoom = 45
    
    With ActiveWorkbook
        .SaveAs destPath & "\spssTable " & Format(Now, "yyyymmdd_hhmm") & ".xlsx"
        .Close False
    End With
End Function

' ================================================================================================
' Phase 2: SPSS output and master table ~~> Results table
' ================================================================================================

Function F01_inputSpssResults()
'
'
    Dim L2 As Long
    Dim L3 As Long
    Dim iVar As Variant
    Dim textList(1 To 2) As Variant
    Dim spssArrs(1 To 3) As Variant
    Dim mastArr As Variant
    
    Dim eSec As Double
    eSec = Timer
    
    MsgBox _
        "Please select 2 files:" & vbCr & _
        "  1.  Master array .txt file" & vbCr & _
        "  2.  SPSS output .txt file" & vbCr & vbCr & _
        "Please do not select more than 2 files.", , "Instructions"
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        If .Show = -1 Then
            For Each iVar In .SelectedItems
                If Right$(iVar, 4) = ".txt" Then
                    L2 = L2 + 1
                    textList(L2) = readTextFile(iVar)
                ElseIf Right$(iVar, 4) = ".xlsx" Then
                    ' Nothing here yet...
                Else
                    MsgBox "File: " & iVar & " is not a supported format. Exiting script."
                    Exit Function
                End If
            Next
        Else
            MsgBox "No files selected. Exiting script."
            Exit Function
        End If
    End With
    
    Application.ScreenUpdating = False
    
    For L2 = 1 To 2
        If Not IsEmpty(textList(L2)) Then
            If InStr(textList(L2)(0), "INFERENTIAL") <> 0 Then
                For L3 = 1 To 3
                    spssArrs(L3) = spssOutToTable(textList(L2), L3)
                Next L3
            Else
                mastArr = linesToTable(textList(L2))
            End If
        End If
    Next L2
    
    If Not IsArray(spssArrs(1)) Then
        MsgBox "SPSS output not found. Exiting script."
        Exit Function
    End If
    
    resetGlobalVariables
    setGlobalVariables , mastArr ' Set numOfShoes, numOfSubjs, numOfSteps, numOfComps
    setVarsUnitMaskArrs mastArr, 0, 1, 2 ' Set varsList, unitList, maskList
        
    E01_masterTableProcess mastArr, spssArrs
    
'    ' Test output.
'    printArray spssArrs(1), "0", 0, 1, 0, "All", "twA"
'    printArray spssArrs(2), "0,2", 1, 1, 0, "All", "swT"
'    printArray spssArrs(3), "0,1", 1, 1, 0, "All", "phT"
    
    Application.ScreenUpdating = True
    
    eSec = Timer - eSec
    Debug.Print "Time elapsed: " & Round(eSec, 5) & " sec"
End Function

Function spssOutToTable(ByRef sourceArr, Optional arrNum = 1) As Variant
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim numOfRows As Long
    Dim numOfCols As Long
    
    For L2 = 0 To UBound(sourceArr)
        If sourceArr(L2) = "Array" & arrNum Then
            numOfRows = sourceArr(L2 + 2)
            numOfCols = sourceArr(L2 + 4) - 1
            
            ReDim newArr(0 To numOfRows, 0 To numOfCols) As Variant
            L2 = L2 + 7 ' <~~ Pre-data (info section) is 6 lines long + 1 separator line
            
            For L3 = 0 To numOfRows
                For L4 = 0 To numOfCols
                    newArr(L3, L4) = sourceArr(L2)
                    L2 = L2 + 1
                Next L4
            Next L3
        End If
    Next L2
    
    spssOutToTable = newArr
End Function

Function arrOfArrTo2dArr(ByRef inArr) As Variant
' Convert an array of 1D arrays (uniform length) to a 2D array.
'
    Dim L2 As Long
    Dim L3 As Long
    
    ReDim outArr(LBound(inArr) To UBound(inArr), LBound(inArr(0)) To UBound(inArr(0))) As Variant
    
    For L2 = LBound(inArr) To UBound(inArr)
'        ReDim Preserve outArr(LBound(inArr) To UBound(inArr), LBound(inArr(L2)) To UBound(inArr(L2))) As Variant
        For L3 = LBound(inArr(L2)) To UBound(inArr(L2))
            outArr(L2, L3) = inArr(L2)(L3)
        Next L3
    Next L2
    
    arrOfArrTo2dArr = outArr
End Function

Function E01_masterTableProcess(ByRef mastArr, Optional spssOut = Empty)
' Array 1: Sort steps into columns, calculate individual mean, SD, and CV.
' Array 2: Sort shoes into columns, calculate pairwise differences.
' Array 3: Group stats (mean/median, SD/IQR, ..., mean within-subject CV, ICC).
' Array 4: Two-way repeated measures ANOVA.
' Array 5: Omnibus tests (one-way repeated measures ANOVA, Friedman's test).
' Array 6: Pairwise tests (paired sample t test, Wilcoxon signed rank test).
' Array 7: Bar chart data (mean/median, 95%CI).
'
    Dim rawArr As Variant
    Dim byShoesArr As Variant
    Dim groupArr As Variant
    Dim twoWayANOVA As Variant
    Dim oneWayOmniArr As Variant
    Dim postHocArr As Variant
    Dim chartDataArr As Variant
    
    Dim normArr() As Boolean
    Dim oneWaySigArr() As Boolean
    Dim twoWayEpsilon() As Long
    Dim oneWayEpsilon() As Long
    
    Dim eSec As Double
    
'    Dim grandArr(1 To 7) As Variant
'
'    grandArr(1) = genArrRaw(mastArr)
'    grandArr(2) = genArrByShoes(grandArr(1))
'    grandArr(3) = genArrGroup(grandArr(1))
'    grandArr(4) = genArrTwoWayANOVA(grandArr(2))
'    grandArr(5) = genArrOmnibus(grandArr(2))
'    grandArr(6) = genArrPostHoc(grandArr(2))
'
'
'    normArr = genNormArr(grandArr(3), 14)
'    grandArr(7) = genArrChartData(grandArr(1), grandArr(1), normArr)
    
    
    
    rawArr = genArrRaw(mastArr)
    If IsEmpty(spssOut) Then
        printArray rawArr, "0,3", 1, 1, 0, "A:E", "Raw"
    Else
        printArray rawArr, "0,3", 0, 1, 0, "A:E", "Raw"
    End If
    
    rightBorders "5,8," & 1 + UBound(rawArr, 2)
    Const r2 As Double = 0.1402
    Const r1 As Double = -7.5265
    Const r0 As Double = 164.93
    If numOfSteps > 10 Then ActiveWindow.Zoom = (r2 * numOfSteps + r1) * numOfSteps + r0
    
    byShoesArr = genArrByShoes(rawArr)
    printArray byShoesArr, "0,2", 1, 1, 0, "A:D", "ByShoes"
    rightBorders "4," & 4 + numOfShoes & "," & 1 + UBound(byShoesArr, 2)
    
    
    eSec = Timer
    
    Dim L2 As Long
    
    groupArr = genArrGroup(rawArr)
    
    If Not IsEmpty(spssOut) Then
        For L2 = 1 To UBound(groupArr)
            groupArr(L2, 14) = spssOut(2)(L2, 4) ' Update p (Shapiro-Wilk) values
        Next L2
    End If
    
    normArr = genNormArr(groupArr, 14)
    printArray groupArr, "0,2", 1, 1, 0, "A:D", "Group"
    rightBorders "4,8,12,15,16," & 1 + UBound(groupArr, 2)
'    rightBorders "4,8,12,15,16," & 15 + numOfSteps
    
    Const g2 As Double = 0.0774
    Const g1 As Double = -4.3492
    Const g0 As Double = 113.13
    ActiveWindow.Zoom = (g2 * numOfSteps + g1) * numOfSteps + g0
    colorSkewKurt Range(Cells(2, 13), Cells(1 + UBound(groupArr), 14)) ' Z(skew) and Z(kurt)
    colorSigVals Range(Cells(2, 15), Cells(1 + UBound(groupArr), 15)) ' p (Shapiro-Wilk)
    colorMeanCVs Range(Cells(2, 16), Cells(1 + UBound(groupArr), 16)) ' mean CVs
    colorMeanICCs Range(Cells(2, 17), Cells(1 + UBound(groupArr), 15 + numOfSteps)) ' mean ICCs
    greyNotNorm normArr, numOfShoes, 5, 8, 9, 12
    
    
    Debug.Print "Group table: " & Round(Timer - eSec, 5) & " sec"
    eSec = Timer
    
    If IsEmpty(spssOut) Then
        twoWayANOVA = genArrTwoWayANOVA(byShoesArr)
    Else
        twoWayANOVA = spssOut(1)
    End If
    twoWayEpsilon = genEpsilonArr(twoWayANOVA, 3, 4)
    printArray twoWayANOVA, "0", 1, 1, 0, "All", "2WayAOV"
    rightBorders "3,6,9,11"
    colorWhiteBlue Range(Cells(2, 11), Cells(1 + UBound(twoWayANOVA), 11)) ' pEtaSq
    colorSigVals Range(Cells(2, 10), Cells(1 + UBound(twoWayANOVA), 10)), True ' p (2-way AOV)
    greyNotEpsilon twoWayEpsilon, 4, 6
    
    oneWayOmniArr = genArrOmnibus(byShoesArr)
    oneWayEpsilon = genEpsilonArr(oneWayOmniArr, 3, 4)
    oneWaySigArr = genSigArr(oneWayOmniArr, normArr, 9, 13)
    printArray oneWayOmniArr, "0,1", 1, 1, 0, "All", "Omnibus"
    rightBorders "3,6,11,15"
    colorWhiteBlue Range(Cells(2, 11), Cells(1 + UBound(oneWayOmniArr), 11)) ' pEtaSq
    colorWhiteBlue Range(Cells(2, 15), Cells(1 + UBound(oneWayOmniArr), 15)) ' Kendall's W
    colorSigVals Range(Cells(2, 10), Cells(1 + UBound(oneWayOmniArr), 10)), True ' p (ANOVA)
    colorSigVals Range(Cells(2, 14), Cells(1 + UBound(oneWayOmniArr), 14)), True ' p (Friedman)
    greyNotNorm normArr, 1, 4, 11, 12, 15
    greyNotEpsilon oneWayEpsilon, 4, 6
    
    
    Debug.Print "Omnibus tables: " & Round(Timer - eSec, 5) & " sec"
    eSec = Timer
    
    
    postHocArr = genArrPostHoc(byShoesArr)
    printArray postHocArr, "0,2", 1, 1, 0, "A:D", "PostHoc"
    rightBorders "4,6,14,22"
    ActiveWindow.Zoom = 80
    Columns(14).NumberFormat = "+0.0%;-0.0%;0.0%" ' % diff
    Columns(22).NumberFormat = "+0.0%;-0.0%;0.0%" ' % diff
    colorRedGreen Range(Cells(2, 13), Cells(1 + UBound(postHocArr), 13)), True ' Cohen's d
    colorRedGreen Range(Cells(2, 21), Cells(1 + UBound(postHocArr), 21)), True ' Cohen's d
    colorRedGreen Range(Cells(2, 14), Cells(1 + UBound(postHocArr), 14)) ' % diff
    colorRedGreen Range(Cells(2, 22), Cells(1 + UBound(postHocArr), 22)) ' % diff
    colorSigVals Range(Cells(2, 12), Cells(1 + UBound(postHocArr), 12)), , True ' p (t test)
    colorSigVals Range(Cells(2, 20), Cells(1 + UBound(postHocArr), 20)), , True ' p (Wilcoxon)
    greyNotNorm normArr, numOfComps, 7, 14, 15, 22
    greyNotSig oneWaySigArr, numOfComps, 7, 22
    
    
    Debug.Print "Post hoc table: " & Round(Timer - eSec, 5) & " sec"
    eSec = Timer
    
    
    chartDataArr = genArrChartData(rawArr, groupArr, normArr)
    printArray chartDataArr, "0,1", 1, 1, 0, "A:D", "ChartData"
    rightBorders "4," & 4 + numOfShoes & "," & 4 + 2 * numOfShoes & "," & 1 + UBound(chartDataArr, 2)
'    rightBorders "4," & 4 + numOfShoes & "," & 4 + 2 * numOfShoes & "," & 4 + 3 * numOfShoes
    colorRedGreen Range(Cells(2, 4), Cells(1 + UBound(chartDataArr), 4)) ' Normality
    
    Debug.Print "Chart data table: " & Round(Timer - eSec, 5) & " sec"
    
    E03_barCharts chartDataArr
End Function

Function genSigArr(ByRef inArr, normArr, pSigCol, npSigCol) As Boolean()
    Dim L2 As Long
    ReDim sigArr(1 To UBound(inArr)) As Boolean
    
    For L2 = 1 To UBound(normArr)
        If normArr(L2) Then
            If inArr(L2, pSigCol) < typeOneErr Then sigArr(L2) = True
        Else
            If inArr(L2, npSigCol) < typeOneErr Then sigArr(L2) = True
        End If
    Next L2
    
    genSigArr = sigArr
End Function

Function genNormArr(ByRef inArr, swCol) As Boolean()
    Dim L2 As Long
    Dim L3 As Long
    Dim nNorm As Long
    ReDim normArr(1 To UBound(inArr) \ numOfShoes) As Boolean
    
    For L2 = 1 To UBound(inArr)
        nNorm = 0
        For L3 = 0 To numOfShoes - 1
            If inArr(L2 + L3, swCol) < typeOneErr Then Exit For
            nNorm = nNorm + 1
        Next L3
        
        L2 = L2 + numOfShoes - 1
        If nNorm = numOfShoes Then normArr(L2 \ numOfShoes) = True
    Next L2
    
    genNormArr = normArr
End Function

Function genEpsilonArr(ByRef inArr, ggCol, hfCol) As Long()
    Dim L2 As Long
    ReDim epsArr(1 To UBound(inArr)) As Long
    
    For L2 = 1 To UBound(inArr)
        If IsNumeric(inArr(L2, ggCol)) And IsNumeric(inArr(L2, hfCol)) Then
            If (CDbl(inArr(L2, ggCol)) + CDbl(inArr(L2, ggCol))) / 2 > 0.75 Then
                epsArr(L2) = 2 ' eHF
            Else
                epsArr(L2) = 1 ' eGG
            End If
        Else
            epsArr(L2) = 3 ' eLB
        End If
    Next L2
    
    genEpsilonArr = epsArr
End Function

Function genArrRaw(ByRef mastArr) As Variant
' Rearrange repeated steps onto a single row.
'
    Const leftCols As String = _
        "Variable," & _
        "Units," & _
        "Mask," & _
        "Model," & _
        "Subject," & _
        "Mean," & _
        "SD," & _
        "CV"
    
    Dim L2 As Long
    Dim lCols As Long
    Dim nStep As Long
    Dim tblHeads As String
    Dim valArr() As Variant
    ReDim rawArr(0) As Variant
    
    lCols = UBound(Split(leftCols, ",")) ' <~~ 7
    
    ' Table headers.
    For L2 = 1 To numOfSteps
        tblHeads = tblHeads & (",Step" & L2)
    Next L2
    tblHeads = leftCols & tblHeads
    rawArr(0) = Split(tblHeads, ",")
    
    ' Scan down master array.
    For L2 = 1 To UBound(mastArr)
        nStep = nStep + 1
        ReDim Preserve valArr(1 To nStep) As Variant
        valArr(nStep) = mastArr(L2, 6)
        
        If L2 = UBound(mastArr) Then
            ' Very last row of table.
            insertRowRawArr mastArr, rawArr, valArr, lCols, L2, nStep
        ElseIf mastArr(L2, 4) <> mastArr(L2 + 1, 4) Then
            ' Last row for current subject.
            insertRowRawArr mastArr, rawArr, valArr, lCols, L2, nStep
            
            ' Re-set for next subject.
            ReDim valArr(1 To 1) As Variant
            nStep = 0
        End If
    Next L2
    
    genArrRaw = arrOfArrTo2dArr(rawArr)
End Function

Function insertRowRawArr(mastArr, rawArr, valArr, lCols, iRow, nStep) As Variant
    Dim L2 As Long
    ReDim newRow(0 To lCols + nStep) As Variant
    ReDim Preserve rawArr(0 To UBound(rawArr) + 1) As Variant
    
    newRow(0) = mastArr(iRow, 0) ' Variable
    newRow(1) = mastArr(iRow, 1) ' Units
    newRow(2) = mastArr(iRow, 2) ' Mask
    newRow(3) = mastArr(iRow, 3) ' Model
    newRow(4) = mastArr(iRow, 4) ' Subject
    newRow(5) = fnArrMean(valArr) ' Mean
    newRow(6) = fnArrSD(valArr) ' SD
    newRow(7) = fnCalcCV(newRow(5), newRow(6)) ' CV
    For L2 = 1 To nStep
        newRow(lCols + L2) = valArr(L2) ' Raw data
    Next L2
    
    rawArr(UBound(rawArr)) = newRow
End Function

Function createNewArray(numOfRows, headerString, Optional deLim = ",") As Variant
' Set the dimensions of new array and input header row.
' Headers would be on row 0.
'
    Dim L2 As Long
    Dim headerArr() As String
    
    headerArr = Split(headerString, deLim)
    
    ReDim newArr(0 To numOfRows, 0 To UBound(headerArr)) As Variant
    
    For L2 = 0 To UBound(headerArr)
        newArr(0, L2) = headerArr(L2)
    Next L2
    
    createNewArray = newArr
End Function

Function genArrByShoes(ByRef rawArr) As Variant
' Rearrange means in rawArr by shoes.
' Calculate pairwise differences.
'
    Const lastIdenCol As Long = 4 ' <~~ First 4 columns for Variable, Units, Mask, and Subject
    
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim readLine As Long
    Dim writeLine As Long
    Dim pwRow As Long
    Dim pwOutCol As Long
    Dim pwInCol2 As Long
    Dim pwInCol1 As Long
    Dim byShoesHeader As String
    Dim byShoesRows As Long
    Dim byShoesArr As Variant
    
    ' Headers for the table.
    byShoesHeader = _
        "Variable," & _
        "Units," & _
        "Mask," & _
        "Subject"
    For L2 = 1 To numOfShoes
        byShoesHeader = byShoesHeader & (",M" & (shoeNumArr(L2)))
    Next L2
    For L2 = 1 To numOfComps
        L4 = pairCompArr2(L2)
        L3 = pairCompArr1(L2)
        byShoesHeader = byShoesHeader & (",M" & (shoeNumArr(L4)) & "-M" & (shoeNumArr(L3)))
    Next L2
    
    ' Count number of rows for the output array.
    byShoesRows = UBound(rawArr) \ numOfShoes
    byShoesArr = createNewArray(byShoesRows, byShoesHeader)
    
    For L2 = 1 To byShoesRows \ numOfSubjs
        readLine = 1 + (L2 - 1) * numOfSubjs * numOfShoes
        writeLine = 1 + (L2 - 1) * numOfSubjs
        
        For L3 = 0 To numOfSubjs - 1
            pwRow = writeLine + L3
            
            ' Copy identifiers.
            byShoesArr(pwRow, 0) = rawArr(readLine + L3, 0) ' Variable
            byShoesArr(pwRow, 1) = rawArr(readLine + L3, 1) ' Units
            byShoesArr(pwRow, 2) = rawArr(readLine + L3, 2) ' Mask
            byShoesArr(pwRow, 3) = rawArr(readLine + L3, 4) ' Subject
            
            ' Copy mean values.
            For L4 = 0 To numOfShoes - 1
                byShoesArr(pwRow, 4 + L4) = rawArr(readLine + L3 + L4 * numOfSubjs, 5) ' Means
            Next L4
            
            ' Calculate pairwise differences.
            For L4 = 0 To numOfComps - 1
                pwOutCol = lastIdenCol + numOfShoes + L4
                pwInCol2 = lastIdenCol - 1 + pairCompArr2(L4 + 1)
                pwInCol1 = lastIdenCol - 1 + pairCompArr1(L4 + 1)
                byShoesArr(pwRow, pwOutCol) = byShoesArr(pwRow, pwInCol2) - byShoesArr(pwRow, pwInCol1)
            Next L4
        Next L3
    Next L2
    
    genArrByShoes = byShoesArr
End Function

Function genArrGroup(ByRef indivArray) As Variant
' Group summary statistics.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim L5 As Long
    Dim inLn As Long
    Dim outLn As Long
    Dim minStepCount As Long
    Dim grpLen As Long
    Dim grpHdr As String
    
    ReDim meanArr(1 To numOfSubjs) As Double
    ReDim cvArr(1 To numOfSubjs) As Variant
    
    Dim statsArr As Variant
    Dim iccOut As Variant
    Dim grpArr As Variant
    Dim hlArr As Variant
    Dim ciArr As Variant
    
    grpHdr = _
        "Variable," & _
        "Units," & _
        "Mask," & _
        "Model," & _
        "gMean," & _
        "gSD," & _
        "95CI lo," & _
        "95CI up," & _
        "gMedian," & _
        "gIQR," & _
        "95CI lo," & _
        "95CI up," & _
        "Z (Skewness)," & _
        "Z (Kurtosis)," & _
        "p (Shapiro-Wilk)," & _
        "CV (mean within-subject)"
    For L2 = 2 To numOfSteps
        grpHdr = grpHdr & (",ICC (" & L2 & " steps)")
    Next L2
    
    ' Count number of rows for the output array.
    grpLen = UBound(indivArray) \ numOfSubjs
    grpArr = createNewArray(grpLen, grpHdr)
    
    For L2 = 1 To grpLen \ numOfShoes
        inLn = 1 + (L2 - 1) * numOfSubjs * numOfShoes
        outLn = 1 + (L2 - 1) * numOfShoes
        
        For L3 = 0 To numOfShoes - 1
            minStepCount = numOfSteps ' <~~ Re-set
            
            For L4 = 0 To numOfSubjs - 1
                For L5 = 0 To numOfSteps - 1
                    If LenB(indivArray(inLn + L3 * numOfSubjs + L4, 8 + L5)) = 0 Then Exit For
                Next L5
                
                ' Find the number of steps that all subjects have for that shoe.
                ' This is to ensure a rectangular array for ICC calculations.
                If minStepCount > L5 Then minStepCount = L5
            Next L4
            
            ' This will be the rectangular array for ICC calculations.
            ReDim stepsArr(1 To numOfSubjs, 1 To minStepCount) As Double
            
            For L4 = 0 To numOfSubjs - 1
                meanArr(1 + L4) = indivArray(inLn + L3 * numOfSubjs + L4, 5) ' Mean
                cvArr(1 + L4) = indivArray(inLn + L3 * numOfSubjs + L4, 7) ' CV
                
                ' Collect data for ICCs here.
                For L5 = 0 To minStepCount - 1
                    stepsArr(1 + L4, 1 + L5) = indivArray(inLn + L3 * numOfSubjs + L4, 8 + L5)
                Next L5
            Next L4
            
            grpArr(outLn + L3, 0) = indivArray(inLn + L3, 0) ' Variable
            grpArr(outLn + L3, 1) = indivArray(inLn + L3, 1) ' Units
            grpArr(outLn + L3, 2) = indivArray(inLn + L3, 2) ' Mask
            grpArr(outLn + L3, 3) = indivArray(inLn + L3 * numOfSubjs, 3) ' Model
            
            ciArr = fnTDist95CI(meanArr)
            hlArr = fnHodgesLehmann(meanArr)
            
            grpArr(outLn + L3, 4) = fnArrMean(meanArr) ' Group mean
            grpArr(outLn + L3, 5) = fnArrSD(meanArr) ' Group SD
            grpArr(outLn + L3, 6) = ciArr(1) ' 95CI lower bound
            grpArr(outLn + L3, 7) = ciArr(2) ' 95CI lower bound
            grpArr(outLn + L3, 8) = hlArr(0) ' Group median
            grpArr(outLn + L3, 9) = hlArr(1) ' Group IQR
            grpArr(outLn + L3, 10) = hlArr(2) ' 95CI lower bound
            grpArr(outLn + L3, 11) = hlArr(3) ' 95CI upper bound
            grpArr(outLn + L3, 12) = fnArrStdSkew(meanArr) ' Standardized skewness
            grpArr(outLn + L3, 13) = fnArrStdKurt(meanArr) ' Standardized kurtosis
            
            grpArr(outLn + L3, 14) = shapiroWilkBelow25(meanArr) ' p (Shapiro-Wilk)
            grpArr(outLn + L3, 15) = fnArrMean(cvArr) ' CV (mean within-subject)
            
            iccOut = meanICCs(stepsArr)
            
            For L4 = 2 To minStepCount
                grpArr(outLn + L3, 14 + L4) = iccOut(L4) ' ICC(3,k) for 2, ..., N steps
            Next L4
        Next L3
    Next L2
    
    genArrGroup = grpArr
End Function

Function genArrTwoWayANOVA(ByRef byShoesArr) As Variant
' Two-way repeated measures ANOVA.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim L5 As Long
    Dim varCount As Long
    Dim readLine As Long
    Dim readBlock As Long
    Dim twoWayAnovaLength As Long
    Dim hdrStr As String
    Dim twoWayAnovaArray As Variant
    Dim outputArr As Variant
    
    ReDim inputArr3D(1 To numOfSubjs, 1 To numOfMasks, 1 To numOfShoes) As Variant
    
    For L2 = LBound(byShoesArr) To UBound(byShoesArr)
        If byShoesArr(L2, 2) = maskArr(1) Then ' Only variables with breakdown by masks
            varCount = varCount + 1
            L2 = L2 + numOfSubjs
        End If
    Next L2
    
    hdrStr = _
        "Effect," & _
        "Variable," & _
        "Units," & _
        "eGG," & _
        "eHF," & _
        "eLB," & _
        "F," & _
        "df1," & _
        "df2," & _
        "p (ANOVA)," & _
        "pEtaSq"
    
    ' Count number of rows for the output array.
    twoWayAnovaLength = varCount * 3 ' <~~ 3 effects
    twoWayAnovaArray = createNewArray(twoWayAnovaLength, hdrStr)
    
    For L2 = 1 To twoWayAnovaLength \ 3
        Do Until byShoesArr(readLine, 2) = maskArr(1)
            readLine = readLine + 1
        Loop
        
        twoWayAnovaArray(L2, 0) = "Mask"
        twoWayAnovaArray(L2 + 1 * varCount, 0) = "Shoe"
        twoWayAnovaArray(L2 + 2 * varCount, 0) = "Mask*Shoe"
        For L3 = 0 To varCount * 2 Step varCount
            twoWayAnovaArray(L2 + L3, 1) = byShoesArr(readLine, 0) ' Variable
            twoWayAnovaArray(L2 + L3, 2) = byShoesArr(readLine, 1) ' Units
        Next L3
        
        For L3 = 1 To numOfMasks
            readBlock = (L3 - 1) * numOfSubjs + readLine
            For L4 = 1 To numOfSubjs
                For L5 = 1 To numOfShoes
                    inputArr3D(L4, L3, L5) = byShoesArr(readBlock - 1 + L4, 3 + L5)
                Next L5
            Next L4
        Next L3
        
        outputArr = rmANOVATwoWay(inputArr3D)
        
        For L3 = 0 To 2
            twoWayAnovaArray(L2 + L3 * varCount, 3) = outputArr(1 + L3, 0) ' eGG
            twoWayAnovaArray(L2 + L3 * varCount, 4) = outputArr(1 + L3, 1) ' eHF
            twoWayAnovaArray(L2 + L3 * varCount, 5) = outputArr(1 + L3, 2) ' eLB
            twoWayAnovaArray(L2 + L3 * varCount, 6) = outputArr(1 + L3, 3) ' F
            twoWayAnovaArray(L2 + L3 * varCount, 7) = outputArr(1 + L3, 4) ' df1
            twoWayAnovaArray(L2 + L3 * varCount, 8) = outputArr(1 + L3, 5) ' df2
            twoWayAnovaArray(L2 + L3 * varCount, 9) = outputArr(1 + L3, 6) ' p (ANOVA)
            twoWayAnovaArray(L2 + L3 * varCount, 10) = outputArr(1 + L3, 7) ' pEtaSq
        Next L3
        
        readLine = readLine + numOfSubjs
    Next L2
    
    genArrTwoWayANOVA = twoWayAnovaArray
End Function

Function genArrOmnibus(ByRef pairwiseArray) As Variant
'
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim readLine As Long
    Dim omnibusLength As Long
    Dim hdrStr As String
    
    Dim rmAOVOutputArr As Variant
    Dim friedmanOutputArr As Variant
    Dim omnibusArray As Variant
    
    ReDim byShoesArr(0 To numOfSubjs - 1, 0 To numOfShoes - 1) As Variant
    
    hdrStr = _
        "Variable," & _
        "Units," & _
        "Mask," & _
        "eGG," & _
        "eHF," & _
        "eLB," & _
        "F," & _
        "df1," & _
        "df2," & _
        "p (ANOVA)," & _
        "pEtaSq," & _
        "chiSq," & _
        "df," & _
        "p (Friedman)," & _
        "W (Kendall)"
    
    ' Count number of rows for the output array.
    omnibusLength = countNewRows(pairwiseArray, 0, 2) ' <~~ Different variable-mask combo
    omnibusArray = createNewArray(omnibusLength, hdrStr)
    
    For L2 = 1 To omnibusLength
        readLine = 1 + (L2 - 1) * numOfSubjs
        
        omnibusArray(L2, 0) = pairwiseArray(readLine, 0) ' Variable
        omnibusArray(L2, 1) = pairwiseArray(readLine, 1) ' Units
        omnibusArray(L2, 2) = pairwiseArray(readLine, 2) ' Mask
        
        For L3 = 0 To numOfSubjs - 1
            For L4 = 0 To numOfShoes - 1
                byShoesArr(L3, L4) = pairwiseArray(readLine + L3, 4 + L4)
            Next L4
        Next L3
        
        rmAOVOutputArr = rmANOVA(byShoesArr)
        friedmanOutputArr = npTestFriedman(byShoesArr)
        
        omnibusArray(L2, 3) = rmAOVOutputArr(0) ' eGG
        omnibusArray(L2, 4) = rmAOVOutputArr(1) ' eHF
        omnibusArray(L2, 5) = rmAOVOutputArr(2) ' eLB
        omnibusArray(L2, 6) = rmAOVOutputArr(3) ' F
        omnibusArray(L2, 7) = rmAOVOutputArr(4) ' df1
        omnibusArray(L2, 8) = rmAOVOutputArr(5) ' df2
        omnibusArray(L2, 9) = rmAOVOutputArr(6) ' p (ANOVA)
        omnibusArray(L2, 10) = rmAOVOutputArr(7) ' pEtaSq
        
        omnibusArray(L2, 11) = friedmanOutputArr(0) ' chiSq
        omnibusArray(L2, 12) = friedmanOutputArr(1) ' df
        omnibusArray(L2, 13) = friedmanOutputArr(2) ' p (Friedman)
        omnibusArray(L2, 14) = friedmanOutputArr(3) ' Kendall's W
    Next L2
    
    genArrOmnibus = omnibusArray
End Function

Function genArrPostHoc(ByRef pwArr) As Variant
' Pairwise post hoc comparisons.
'
    Const pwLastIdenCol As Long = 4
    
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim readPwLine As Long
    Dim postHocLength As Long
    Dim hdrStr As String
    Dim outputArrPNP As Variant
    Dim postHocArr As Variant
    
    ReDim diffArr(1 To numOfSubjs) As Double
    ReDim shoeArr1(1 To numOfSubjs) As Double
    ReDim shoeArr2(1 To numOfSubjs) As Double
    
    hdrStr = _
        "Variable," & _
        "Units," & _
        "Mask," & _
        "Comparison," & _
        "Correlation," & _
        "," & _
        "dMean," & _
        "dSD," & _
        "95CI lo," & _
        "95CI up," & _
        "t," & _
        "p (Paired t)," & _
        "d," & _
        "% diff.," & _
        "dMedian," & _
        "dIQR," & _
        "95CI lo," & _
        "95CI up," & _
        "Z," & _
        "p (Wilcoxon)," & _
        "d," & _
        "% diff."
    
    ' Count number of rows for the output array.
    postHocLength = countNewRows(pwArr, 0, 2) * numOfComps ' <~~ Different variable-mask combo
    postHocArr = createNewArray(postHocLength, hdrStr)
    
    For L2 = 1 To postHocLength Step numOfComps
        readPwLine = 1 + (L2 - 1) * numOfSubjs \ numOfComps
        
        For L3 = 0 To numOfComps - 1
            For L4 = 0 To numOfSubjs - 1
                shoeArr1(1 + L4) = pwArr(readPwLine + L4, 3 + pairCompArr1(1 + L3))
                shoeArr2(1 + L4) = pwArr(readPwLine + L4, 3 + pairCompArr2(1 + L3))
                diffArr(1 + L4) = pwArr(readPwLine + L4, pwLastIdenCol + numOfShoes + L3)
            Next L4
            
            postHocArr(L2 + L3, 0) = pwArr(1 + readPwLine, 0) ' Variable
            postHocArr(L2 + L3, 1) = pwArr(1 + readPwLine, 1) ' Units
            postHocArr(L2 + L3, 2) = pwArr(1 + readPwLine, 2) ' Mask
            postHocArr(L2 + L3, 3) = pwArr(0, pwLastIdenCol + numOfShoes + L3) ' Comparison
            
            outputArrPNP = calcPostHocStats(diffArr, shoeArr1, shoeArr2)
            
            postHocArr(L2 + L3, 4) = outputArrPNP(0) ' r
            For L4 = 1 To 16
                postHocArr(L2 + L3, 5 + L4) = outputArrPNP(L4)
            Next L4
        Next L3
    Next L2
    
    genArrPostHoc = postHocArr
End Function

Function calcPostHocStats(ByRef diffArr, shoeArr1, shoeArr2) As Variant
' Specialized function for calculating post hoc test statistics.
'
    Dim L2 As Long
    Dim arrLen As Long
    Dim diffMean As Double
    Dim diffSD As Double
    Dim tStat As Double
    Dim pVal As Double
    Dim gMean1 As Double
    Dim gMedian1 As Double
    Dim rCorrel As Variant
    Dim diff95CIArr As Variant
    Dim hlArr As Variant
    Dim wilcoxArr As Variant
    Dim outArr(0 To 16) As Variant
    
    arrLen = UBound(diffArr) - LBound(diffArr) + 1
    rCorrel = fnCorrel(shoeArr1, shoeArr2)
    diffMean = fnArrMean(diffArr)
    diffSD = fnArrSD(diffArr)
    diff95CIArr = fnTDist95CI(diffArr, typeOneErr / numOfComps)
    hlArr = fnHodgesLehmann(diffArr)
    wilcoxArr = npTestWilcoxonSR(diffArr)
    gMean1 = fnArrMean(shoeArr1)
    gMedian1 = fnHodgesLehmann(shoeArr1)(0)
    
    ' Set default values to "NA".
    For L2 = LBound(outArr) To UBound(outArr)
        outArr(L2) = "NA"
    Next L2
    
    outArr(0) = rCorrel ' Correlation
    outArr(1) = diffMean ' Group mean
    outArr(2) = diffSD ' Group SD
    If diffSD <> 0 Then
        tStat = diffMean / (diffSD / Sqr(arrLen)) ' t = diffMean / SE
        pVal = Application.TDist(Abs(tStat), arrLen - 1, 2) * numOfComps ' Bonf. p
        If pVal > 1 Then pVal = 1
        
        outArr(3) = diff95CIArr(1) ' 95CI lower bound
        outArr(4) = diff95CIArr(2) ' 95CI upper bound
        outArr(5) = tStat ' t
        outArr(6) = pVal ' Bonf. p
        outArr(7) = fnCohensD(diffMean, diffSD, rCorrel) ' Cohen's d (parametric)
    End If
    If gMean1 <> 0 Then outArr(8) = diffMean / gMean1 ' % difference (parametric)
    
    ' Nonparametric section.
    outArr(9) = hlArr(0) ' Group median
    outArr(10) = hlArr(1) ' Group IQR
    outArr(11) = hlArr(2) ' 95CI lower bound
    outArr(12) = hlArr(3) ' 95CI upper bound
    outArr(13) = wilcoxArr(0) ' Z (Wilcoxon)
    outArr(14) = wilcoxArr(1) ' p (Wilcoxon)
    If hlArr(1) <> 0 Then outArr(15) = fnCohensD(hlArr(0), hlArr(1), rCorrel) ' Cohen's d (np)
    If gMedian1 <> 0 Then outArr(16) = hlArr(0) / gMedian1 ' % difference (nonparametric)
    
    calcPostHocStats = outArr
End Function

Function genArrChartData(ByRef indivArr, grpArr, normArr) As Variant
' Copy data for bar charts.
'
    Const lastIdenCol As Long = 4
    
    Dim L2 As Long
    Dim L3 As Long
    Dim readLine As Long
    Dim chartsLength As Long
    Dim errLB As Double
    Dim errUB As Double
    Dim hdrStr As String
    Dim hdrStr2 As String
    Dim hdrStr3 As String
    Dim hdrStr4 As String
    Dim hdrStr5 As String
    Dim chartsArray As Variant
    
    hdrStr = _
        "Variable," & _
        "Units," & _
        "Mask," & _
        "Normality"
    For L2 = 1 To numOfShoes
        hdrStr2 = hdrStr2 & (",M" & shoeNumArr(L2))
        hdrStr3 = hdrStr3 & (",errLo" & shoeNumArr(L2))
        hdrStr4 = hdrStr4 & (",errUp" & shoeNumArr(L2))
        hdrStr5 = hdrStr5 & (",sig" & shoeNumArr(L2))
    Next L2
    hdrStr = hdrStr & hdrStr2 & hdrStr3 & hdrStr4 & hdrStr5
    
    ' Count number of rows for the output array.
    chartsLength = countNewRows(indivArr, 0, 2) ' <~~ Different variable-mask combo
    chartsArray = createNewArray(chartsLength, hdrStr)
    
    For L2 = 1 To chartsLength
        readLine = 1 + (L2 - 1) * numOfShoes
        
        chartsArray(L2, 0) = grpArr(readLine, 0) ' Variable
        chartsArray(L2, 1) = grpArr(readLine, 1) ' Units
        chartsArray(L2, 2) = grpArr(readLine, 2) ' Mask
        chartsArray(L2, 3) = normArr(L2) ' Normality
        
        ' Copy data from group summary stats table.
        For L3 = 0 To numOfShoes - 1
            If normArr(L2) Then
                errLB = grpArr(readLine + L3, 4) - grpArr(readLine + L3, 6) ' Med - 95CI lo
                errUB = grpArr(readLine + L3, 7) - grpArr(readLine + L3, 4) ' 95CI up - Med
                
                chartsArray(L2, lastIdenCol + L3) = grpArr(readLine + L3, 4) ' Group mean
                chartsArray(L2, lastIdenCol + numOfShoes + L3) = errLB
                chartsArray(L2, lastIdenCol + 2 * numOfShoes + L3) = errUB
            Else
                errLB = grpArr(readLine + L3, 8) - grpArr(readLine + L3, 10) ' Error lo
                errUB = grpArr(readLine + L3, 11) - grpArr(readLine + L3, 8) ' Error up
                
                chartsArray(L2, lastIdenCol + L3) = grpArr(readLine + L3, 8) ' Group median
                chartsArray(L2, lastIdenCol + numOfShoes + L3) = errLB
                chartsArray(L2, lastIdenCol + 2 * numOfShoes + L3) = errUB
            End If
        Next L3
    Next L2
    
    
    ' Testing
    chartsArray(1, lastIdenCol + 3 * numOfShoes + 0) = "A"
    chartsArray(1, lastIdenCol + 3 * numOfShoes + 1) = "B"
    chartsArray(1, lastIdenCol + 3 * numOfShoes + 2) = "C"
    chartsArray(1, lastIdenCol + 3 * numOfShoes + 3) = "D"
    
    chartsArray(2, lastIdenCol + 3 * numOfShoes + 0) = "E"
    chartsArray(2, lastIdenCol + 3 * numOfShoes + 1) = "F"
    chartsArray(2, lastIdenCol + 3 * numOfShoes + 2) = "G"
    chartsArray(2, lastIdenCol + 3 * numOfShoes + 3) = "H"
    
    chartsArray(48, lastIdenCol + 3 * numOfShoes + 0) = "A"
    chartsArray(48, lastIdenCol + 3 * numOfShoes + 1) = "B"
    chartsArray(48, lastIdenCol + 3 * numOfShoes + 2) = "C"
    chartsArray(48, lastIdenCol + 3 * numOfShoes + 3) = "D"
    
    chartsArray(49, lastIdenCol + 3 * numOfShoes + 0) = "E"
    chartsArray(49, lastIdenCol + 3 * numOfShoes + 1) = "F"
    chartsArray(49, lastIdenCol + 3 * numOfShoes + 2) = "G"
    chartsArray(49, lastIdenCol + 3 * numOfShoes + 3) = "H"
    
    genArrChartData = chartsArray
End Function

Function countNewRows(ByRef inputArr, iCol1, Optional iCol2 = -1) As Long
' Returns the number of rows needed for the new array.
'
    Dim L2 As Long
    Dim s2 As String
    Dim s3 As String
    
    If iCol2 = -1 Then iCol2 = iCol1
    
    s2 = inputArr(0, iCol1)
    s3 = inputArr(0, iCol2)
    For L2 = 1 To UBound(inputArr, 1)
        If LenB(inputArr(L2, iCol1)) <> 0 Then
            If inputArr(L2, iCol1) <> s2 Or inputArr(L2, iCol2) <> s3 Then
                s2 = inputArr(L2, iCol1)
                s3 = inputArr(L2, iCol2)
                countNewRows = countNewRows + 1 ' <~~ +1 for different condition
            End If
        End If
    Next L2
End Function

Function greyNotNorm(ByRef normArr, bHeight, pC1, pCN, npC1, npCN)
' Grey out parametric or nonparametric sections of array based on normality.
'
    Dim L2 As Long
    Dim c1 As Long
    Dim cN As Long
    
    For L2 = LBound(normArr) To UBound(normArr)
        If normArr(L2) Then
            c1 = npC1
            cN = npCN
        Else
            c1 = pC1
            cN = pCN
        End If
        
        With Range(Cells(2, c1), Cells(1 + bHeight, cN)).Offset((L2 - 1) * bHeight, 0)
            .Font.Color = RGB(225, 225, 225) ' Grey
            .Interior.colorIndex = 0 ' No fill
        End With
    Next L2
End Function

Function greyNotSig(ByRef sigArr, numOfRows, colFirst, colLast)
' Grey out sections of array based on omnibus tests.
'
    Dim L2 As Long
    Dim startRow As Long
    Dim endRow As Long
    
    For L2 = LBound(sigArr) To UBound(sigArr)
        startRow = 2 + (L2 - 1) * numOfRows
        endRow = 1 + L2 * numOfRows
        
        If sigArr(L2) = False Then
            Range(Cells(startRow, colFirst), Cells(endRow, colLast)).Font.Color = RGB(225, 225, 225)
            'Range(Cells(startRow, colFirst), Cells(endRow, colLast)).Interior.ColorIndex = 0
        End If
    Next L2
End Function

Function greyNotEpsilon(ByRef epsilonArr, colFirst, colLast)
' Grey out sections of array based on omnibus tests.
'
    Dim L2 As Long
    
    For L2 = LBound(epsilonArr) To UBound(epsilonArr)
        If epsilonArr(L2) = 1 Then
            Cells(L2 + 1, colFirst + 1).Font.Color = RGB(225, 225, 225)
            Cells(L2 + 1, colFirst + 2).Font.Color = RGB(225, 225, 225)
        ElseIf epsilonArr(L2) = 2 Then
            Cells(L2 + 1, colFirst).Font.Color = RGB(225, 225, 225)
            Cells(L2 + 1, colFirst + 2).Font.Color = RGB(225, 225, 225)
        ElseIf epsilonArr(L2) = 3 Then
            Cells(L2 + 1, colFirst).Font.Color = RGB(225, 225, 225)
            Cells(L2 + 1, colFirst + 1).Font.Color = RGB(225, 225, 225)
        End If
    Next L2
End Function

'
'
'

Function colorSkewKurt(iRng As Range)
' Highlight Z scores of skewnesses and kurtoses more than 1.96 standard errors from 0.
'
    ' Set darkest color shade for most extreme values, 0 (darkest) to 255 (white).
    Const minR As Long = 64
    Const mP As Long = 100 ' Multiplier
    
    Dim L2 As Long
    Dim indR As Long
    Dim rngMaxAbs As Long
    Dim SE196 As Double
    Dim iCell As Range
    
    ' Find the max absolute value during first pass.
    For Each iCell In iRng
        With iCell
            If IsNumeric(.Value2) Then
                If LenB(.Value2) <> 0 Then
                    If rngMaxAbs < Abs(.Value2 * mP) Then rngMaxAbs = Abs(.Value2 * mP)
                End If
            End If
        End With
    Next iCell
    
    ReDim colorArr(0 To rngMaxAbs) As Long
    For L2 = 0 To rngMaxAbs
        indR = L2 / rngMaxAbs * (255 - minR)
        colorArr(L2) = RGB(255, 255 - indR, 255 - indR)
    Next L2
    
    SE196 = stdNormCdfInv(0.975) ' 1.96 standard errors
    
    For Each iCell In iRng
        With iCell
            If IsNumeric(.Value2) Then
                If LenB(.Value2) <> 0 Then
                    If Abs(.Value2) >= SE196 Then
                        .Interior.Color = colorArr(CLng(Abs(.Value2 * mP)))
                    End If
                End If
            End If
        End With
    Next iCell
End Function

Function colorSigVals(iRng As Range, Optional Omni = False, Optional pHoc = False)
' Highlight p values less than 0.05.
'
    ' Set darkest color shade for most extreme values, 0 (darkest) to 255 (white).
    Const minY As Long = 96
    Const mP As Long = 10000 ' Multiplier
    
    Dim L2 As Long
    Dim indY As Long
    Dim iCell As Range
    Dim colorArr(0 To 0.1 * mP) As Long
    
    ' Define for 0% to 10%.
    For L2 = 0 To 0.1 * mP
        indY = L2 / (0.1 * mP) * (255 - minY)
        colorArr(L2) = RGB(255, 255, minY + indY)
    Next L2
    
    For Each iCell In iRng
        With iCell
            If LenB(.Value2) <> 0 And IsNumeric(.Value2) Then
                If .Value2 < 0.05 Then
                    .Interior.Color = colorArr(CLng(.Value2 * mP))
                Else
                    If Omni Then
                        .Offset(0, 1).Interior.colorIndex = 0 ' Effect size estimate
                    ElseIf pHoc Then
                        .Offset(0, -5).Interior.colorIndex = 0 ' dMean or dMedian
                        .Offset(0, -3).Interior.colorIndex = 0 ' 95CI lower bound
                        .Offset(0, -2).Interior.colorIndex = 0 ' 95CI upper bound
                        .Offset(0, 1).Interior.colorIndex = 0 ' Cohen's d
                        .Offset(0, 2).Interior.colorIndex = 0 ' % difference
                    End If
                End If
            End If
        End With
    Next iCell
End Function

Function colorMeanCVs(iRng As Range)
' Highlight mean within-subject CVs greater than 17% (reference required).
'
    ' Set darkest color shade for most extreme values, 0 (darkest) to 255 (white).
    Const minR As Long = 64
    Const mP As Long = 1000 ' Multiplier
    
    Dim L2 As Long
    Dim indR As Long
    Dim rngMax As Long
    Dim iCell As Range
    
    For Each iCell In iRng
        With iCell
            If IsNumeric(.Value2) Then
                If LenB(.Value2) <> 0 Then
                    If rngMax < .Value2 * mP Then rngMax = .Value2 * mP
                End If
            End If
        End With
    Next iCell
    
    ReDim colorArr(0 To rngMax) As Long
    For L2 = 0 To rngMax
        indR = L2 / rngMax * (255 - minR)
        colorArr(L2) = RGB(255, 255 - indR, 255 - indR)
    Next L2
    
    For Each iCell In iRng
        With iCell
            If IsNumeric(.Value2) Then
                If LenB(.Value2) <> 0 Then
                    If .Value2 >= 0.17 Then
                        .Interior.Color = colorArr(CLng(.Value2 * mP))
                    End If
                End If
            End If
        End With
    Next iCell
End Function

Function colorMeanICCs(iRng As Range)
' Highlight mean intraclass correlation coefficients (ICCs).
'
' Reference:
'     Portney, L. G., & Watkins, M. P. (2009).
'     Foundations of clinical research: Applications to practice.
'     Upper Saddle River, N.J: Pearson/Prentice Hall.
'         >0.90 ~~> 'Reasonable for clinical measurements'.
'         0.75 ~~> Good.
'
    ' Set darkest color shade for most extreme values, 0 (darkest) to 255 (white).
    Const minY As Long = 96
    Const minG As Long = 64
    Const maxG As Long = 224
    Const mP As Long = 1000 ' Multiplier
    
    Dim L2 As Long
    Dim indY As Long
    Dim indG As Long
    Dim colorArr(0 To mP) As Long
    Dim iCell As Range
    
    For L2 = 0 To mP
        If L2 < 0.9 * mP Then
            indY = L2 / (0.9 * mP) * (255 - minY)
            colorArr(L2) = RGB(255, 255, minY + indY)
        Else
            indG = (L2 - 0.9 * mP) / (mP - 0.9 * mP) * (maxG - minG)
            colorArr(L2) = RGB(maxG - indG, 255, maxG - indG)
        End If
    Next L2
    
    For Each iCell In iRng
        With iCell
            If LenB(.Value2) <> 0 And IsNumeric(.Value2) Then
                If .Value2 >= 0 And .Value2 <= 1 Then
                    .Interior.Color = colorArr(CLng(.Value2 * mP))
                Else
                    .Interior.Color = RGB(255, 96, 96)
                End If
            End If
        End With
    Next iCell
End Function

Function colorWhiteBlue(iRng As Range)
' Highlight non-negative numbers.
'
    ' Set darkest color shade for most extreme values, 0 (darkest) to 255 (white).
    Const minB As Long = 96
    Const mP As Long = 1000 ' Multiplier
    
    Dim L2 As Long
    Dim indB As Long
    Dim rngMax As Long
    Dim iCell As Range
    
    For Each iCell In iRng
        With iCell
            If LenB(.Value2) <> 0 And IsNumeric(.Value2) Then
                If rngMax < .Value2 * mP Then rngMax = .Value2 * mP
            End If
        End With
    Next iCell
    
    ReDim colorArr(0 To rngMax) As Long
    
    For L2 = 0 To rngMax
        indB = L2 / rngMax * (255 - minB)
        colorArr(L2) = RGB(255 - indB, 255 - indB, 255)
    Next L2
    
    For Each iCell In iRng
        With iCell
            If LenB(.Value2) <> 0 And IsNumeric(.Value2) Then
                .Interior.Color = colorArr(CLng(.Value2 * mP))
            End If
        End With
    Next iCell
End Function

Function colorRedGreen(iRng As Range, Optional cD = False)
' Highlight real numbers.
' For Cohen's d, highlight mean/median and 95%CI as well.
'
    ' Set darkest color shade for most extreme values, 0 (darkest) to 255 (white).
    Const minR As Long = 64
    Const minG As Long = 64
    Const mP As Long = 1000 ' Multiplier
    
    Dim L2 As Long
    Dim indR As Long
    Dim indG As Long
    Dim rngMin As Long
    Dim rngMax As Long
    Dim iCell As Range
    
    rngMin = maxLngVal
    For Each iCell In iRng
        With iCell
            If IsNumeric(.Value2) Then
                If LenB(.Value2) <> 0 Then
                    If rngMax < .Value2 * mP Then rngMax = .Value2 * mP
                    If rngMin > .Value2 * mP Then rngMin = .Value2 * mP
                End If
            End If
        End With
    Next iCell
    
    ReDim colorArr(rngMin To rngMax) As Long
    For L2 = rngMin To rngMax
        If L2 < 0 Then
            indG = L2 / rngMin * (255 - minG)
            colorArr(L2) = RGB(255 - indG, 255, 255 - indG)
        ElseIf L2 = 0 Then
            colorArr(L2) = RGB(255, 255, 255)
        Else
            indR = L2 / rngMax * (255 - minR)
            colorArr(L2) = RGB(255, 255 - indR, 255 - indR)
        End If
    Next L2
    
    For Each iCell In iRng
        With iCell
            If IsNumeric(.Value2) Then
                If LenB(.Value2) <> 0 Then
                    .Interior.Color = colorArr(CLng(.Value2 * mP))
                    
                    If cD Then
                        .Offset(0, -6).Interior.Color = colorArr(CLng(.Value2 * mP)) ' dMean or dMedian
                        .Offset(0, -4).Interior.Color = colorArr(CLng(.Value2 * mP)) ' 95CI lower bound
                        .Offset(0, -3).Interior.Color = colorArr(CLng(.Value2 * mP)) ' 95CI upper bound
                    End If
                End If
            End If
        End With
    Next iCell
End Function

Function E03_barCharts(ByRef chartArr, Optional rRow = 2, Optional rCol = 2)
'
    Const styleSwitch As Long = 201 ' 201-print, 209-presentation
    Const fixedCols As Long = 4
    Const xWidth As Long = 1100
    Const yHeight As Long = 500
    
    Dim L2 As Long
    Dim L3 As Long
    Dim L4 As Long
    Dim xRef As Long
    Dim yRef As Long
    Dim nVars As Long
    Dim checkVar As String
    Dim checkUnit As String
    Dim c1(1 To 4) As String
    Dim cN(1 To 4) As String
    
    ReDim r1(1 To 1) As Long
    ReDim rN(1 To 1) As Long
    ReDim yPos(1 To 1) As Long
    ReDim yTitle(1 To 1) As String
    
    xRef = Cells(rRow, rCol).Left
    yRef = Cells(rRow, rCol).Top
    
    ' Set column boundaries.
    For L2 = 1 To 4
        c1(L2) = numToLetter(1 + fixedCols + (L2 - 1) * numOfShoes)
        cN(L2) = numToLetter(fixedCols + L2 * numOfShoes)
    Next L2
    
    ' Calculate r1 and nBars values.
    For L2 = 1 To UBound(chartArr)
        If checkVar <> chartArr(L2, 0) Or checkUnit <> chartArr(L2, 1) Then
            checkVar = chartArr(L2, 0)
            checkUnit = chartArr(L2, 1)
            
            nVars = nVars + 1
            
            r1(nVars) = L2 + 1
            ReDim Preserve r1(1 To nVars + 1) As Long
            
            yPos(nVars) = yRef + yHeight * (nVars - 1)
            ReDim Preserve yPos(1 To nVars + 1) As Long
            
            yTitle(nVars) = varsArr(nVars) & " " & unitArr(nVars)
            ReDim Preserve yTitle(1 To nVars + 1) As String
            
            If nVars > 1 Then
                rN(nVars - 1) = r1(nVars) - 1
                ReDim Preserve rN(1 To nVars) As Long
            End If
        End If
    Next L2
    rN(UBound(rN)) = L2
    
    ReDim rngStrArr(1 To nVars) As String
    ReDim errLo(1 To nVars, 1 To numOfShoes) As Range
    ReDim errUp(1 To nVars, 1 To numOfShoes) As Range
    ReDim dLab(1 To nVars, 1 To numOfShoes) As Range
    ReDim dLabStr(1 To nVars, 1 To numOfShoes) As String
    
    ' Variables for chart data.
    For L2 = 1 To nVars
        ' Data range reference.
        rngStrArr(L2) = _
            c1(1) & r1(L2) & ":" & cN(1) & rN(L2) & "," & _
            c1(1) & "1:" & cN(1) & "1,C" & r1(L2) & ":C" & rN(L2) & ",C1"
        
        If rN(L2) - r1(L2) + 1 <> 1 Then
            ' Is a cluster.
            For L3 = 1 To numOfShoes
                Set errLo(L2, L3) = Range(c1(2) & r1(L2) & ":" & c1(2) & rN(L2)).Offset(0, L3 - 1)
                Set errUp(L2, L3) = Range(c1(3) & r1(L2) & ":" & c1(3) & rN(L2)).Offset(0, L3 - 1)
                Set dLab(L2, L3) = Range(c1(4) & r1(L2) & ":" & cN(4) & rN(L2)).Offset(0, L3 - 1)
                dLabStr(L2, L3) = "='" & ActiveSheet.Name & "'!" & dLab(L2, L3).Address
            Next L3
        Else
            ' Not a cluster.
            Set errLo(L2, 1) = Range(c1(2) & r1(L2) & ":" & cN(2) & r1(L2))
            Set errUp(L2, 1) = Range(c1(3) & r1(L2) & ":" & cN(3) & r1(L2))
            Set dLab(L2, 1) = Range(c1(4) & r1(L2) & ":" & cN(4) & rN(L2))
            dLabStr(L2, 1) = "='" & ActiveSheet.Name & "'!" & dLab(L2, 1).Address
        End If
    Next L2
    
    ' Generate charts.
    Sheets.Add , Sheets(Sheets.Count)
    
    For L2 = 1 To nVars
        ActiveSheet.Shapes.AddChart2( _
            styleSwitch, _
            xlColumnClustered, _
            xRef, _
            yPos(L2), _
            xWidth - 5, _
            yHeight - 5).Select
        
        With ActiveChart
            .SetSourceData Source:=Sheets(ActiveSheet.Previous.Name).Range(rngStrArr(L2))
            .ApplyLayout 5
            .ChartTitle.Text = yTitle(L2)
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).AxisTitle.Delete
            
            If rN(L2) - r1(L2) + 1 <> 1 Then
                For L3 = 1 To numOfShoes
                    With .SeriesCollection(L3)
                        .HasErrorBars = True
                        .ErrorBar xlY, xlBoth, xlCustom, errUp(L2, L3), errLo(L2, L3)
                        
                        .HasDataLabels = True
                        With .DataLabels
                            .ShowValue = False
                            .ShowRange = True
                            With .Format.TextFrame2
                                .TextRange.InsertChartField msoChartFieldRange, dLabStr(L2, L3), 0
                                .VerticalAnchor = msoAnchorTop
                            End With
                            .Font.Size = 14
                        End With
                        
                        For L4 = 1 To rN(L2) - r1(L2) + 1
                            .DataLabels(L4).Format.TextFrame2.TextRange.InsertAfter Chr(13)
                        Next L4
                    End With
                Next L3
            Else
                With .SeriesCollection(1)
                    .HasErrorBars = True
                    .ErrorBar xlY, xlBoth, xlCustom, errUp(L2, 1), errLo(L2, 1)
                    
                    .HasDataLabels = True
                    With .DataLabels
                        .ShowValue = False
                        .ShowRange = True
                        With .Format.TextFrame2
                            .TextRange.InsertChartField msoChartFieldRange, dLabStr(L2, 1), 0
                            .VerticalAnchor = msoAnchorTop
                        End With
                        .Font.Size = 14
                    End With
                End With
            End If
        End With
    Next L2
    
    ActiveSheet.Name = "BarCharts"
    ActiveWindow.Zoom = 80
    Cells(1, 1).Select
End Function

' ================================================================================================
' Common Functions
' ================================================================================================

Function readTextFile(filePath) As String()
' Read specified text file line-by-line into a 1D array of strings.
'
    Dim fileNum As Long
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
        readTextFile = Split(Input$(LOF(fileNum), #fileNum), vbNewLine) ' MUST be vbNewLine
    Close #fileNum
End Function

Function linesToTable(ByRef inputArr, Optional deLim = vbTab) As Variant
' Convert a 1D array of delimited (tab, comma, etc.) strings to a 2D array.
' Each string should have the same number of delimiters, or be completely empty.
'
    Dim L2 As Long
    Dim L3 As Long
    Dim lastCol As Long
    Dim tempRow As Variant
    
    lastCol = UBound(Split(inputArr(0), deLim)) ' Index of last column of table
    ReDim outputArr(LBound(inputArr) To UBound(inputArr), 0 To lastCol) As Variant
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        If LenB(inputArr(L2)) <> 0 Then ' <~~ Not an empty row
            tempRow = Split(inputArr(L2), deLim)
            For L3 = LBound(tempRow) To UBound(tempRow)
                outputArr(L2, L3) = tempRow(L3)
            Next L3
        End If
    Next L2
    
    linesToTable = outputArr
End Function

Function fnSwapValues(ByRef Variable1, ByRef Variable2)
    Dim tempVal As Variant
    tempVal = Variable1
    Variable1 = Variable2
    Variable2 = tempVal
End Function

Function fnPartition(ByRef inputArr, Lo As Long, Hi As Long, pivotIndex) As Long
' Partition function for QuickSelect algorithm.
'
    Dim L2 As Long
    Dim tempIndex As Long
    
    ' Store pivot value as the last element in array.
    fnSwapValues inputArr(pivotIndex), inputArr(Hi)
    
    ' Move all values smaller than the pivot to the left.
    tempIndex = Lo
    For L2 = Lo To Hi - 1
        If inputArr(L2) < inputArr(Hi) Then
            fnSwapValues inputArr(tempIndex), inputArr(L2)
            tempIndex = tempIndex + 1
        End If
    Next L2
    
    ' Move pivot to the final position.
    ' Partition is now complete, all smaller values are to the left of pivot.
    fnSwapValues inputArr(tempIndex), inputArr(Hi)
    
    ' Return the final position index.
    fnPartition = tempIndex
End Function

Function QuickSelect(ByRef inputArr, Lo As Long, Hi As Long, k) As Double
' Receive an array, which may be unsorted.
' Return the k-th smallest value.
'
    If Lo = Hi Then ' <~~ Only 1 value
        QuickSelect = inputArr(Lo)
        Exit Function
    End If
    
    Dim pivotIndex As Long
    
    pivotIndex = (Lo + Hi) \ 2 ' <~~ 1st guess at midpoint
    pivotIndex = fnPartition(inputArr, Lo, Hi, pivotIndex)
    
    If k = pivotIndex Then
        QuickSelect = inputArr(k)
    ElseIf k < pivotIndex Then
        QuickSelect = QuickSelect(inputArr, Lo, pivotIndex - 1, k)
    ElseIf k > pivotIndex Then
        QuickSelect = QuickSelect(inputArr, pivotIndex + 1, Hi, k)
    End If
End Function

Function QuickSort(ByRef inputArr, Lo, Hi, numOfDims, Optional refCol = 0)
' Sort an array in ascending lexicographic order.
' numOfDims:
'     1 - Sort 1D array of values.
'     2 - Sort 1D array of 1D arrays (requires a reference column to sort by).
'
    Dim tmpLo As Long
    Dim tmpHi As Long
    Dim tempVal As Variant
    Dim pivotVal As Variant
    
    If numOfDims = 1 Then
        pivotVal = inputArr((Lo + Hi) \ 2) ' <~~ Pivot at the middle
    ElseIf numOfDims = 2 Then
        pivotVal = inputArr((Lo + Hi) \ 2)(refCol) ' <~~ Pivot at the middle
    End If
    
    tmpLo = Lo
    tmpHi = Hi
    Do While tmpLo <= tmpHi
        If numOfDims = 1 Then
            ' Scan for leftmost value larger than pivot.
            Do While inputArr(tmpLo) < pivotVal And tmpLo < Hi
              tmpLo = tmpLo + 1
            Loop
            
            ' Scan for rightmost value smaller than pivot.
            Do While inputArr(tmpHi) > pivotVal And tmpHi > Lo
              tmpHi = tmpHi - 1
            Loop
        ElseIf numOfDims = 2 Then
            ' Scan for leftmost array with reference value larger than pivot.
            Do While inputArr(tmpLo)(refCol) < pivotVal And tmpLo < Hi
              tmpLo = tmpLo + 1
            Loop
            
            ' Scan for rightmost array with reference value smaller than pivot.
            Do While inputArr(tmpHi)(refCol) > pivotVal And tmpHi > Lo
              tmpHi = tmpHi - 1
            Loop
        End If
        
        ' Swap values.
        If tmpLo <= tmpHi Then
            fnSwapValues inputArr(tmpLo), inputArr(tmpHi)
            tmpLo = tmpLo + 1
            tmpHi = tmpHi - 1
        End If
    Loop
    
    If Lo < tmpHi Then QuickSort inputArr, Lo, tmpHi, numOfDims, refCol
    If tmpLo < Hi Then QuickSort inputArr, tmpLo, Hi, numOfDims, refCol
End Function

Function fnRoundArrNums(ByVal inputArr, Optional decimalPlaces = 4) As Variant
' Round off the numbers in a 2D array to a number of decimal places.
'
    Dim L2 As Long
    Dim L3 As Long
    
    For L2 = LBound(inputArr) To UBound(inputArr)
        For L3 = LBound(inputArr, 2) To UBound(inputArr, 2)
            If LenB(inputArr(L2, L3)) <> 0 Then ' Ignore empty cells
                If IsNumeric(inputArr(L2, L3)) Then
                    inputArr(L2, L3) = Round(inputArr(L2, L3), decimalPlaces)
                End If
            End If
        Next L3
    Next L2
    
    fnRoundArrNums = inputArr
End Function

Function printArray( _
    ByRef inArr, _
    Optional hBordRefCols As String = "-1", _
    Optional newBookOrSheet As Long = 0, _
    Optional freezeRow As Long = 1, _
    Optional freezeCol As Long = 0, _
    Optional colAutoFit As String = "None", _
    Optional sheetName As String = "None")
' Output contents of a 2D array.
'
    ' Choose to output onto a new workbook or a new sheet.
    If newBookOrSheet = 0 Then
        Workbooks.Add
    ElseIf newBookOrSheet = 1 Then
        Sheets.Add , ActiveSheet
    End If
    
    ' Output array values.
    With Range(Cells(1, 1), Cells(1 + UBound(inArr, 1), 1 + UBound(inArr, 2)))
        .Value2 = fnRoundArrNums(inArr, 3)
    End With
    
    ' Dark blue background and white font for header row(s).
    With Rows("1:" & freezeRow)
        .Interior.Color = 6299648 ' Dark blue
        .Font.Color = 16777215 ' White
    End With
    
    ' Freeze specified rows and columns.
    With ActiveWindow
        .SplitRow = freezeRow
        .SplitColumn = freezeCol
        .FreezePanes = True
    End With
    
    ' Draw horizonal borders.
    horizontalBorders inArr, hBordRefCols
    
    ' Autofit columns.
    If colAutoFit <> "None" Then
        If colAutoFit = "All" Then
            Columns.AutoFit
        Else
            Columns(colAutoFit).AutoFit
        End If
    End If
    
    ' Rename sheet.
    If sheetName <> "None" Then ActiveSheet.Name = sheetName
End Function

Function horizontalBorders(ByRef inArr, colList, Optional deLim = ",")
' Draw horizontal borders for specified rows.
'
    If colList = "-1" Then Exit Function
    
    Dim L2 As Long
    Dim L3 As Long
    Dim currStr As String
    Dim refStr As String
    Dim colArr As Variant
    
    colArr = Split(colList, deLim)
    
    For L2 = LBound(inArr) + 1 To UBound(inArr)
        ' Concatenate reference values for current row.
        currStr = vbNullString
        For L3 = LBound(colArr) To UBound(colArr)
            currStr = currStr & inArr(L2, colArr(L3))
        Next L3
        
        If refStr <> currStr Then
            ' Current row is in different section from previous row.
            Rows(L2).Borders(xlEdgeBottom).Weight = xlThin
            refStr = currStr
        End If
    Next L2
    
    Rows(L2).Borders(xlEdgeBottom).Weight = xlThin
End Function

Function rightBorders(strOfNums, Optional deLim = ",")
' Draw right borders for specified columns.
' Indicate column numbers in a comma-delimited string,
'
    Dim L2 As Long
    Dim numArr As Variant
    
    numArr = Split(strOfNums, deLim)
    
    For L2 = LBound(numArr) To UBound(numArr)
        Columns(CLng(numArr(L2))).Borders(xlEdgeRight).Weight = xlThin
    Next L2
End Function

Function numToLetter(num As Long) As String
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

Function fnMultiTo1DArr(ByRef inputArrND) As Variant
' Reduce a multidimensional array to a 1D array.
'
    Dim arrLen As Long
    Dim arrVal As Variant
    
    ReDim tempArr(0) As Double
    For Each arrVal In inputArrND
        If LenB(arrVal) <> 0 And IsNumeric(arrVal) And arrVal > 0 Then
            arrLen = arrLen + 1
            ReDim Preserve tempArr(arrLen) As Double
            tempArr(arrLen) = arrVal
        End If
    Next arrVal
    
    fnMultiTo1DArr = tempArr
End Function

Sub Z01_deleteFiles()
' Let user select which folder to operate on.
' Delete all files except for StepID.lin and group.sta.
'
    Dim L2 As Long
    Dim n1 As Long
    Dim FSO As Object
    Dim fldPath As String
    
    ReDim delRep(0) As String
    
    delRep(0) = "File deleted" & vbTab & "File path"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    L2 = MsgBox( _
        "Please select folder to DELETE files from." & vbCr & vbCr & _
        "The following files will be deleted:" & vbCr & _
        "  1. group.lin" & vbCr & _
        "  2. group.lst" & vbCr & _
        "  3. group.rec" & vbCr & _
        "  4. StepID.txt" & vbCr & _
        "  5. StepID.vel" & vbCr & vbCr & _
        "Are you sure you want to proceed?" & vbCr, vbOKCancel, "DELETE FILES")
    If L2 = vbCancel Then
        MsgBox "No changes made. Exiting script."
        Exit Sub
    End If
    
    ' Obtain the address of the desired folder.
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then
            MsgBox "No folder selected. Exiting script."
            Exit Sub
        End If
        fldPath = .SelectedItems(1)
    End With
    
    ' This will be the user's last chance to abort the procedure.
    L2 = MsgBox( _
        "The folder with the address: " & vbCr & vbCr & _
        "'" & fldPath & "'" & vbCr & vbCr & _
        "has been selected." & vbCr & vbCr & _
        "THIS WILL BE THE LAST CHANCE TO ABORT! Proceed?", vbOKCancel, "Confirm folder")
    If L2 = vbOK Then
        deleteFolderScan FSO.getFolder(fldPath), delRep, n1
        printArray linesToTable(delRep), , 0, 1, 0, "All", "None"
    Else
        MsgBox "No changes made. Exiting script."
    End If
End Sub

Function deleteFolderScan(fsoFld, delRep, n1)
' Iteratively search for subfolders within the selected folder.
' Once the terminal folder is reached, scan files within that folder.
'
    Dim FSO As Object
    Dim iObj As Object
    Dim passStepID As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With fsoFld
        ' Access subdirectories.
        For Each iObj In .SubFolders
            deleteFolderScan FSO.getFolder(iObj), delRep, n1
        Next iObj
        
        ' Get StepID from .lin files (with same file name as .sol files) during first-pass.
        For Each iObj In .Files
            With iObj
                If Right$(.Name, 4) = ".lin" Then
                    If Left$(.Name, 5) <> "group" Then ' Disregard group.lin files
                        passStepID = Left$(.Name, Len(.Name) - 4)
                        Exit For
                    End If
                End If
            End With
        Next iObj
        
        ' Delete files.
        For Each iObj In .Files
            With iObj
                If _
                    .Name = "group.lin" Or _
                    .Name = "group.lst" Or _
                    .Name = "group.rec" Or _
                    .Name = passStepID & ".txt" Or _
                    .Name = passStepID & ".vel" Then
                    ' Delete only these files.
                    ' Add to the list if you want to delete others.
                    ' The "delete everthing else" system is just a disaster waiting to happen.
                    
                    ' Update report.
                    n1 = n1 + 1
                    ReDim Preserve delRep(0 To n1) As String
                    delRep(n1) = n1 & vbTab & .Path
                    
                    ' Delete file.
                    FSO.deleteFile .Path
                End If
            End With
        Next iObj
    End With
End Function

' ================================================================================================
' Math Functions
'     Functions available in Excel
'     The objective here is future-proofing
' ================================================================================================

Function fnArrNumLen(ByRef inputArr) As Long
' Count the number of numerical values in an array (can be multidimensional).
'
    Dim arrVal As Variant
    
    For Each arrVal In inputArr
        If IsNumeric(arrVal) Then
            fnArrNumLen = fnArrNumLen + 1
        End If
    Next arrVal
End Function

Function fnArrRawSum(ByRef inputArr, Optional mOrdinal = 1) As Double
' Calculate the sum of the n-th (default of 1) power of numbers in an array.
'
    Dim arrVal As Variant
    
    For Each arrVal In inputArr
        If IsNumeric(arrVal) Then
            fnArrRawSum = fnArrRawSum + (arrVal ^ mOrdinal)
        End If
    Next arrVal
End Function

Function fnArrMean(ByRef inputArr) As Variant
' Calculate the average value of an array of numbers.
'
    Dim arrLen As Long
    
    arrLen = fnArrNumLen(inputArr) ' Number of numerical values
    
    If arrLen <> 0 Then
        fnArrMean = fnArrRawSum(inputArr, 1) / arrLen
    Else
        fnArrMean = errDiv0
    End If
End Function

Function fnArrCenSum(ByRef inputArr, Optional mOrdinal = 2) As Double
' Calculate the n-th (default of 2) order centralized sum of an array of numbers.
'
    Dim arrMean As Double
    Dim arrVal As Variant
    
    arrMean = fnArrMean(inputArr)
    
    If IsNumeric(arrMean) Then
        For Each arrVal In inputArr
            If IsNumeric(arrVal) Then
                fnArrCenSum = fnArrCenSum + ((arrVal - arrMean) ^ mOrdinal)
            End If
        Next arrVal
    End If
End Function

Function fnArrSD(ByRef inputArr) As Variant
' Calculate the variance of an array of numbers.
'
    Dim arrLen As Long
    Dim arrVar As Double
    
    arrLen = fnArrNumLen(inputArr)
    
    If arrLen > 1 Then
        arrVar = fnArrCenSum(inputArr, 2) / (arrLen - 1) ' Bessel's correction
        fnArrSD = Sqr(arrVar)
    Else
        fnArrSD = errDiv0
    End If
End Function

Function fnArrNormCenSum(ByRef inputArr, Optional mOrdinal = 3) As Variant
' Calculate the n-th (default of 3) order standardized sum of an array of numbers.
'
    Dim arrSD As Variant
    
    arrSD = fnArrSD(inputArr)
    
    If IsNumeric(arrSD) Then
        If arrSD <> 0 Then
            fnArrNormCenSum = fnArrCenSum(inputArr, mOrdinal) / (arrSD ^ mOrdinal)
        Else
            fnArrNormCenSum = errDiv0
        End If
    Else
        fnArrNormCenSum = "NA"
    End If
End Function

Function fnArrStdSkew(ByRef inputArr) As Variant
' Calculate the standardized skewness of an array of numbers.
'
    Dim arrLen As Long
    Dim normCenSum3 As Variant
    Dim skewConst As Double
    Dim arrSkew As Double
    Dim skewVariance As Double
    
    arrLen = fnArrNumLen(inputArr)
    
    If arrLen > 2 Then
        normCenSum3 = fnArrNormCenSum(inputArr, 3)
        
        If IsNumeric(normCenSum3) Then
            skewConst = arrLen / ((arrLen - 3) * arrLen + 2)
            arrSkew = normCenSum3 * skewConst ' <~~ Raw skewness
            
            skewVariance = 6 * (arrLen - 1) * arrLen / (((arrLen + 2) * arrLen - 5) * arrLen - 6)
            fnArrStdSkew = arrSkew / Sqr(skewVariance)
        Else
            fnArrStdSkew = normCenSum3
        End If
    Else
        fnArrStdSkew = errDiv0
    End If
End Function

Function fnArrStdKurt(ByRef inputArr) As Variant
' Calculate the standardized kurtosis of an array of numbers.
'
    Dim arrLen As Long
    Dim normCenSum4 As Variant
    Dim kurtConst1 As Double
    Dim kurtConst2 As Double
    Dim arrKurt As Double
    Dim skewVariance As Double
    Dim kurtVariance As Double
    
    arrLen = fnArrNumLen(inputArr)
    
    If arrLen > 3 Then
        normCenSum4 = fnArrNormCenSum(inputArr, 4)
        
        If IsNumeric(normCenSum4) Then
            kurtConst1 = (arrLen + 1) * arrLen / (((arrLen - 6) * arrLen + 11) * arrLen - 6)
            kurtConst2 = ((3 * arrLen - 6) * arrLen + 3) / ((arrLen - 5) * arrLen + 6)
            arrKurt = normCenSum4 * kurtConst1 - kurtConst2 ' <~~ Raw kurtosis
            
            skewVariance = 6 * (arrLen - 1) * arrLen / (((arrLen + 2) * arrLen - 5) * arrLen - 6)
            kurtVariance = 4 * skewVariance * (arrLen * arrLen - 1) / ((arrLen + 2) * arrLen - 15)
            fnArrStdKurt = arrKurt / Sqr(kurtVariance)
        Else
            fnArrStdKurt = normCenSum4
        End If
    Else
        fnArrStdKurt = errDiv0
    End If
End Function

Function fnSumXY(inArr1, inArr2) As Double
' Calculate the sum of the products of the deviations from two arrays.
' Arrays should have been checked for equal lengths beforehand.
'
    Dim L2 As Long
    Dim arrMean1 As Double
    Dim arrMean2 As Double
    
    arrMean1 = fnArrMean(inArr1)
    arrMean2 = fnArrMean(inArr2)
    
    If IsNumeric(arrMean1) And IsNumeric(arrMean2) Then
        For L2 = LBound(inArr1) To UBound(inArr1)
            fnSumXY = fnSumXY + ((inArr1(L2) - arrMean1) * (inArr2(L2) - arrMean2))
        Next L2
    End If
End Function

Function fnCovar(inArr1, inArr2) As Variant
' Calculate the covariance of two arrays.
'
    Dim arrLen1 As Long
    Dim arrLen2 As Long
    
    arrLen1 = fnArrNumLen(inArr1)
    arrLen2 = fnArrNumLen(inArr2)
    
    If arrLen1 = arrLen2 And arrLen1 > 1 And arrLen2 > 1 Then
        fnCovar = fnSumXY(inArr1, inArr2) / arrLen1
    Else
        fnCovar = "NA"
    End If
End Function

Function fnCorrel(inArr1, inArr2) As Variant
' Calculate the correlation of two arrays.
'
    Dim arrLen1 As Long
    Dim arrLen2 As Long
    Dim sumXX As Double
    Dim sumYY As Double
    
    arrLen1 = fnArrNumLen(inArr1)
    arrLen2 = fnArrNumLen(inArr2)
    
    If arrLen1 = arrLen2 And arrLen1 > 1 And arrLen2 > 1 Then
        sumXX = fnArrCenSum(inArr1, 2)
        sumYY = fnArrCenSum(inArr2, 2)
        
        If sumXX <> 0 And sumYY <> 0 Then
            fnCorrel = fnSumXY(inArr1, inArr2) / Sqr(sumXX) / Sqr(sumYY)
        Else
            fnCorrel = "NA"
        End If
    Else
        fnCorrel = "NA"
    End If
End Function

Function fnArrQuartile(ByRef inputArr, quartileNum As Long) As Variant
' Calculate the 0th, 1st, 2nd, 3rd, or 4th quartile of an array of numbers.
'
    If quartileNum < 0 Or quartileNum > 4 Then
        Debug.Print "Error: Invalid quartile entry (must be 0, 1, 2, 3, or 4)."
        fnArrQuartile = "Err"
        Exit Function
    End If
    
    Dim arrLen As Long
    Dim quartInt As Long
    Dim quartDbl As Double
    Dim quartLBound As Double
    Dim quartUBound As Double
    
    arrLen = UBound(inputArr) - LBound(inputArr) + 1
    quartDbl = (arrLen - LBound(inputArr)) * quartileNum / 4 + LBound(inputArr)
    quartInt = Int(quartDbl)
    quartLBound = QuickSelect(inputArr, LBound(inputArr), UBound(inputArr), quartInt)
    quartUBound = QuickSelect(inputArr, LBound(inputArr), UBound(inputArr), quartInt + 1)
    
    fnArrQuartile = quartLBound + (quartDbl - quartInt) * (quartUBound - quartLBound)
End Function

Function fnTDist95CI(ByRef inputArr, Optional alpha = typeOneErr) As Variant
' Calculate the (100 - alpha)% confidence interval of an array of numbers.
'
    Dim arrLen As Long
    Dim arrMean As Double
    Dim arrSE As Double
    Dim ci95span As Double
    Dim outputArr(1 To 2) As Variant
    
    arrLen = fnArrNumLen(inputArr)
    
    If arrLen > 0 Then
        arrMean = fnArrMean(inputArr)
        arrSE = fnArrSD(inputArr) / Sqr(arrLen)
        ci95span = Application.TInv(alpha, arrLen - 1) * arrSE
        
        outputArr(1) = arrMean - ci95span ' 95CI lower bound
        outputArr(2) = arrMean + ci95span ' 95CI upper bound
    Else
        outputArr(1) = "NA"
        outputArr(2) = "NA"
    End If
    
    fnTDist95CI = outputArr
End Function

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

' ================================================================================================
' Other unused functions
' ================================================================================================

Function determineOS()
' Mac OS not supported at the moment.
'
' Possible returns:
'     "Windows (32-bit) NT 10.00" <~~ Roscoe's PC
'     "Windows (32-bit) NT :.00" <~~ KG R block level 6 PC
'     "Macintosh ..."
'
' Should be sufficient to distinguish OS.
'
    Debug.Print Application.OperatingSystem
End Function

Function nChooseK(ByVal N, ByVal k) As Long
' Return number of ways to choose k items fron n items.
' n choose k = n! / ((n-k)! * k!).
'
' Up to n = 29 with no issues.
' Overflow error when:
'     30 choose 15,
'     46342 choose 2.
'
    If k < 0 Or k > N Then
        Debug.Print "Invalid input for nChooseK (k cannot be negative or larger than n)."
        Exit Function
    End If
    
    If k > N - k Then k = N - k ' <~~ By symmetry, n choose k = n choose (n-k)
    
    If k = 0 Then
        nChooseK = 1
        Exit Function
    End If
    
    nChooseK = nChooseK(N - 1, k - 1) * N \ k ' <~~ Integer division must be done last
End Function

Function zFigureFactory()
' This will be the next level.
' Eliminates the need for the artist.
'
    Dim filePath As String
    
    MsgBox "Please select the desired footprint image file."
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show <> -1 Then MsgBox "No file selected. Exiting script.": Exit Function
        filePath = .SelectedItems(1)
    End With
    
    Sheets.Add , Sheets(Sheets.Count)
    
    ActiveSheet.Pictures.Insert(filePath).Select
    With Selection.ShapeRange.ThreeD
        .SetPresetCamera (msoCameraIsometricTopUp)
        .RotationX = 45
        .RotationY = -30
        .RotationZ = -60
    End With
    Selection.ShapeRange.IncrementLeft 400
    
    ActiveWindow.Zoom = 50
End Function

' ================================================================================================
' Log:
'
' 2018/06/30.
' Changed method for generating rawArr.
' Different numbers of steps now accepted for stats analysis.
' Watermelon table now accepts non-uniform numbers of steps.
' Force-time graph function now accepts non-uniform numbers of steps.
' ICC function now receives input truncated to lowest common number of steps across subjects.
' Non-uniform numbers of shoes still not accepted.
'
' 2018/06/27.
' Changed the way StepID is obtained.
'     The old method was to read the name of the terminal folder.
'     Terminal folders have to be named manually, and is thus very prone to human error.
'     The new way is to find the .lin file with the same name as the .sol file used to create it.
'     The condition is that the EMASCII files generated must not be renamed.
'     Still vulnerable to human error, but less so.
'     The risk of group.sta files being overwritten still persists.
'     Thus, proper file organization into folders will still be crucial.
'
' 2018/06/10.
' Test and debug program on 3 shoe models and 16 subjects.
' The next challenge will be implementing the Shapiro-Wilk test for n != 16.
' Shapiro-Wilk test now available for sample sizes up to n = 20.
' Test and debug program on 3 shoe models and 15 subjects.
' Test program on 2 shoe models and 16 subjects.
' Test program on 2 shoe models and 15 subjects.
' Test and debug program on 2 shoe models and 5 subjects (some strange results).
' At least 8 subjects recommended.
' Test program on 5 shoe models and 14 subjects.
'
' 2018/06/07.
' Hard-coded variables removed (FINALLY!).
' Algorithms used to generate the global lists have not been thoroughly verified, yet.
'
' 2018/06/04.
' Fix Wilcoxon signed rank test.
'     Some variables mis-defined as Long instead of Double type, leading to rounding errors.
'     Results now more consistent with SPSS output in general.
'     The remaining inconsistencies involve questionable SPSS output.
'     e.g. significant p values where the estimated difference is 0, or 95% CIs containing 0.
'
' 2018/06/03.
' Recursive algorithm for generating non-repeated n choose k lists.
'
' 2018/05/19.
' Did not greet mum for birthday. Depressed.
'
' 2018/05/15.
' Tidying up.
' "Do Everything" button.
'
' 2018/05/14.
' Insert references for procedures.
'
' 2018/05/13.
' "Do Everything" function.
' 18.3 seconds from separate data files to final results table.
' Updated bar charts and error bars (medians and 95% CIs).
' Did not greet mum for Mother's Day. Sad.
'
' 2018/05/12.
' Housekeeping.
'
' 2018/05/11.
' Optimize ICC calculation speed (remove one inefficiency).
'     Old time: ~ 3 minutes.
'     New time: ~ 15 seconds.
'
' 2018/05/10.
' Use lower bound estimate of sphericity for interaction term.
' This will be a limitation until the algorithm is figured out.
'
' 2018/05/09.
' Two-way repeated measures ANOVA.
' All values available except sphericity estimates for interaction term.
'
' 2018/05/08.
' One-way repeated measures ANOVA.
' Greenhouse-Geisser and Huynh-Feldt epsilon estimates.
' Mauchly's test not implemented, corrections applied according to average of estimates.
' Biharmonic interpolation of adjusted degree of freedoms for calculating p values.
' Require algorithm for finding eigenvalues of a matrix.
'
' 2018/05/07.
' Shapiro-Wilk test based on Shapiro and Wilk (1946).
' Linear interpolation of transformed test statistic (raised to the power of 20).
' No significant disagreements with SPSS output.
'
' 2018/05/04.
' ICC calculations much quicker than expected (actual 2.5 min vs. expected 30 min).
'
' 2018/05/01.
' Remove Walsh average calculation sheet.
' Create hodgeLehmann (Hodges-Lehmann estimator) function.
'
' 2018/04/30.
' Inconsistencies with Wilcoxon signed-rank test values.
' SPSS reports significant differences between conditions, with median difference of 0.
'
' 2018/04/27.
' Wilcoxon signed-rank test.
'
' 2018/04/26.
' Calculate ICC(3,k) absolute agreement for single array.
' Multiple column combinations.
'
' 2018/04/25.
' Remove excelSort completely (archived).
'
' 2018/04/24.
' Tidy tTestArray.
' Allow multiple input files.
' ICC optional.
'
' 2018/04/23.
' New color scheme for "watermelon" sheet.
' Fix empty column in "watermelon" sheet.
'
' 2018/04/22.
' Remove excelSort (mostly).
' Fix damages.
'
' 2018/04/21.
' Auto border lines for Omnibus.
' Attempt to remove excelSort.
'
' 2018/04/20.
' Reduce hard-coding for BarCharts sub.
' Cohen's d function.
' Standardize highlight subs.
'
' 2018/04/19.
' Added ICC data.
' Rearranged Groups columns.
' Color for ICCs and for Walsh averages in Hodges-Lehmann algorithm.
'
' 2018/04/18.
' Fixed error bars.
'
' 2018/04/17.
' Friedman's test successful! One less thing for SPSS to do!
'
' 2018/04/16.
' SPSS output process overhaul.
'
' 2018/04/15.
' Break masterTable sub into separate pieces.
'
' 2018/04/14.
' Got rid of workBook_3, the last public/global variable. Finally!
' Data extraction procedure is now fully based on arrays.
' Restructure workflow, process .lin data first.
' 11.9 s to get from raw data to secondMasterTable.
' 15.3 s to get from raw data to pre-SPSS (+ file report).
'
' 2018/04/13.
' Removed hard-coding for column letters for force-time graphs.
'
' 2018/04/12.
' Removed cLists.
' Fixed error bars for most bar charts.
' Faster data processing time:
'     Old: 21.0 sec
'     New: 9.8 sec for firstMasterTable (likely due to ByRef)
'
' 2018/04/11.
' Overhaul force-time subroutines.
' Most subs running on .txt
' Bar charts (without error bars).
'
' 2018/04/10.
' Recoded SPSS table generator.
'
' 2018/04/09.
' Discarded old post-SPSS subroutines.
' Complete master/results table without bar charts.
' Read tab delimited text files(!!!) OMG! This is going to solve problems!
' Significantly faster data gathering time:
'     Old: 8 min 25 sec
'     New: 0 min 21 sec (24 times faster!)
'
' 2018/04/08.
' Read from text file procedure.
' Back to SPSS scripting.
'
' 2018/04/07.
' Declare public constants.
'
' 2018/04/02.
' No major changes, working on SPSS script.
' Try to stay within 100 columns.
'
' 2018/04/01.
' Major reorganization of code.
'
' 2018/03/31.
' Group stats done.
' Chart data done.
'
' 2018/03/30.
' Continue work on Hodges-Lehmann algorithm.
' Stats output done in single sheet.
'
' 2018/03/29.
' Work on Hodges-Lehmann algorithm.
' Issue with sorting.
' Implemented Quicksort. It works spectacularly.
'
' 2018/03/28.
' Purchased new keyboard.
' Remove redundant codes.
' Remove more global variables.
' Tidy up new master table macro.
'
' 2018/03/27.
' Keyboard issues causing Excel to crash.
' Rest.
'
' 2018/03/26.
' t-tests done.
'
' 2018/03/25.
' Correlation is proving to be challenging.
' Rest.
'
' 2018/03/24.
' Working up to t-tests.
' Unable to calculate correlation, yet.
'
' 2018/03/23.
' Pairwise differences done.
'
' 2018/03/22.
' Moving on to processing master table.
' Individual means done.
'
' 2018/03/21.
' 2 error in calculations detected.
' Errors in calculations corrected.
'
' 2018/03/20.
' New method for finding local extrema.
' All calculations replicated.
'
' 2018/03/19.
' Timing: 8 min 25 sec.
' Moving on to force-time processing.
' Scanning a window of 15 cells does not work for arrays.
'
' 2018/03/18.
' Removed some one-line functions.
' Slightly improved timing: 8 min 22 sec.
' But file size is reduced to about 1/3.
'
' 2018/03/17.
' Trying out reading without opening workbook.
' Seems to work ok for .lin data.
' Not so applicable for .sta data.
' Removed a GoTo.
' Removed several global/public variables.
' Core functionality is still operational.
'
' 2018/03/16.
' Slower timing: 9 min 51 sec.
'
' 2018/03/15.
' On second thought, I don't know what to do with time process data for individual masks.
' Work on getting the other .lin data for the whole foot.
'
' 2018/03/14.
' Faster timing:
'     Old: ~30 min
'     New: 5 min 46 sec (~5x to 6x faster)
' Data extraction series successfully revamped.
' Requires tidying up.
'
' 2018/03/13.
' The new approach works!
' Less workbooks required.
' Resolved the "semi-final master table" problem.
' Both .sta and .lin data extracted simultaneously, not separately.
' Timing is about the same, but more data extracted.
'
' 2018/03/12.
' The work continues after a long hiatus.
' Major renovations planned for the module.
' Discard as much as possible (minimalism principle).
' Workflow is mostly the same:
'     1. Extract and collate data.
'     2. Process data.
'     3. Statistics.
'     4. Presentation.
' ================================================================================================
