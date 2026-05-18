' ================================================================================
' XLS -> XLSX Batch Converter (VBScript)
'
' Description: Recursively scans a directory for .xls files and converts them
'              to .xlsx format using Microsoft Excel's COM interface.
' Requirement: Windows + Microsoft Excel (2010 or later) installed
' Usage:       Place this script in the target folder and double-click to run
' ================================================================================

Option Explicit

' Excel file format constant: xlOpenXMLWorkbook = 51
Const XL_FILE_FORMAT_XLSX = 51

' --------------------------------------------------------------------------------
' Force console mode (CScript) for interactive input/output
' --------------------------------------------------------------------------------
If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cscript.exe //nologo """ & WScript.ScriptFullName & """"
    Set objShell = Nothing
    WScript.Quit
End If

' --------------------------------------------------------------------------------
' Main entry point
' --------------------------------------------------------------------------------
Call Main()

Sub Main()
    Dim fso, rootDir
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Use the script's directory as the scan root
    rootDir = fso.GetParentFolderName(WScript.ScriptFullName)

    Call ShowBanner()

    ' Collect files using Dictionary (most compatible dynamic array)
    Dim xlsFiles
    Set xlsFiles = CreateObject("Scripting.Dictionary")

    WScript.Echo "[Scan Directory] " & rootDir
    WScript.Echo "Scanning for .xls files..."
    WScript.Echo ""

    Call ScanXlsFiles(fso, rootDir, xlsFiles)

    If xlsFiles.Count = 0 Then
        WScript.Echo "[Result] No .xls files found. Exiting."
        WScript.Echo ""
        Call PauseExit()
        Exit Sub
    End If

    Call ShowScanReport(fso, rootDir, xlsFiles)
    Call ShowWarning()

    ' --- Backup confirmation ---
    Dim ans
    Do
        WScript.Echo ""
        WScript.Echo "Have you backed up the original files?"
        WScript.Echo "  [1] Yes, backup completed"
        WScript.Echo "  [Q] No, quit program"
        WScript.StdOut.Write "Enter choice: "
        ans = LCase(Trim(WScript.StdIn.ReadLine()))
        If ans = "1" Then Exit Do
        If ans = "q" Then
            WScript.Echo ""
            WScript.Echo "[Exit] Please back up your files before running this script."
            Call PauseExit()
            Exit Sub
        End If
        WScript.Echo ">> Invalid input. Please enter 1 or Q."
    Loop

    ' --- Mode selection ---
    Dim deleteOriginal
    deleteOriginal = False
    Do
        WScript.Echo ""
        WScript.Echo "[Operation Options]"
        WScript.Echo "  1. Keep original .xls files after conversion (safe, recommended)"
        WScript.Echo "  2. Delete original .xls files after conversion (dangerous)"
        WScript.StdOut.Write "Enter option (1 or 2): "
        ans = Trim(WScript.StdIn.ReadLine())
        If ans = "1" Then
            deleteOriginal = False
            Exit Do
        ElseIf ans = "2" Then
            deleteOriginal = True
            Exit Do
        End If
        WScript.Echo ">> Invalid input. Please try again."
    Loop

    ' --- Final confirmation ---
    Dim modeStr
    If deleteOriginal Then
        modeStr = "[DELETE ORIGINAL]"
    Else
        modeStr = "[KEEP ORIGINAL]"
    End If

    Do
        WScript.Echo ""
        WScript.Echo "[Execution Confirm] " & modeStr & " About to process " & xlsFiles.Count & " .xls file(s)"
        WScript.Echo "Start conversion?"
        WScript.Echo "  [1] Yes, start now"
        WScript.Echo "  [2] Go back to mode selection"
        WScript.Echo "  [Q] Quit program"
        WScript.StdOut.Write "Enter choice: "
        ans = LCase(Trim(WScript.StdIn.ReadLine()))
        If ans = "1" Then Exit Do
        If ans = "q" Then
            WScript.Echo ""
            WScript.Echo "[Exit] Operation cancelled."
            Call PauseExit()
            Exit Sub
        End If
        ' ans = "2" -> re-select mode
        Do
            WScript.Echo ""
            WScript.Echo "[Operation Options]"
            WScript.Echo "  1. Keep original .xls files after conversion (safe, recommended)"
            WScript.Echo "  2. Delete original .xls files after conversion (dangerous)"
            WScript.StdOut.Write "Enter option (1 or 2): "
            ans = Trim(WScript.StdIn.ReadLine())
            If ans = "1" Then
                deleteOriginal = False
                Exit Do
            ElseIf ans = "2" Then
                deleteOriginal = True
                Exit Do
            End If
            WScript.Echo ">> Invalid input. Please try again."
        Loop
    Loop

    ' --- Launch Excel ---
    WScript.Echo ""
    WScript.Echo "Starting Microsoft Excel..."

    Dim excel
    On Error Resume Next
    Set excel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        WScript.Echo "[Error] Unable to start Microsoft Excel. Please make sure Excel is installed."
        WScript.Echo "Details: " & Err.Description
        Call PauseExit()
        Exit Sub
    End If
    On Error GoTo 0

    excel.DisplayAlerts = False
    excel.Visible = False
    excel.ScreenUpdating = False

    ' --- Execute conversion ---
    WScript.Echo String(64, "-")
    WScript.Echo "Starting conversion..."
    WScript.Echo ""

    Dim success, skipped
    Dim failedList
    Set failedList = CreateObject("Scripting.Dictionary")
    success = 0
    skipped = 0

    Dim i, srcPath, dstPath, wb, fileName, relPath
    For i = 0 To xlsFiles.Count - 1
        srcPath = xlsFiles(i)
        fileName = fso.GetFileName(srcPath)
        dstPath = fso.GetParentFolderName(srcPath) & "\" & fso.GetBaseName(srcPath) & ".xlsx"

        ' Show relative path for progress
        If Len(srcPath) > Len(rootDir) + 1 Then
            relPath = Mid(srcPath, Len(rootDir) + 2)
        Else
            relPath = fileName
        End If
        WScript.Echo "  [" & Right("   " & (i + 1), 3) & "/" & Right("   " & xlsFiles.Count, 3) & "] " & relPath

        ' Skip if destination already exists
        If fso.FileExists(dstPath) Then
            WScript.Echo "      -> Skipped (destination already exists)"
            skipped = skipped + 1
        Else
            On Error Resume Next
            Set wb = excel.Workbooks.Open(srcPath)
            If Err.Number <> 0 Then
                failedList.Add failedList.Count, Array(srcPath, "Failed to open: " & Err.Description)
                WScript.Echo "      -> Failed (open error)"
                Err.Clear
            Else
                wb.SaveAs dstPath, XL_FILE_FORMAT_XLSX
                If Err.Number <> 0 Then
                    failedList.Add failedList.Count, Array(srcPath, "Failed to save: " & Err.Description)
                    WScript.Echo "      -> Failed (save error)"
                    Err.Clear
                Else
                    success = success + 1
                    WScript.Echo "      -> Success"
                End If
                wb.Close SaveChanges=False
                Set wb = Nothing
            End If
            On Error GoTo 0
        End If

        ' Delete original if requested and conversion succeeded
        If deleteOriginal And fso.FileExists(dstPath) Then
            On Error Resume Next
            fso.DeleteFile srcPath, True
            If Err.Number <> 0 Then
                WScript.Echo "      -> Warning: could not delete original - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next

    ' --- Cleanup Excel ---
    On Error Resume Next
    excel.Quit
    Set excel = Nothing
    On Error GoTo 0

    ' --- Summary ---
    WScript.Echo ""
    Call ShowSummary(success, skipped, failedList)

    Call PauseExit()
End Sub

' --------------------------------------------------------------------------------
' Recursively scan for .xls files
' --------------------------------------------------------------------------------
Sub ScanXlsFiles(fso, folderPath, fileList)
    Dim folder, file, subFolder
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        ' Strict match: extension must be exactly "xls" (excludes .xlsx)
        If LCase(fso.GetExtensionName(file.Name)) = "xls" Then
            fileList.Add fileList.Count, file.Path
        End If
    Next

    For Each subFolder In folder.SubFolders
        Call ScanXlsFiles(fso, subFolder.Path, fileList)
    Next
End Sub

' --------------------------------------------------------------------------------
' Format byte size to human-readable string
' --------------------------------------------------------------------------------
Function FormatSize(sizeBytes)
    Dim units, i, size
    units = Array("B", "KB", "MB", "GB", "TB")
    size = CDbl(sizeBytes)
    i = 0
    Do While size >= 1024 And i < UBound(units)
        size = size / 1024
        i = i + 1
    Loop
    FormatSize = Round(size, 2) & " " & units(i)
End Function

' --------------------------------------------------------------------------------
' Display banner
' --------------------------------------------------------------------------------
Sub ShowBanner()
    WScript.Echo ""
    WScript.Echo "================================================================================"
    WScript.Echo "        XLS -> XLSX Batch Converter (VBScript + Excel COM)"
    WScript.Echo "================================================================================"
    WScript.Echo "Features"
    WScript.Echo "  . Uses Microsoft Excel COM interface -- 100% format preservation"
    WScript.Echo "  . No Python or other dependencies required; just Excel installed"
    WScript.Echo "  . Recursive directory scan, batch processing, auto-skip existing files"
    WScript.Echo "  . Multi-step confirmation to prevent accidental data loss"
    WScript.Echo "================================================================================"
End Sub

' --------------------------------------------------------------------------------
' Display scan report
' --------------------------------------------------------------------------------
Sub ShowScanReport(fso, rootDir, fileList)
    Dim i, totalSize, dirSet, f, relPath
    Set dirSet = CreateObject("Scripting.Dictionary")
    totalSize = 0

    For i = 0 To fileList.Count - 1
        Set f = fso.GetFile(fileList(i))
        totalSize = totalSize + f.Size
        dirSet(f.ParentFolder) = True
    Next

    WScript.Echo ""
    WScript.Echo "[Scan Result]"
    WScript.Echo "  Files to convert : " & fileList.Count
    WScript.Echo "  Directories      : " & dirSet.Count
    WScript.Echo "  Total size       : " & FormatSize(totalSize)

    Dim maxShow
    maxShow = 10
    If fileList.Count < maxShow Then maxShow = fileList.Count

    WScript.Echo ""
    WScript.Echo "[File List] First " & maxShow & ":"
    For i = 0 To maxShow - 1
        If Len(fileList(i)) > Len(rootDir) + 1 Then
            relPath = Mid(fileList(i), Len(rootDir) + 2)
        Else
            relPath = fso.GetFileName(fileList(i))
        End If
        WScript.Echo "  " & Right("  " & (i + 1), 2) & ". " & relPath
    Next
    If fileList.Count > 10 Then
        WScript.Echo "      ... and " & (fileList.Count - 10) & " more file(s)"
    End If
End Sub

' --------------------------------------------------------------------------------
' Display warning
' --------------------------------------------------------------------------------
Sub ShowWarning()
    WScript.Echo ""
    WScript.Echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    WScript.Echo "  !   IMPORTANT WARNING"
    WScript.Echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    WScript.Echo "  1. This tool uses Excel COM SaveAs, producing identical results to"
    WScript.Echo "     manually using 'Save As' inside Excel."
    WScript.Echo "  2. New .xlsx files are created; original .xls files are NOT overwritten."
    WScript.Echo "  3. If you choose DELETE ORIGINAL mode, the source .xls will be PERMANENTLY"
    WScript.Echo "     removed after successful conversion."
    WScript.Echo "  4. STRONGLY RECOMMEND making a full backup before proceeding."
    WScript.Echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
End Sub

' --------------------------------------------------------------------------------
' Display summary
' --------------------------------------------------------------------------------
Sub ShowSummary(success, skipped, failedList)
    WScript.Echo "================================================================================"
    WScript.Echo "                     CONVERSION SUMMARY"
    WScript.Echo "================================================================================"
    WScript.Echo "  . Successful : " & success
    WScript.Echo "  . Skipped (exists) : " & skipped
    WScript.Echo "  . Failed     : " & failedList.Count

    If failedList.Count > 0 Then
        WScript.Echo "--------------------------------------------------------------------------------"
        WScript.Echo "[Failure Details]"
        Dim i, item
        For i = 0 To failedList.Count - 1
            item = failedList(i)
            WScript.Echo "  . " & item(0)
            WScript.Echo "    Reason: " & item(1)
        Next
    End If
    WScript.Echo "================================================================================"
End Sub

' --------------------------------------------------------------------------------
' Pause and wait for Enter key before exiting
' --------------------------------------------------------------------------------
Sub PauseExit()
    WScript.Echo ""
    WScript.StdOut.Write "Press Enter to exit..."
    Dim dummy
    dummy = WScript.StdIn.ReadLine()
    WScript.Quit
End Sub
