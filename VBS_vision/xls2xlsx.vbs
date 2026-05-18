' ================================================================================
' XLS -> XLSX 批量转换工具 (VBScript)
' 
' 功能：递归扫描目录中的所有 .xls 文件，调用 Microsoft Excel 另存为 .xlsx
' 要求：Windows + 已安装 Microsoft Excel (2010+)
' 用法：将本脚本放到目标文件夹，双击运行
' ================================================================================

Option Explicit

' Excel 文件格式常量: xlOpenXMLWorkbook = 51
Const XL_FILE_FORMAT_XLSX = 51

' --------------------------------------------------------------------------------
' 强制使用 CScript 控制台模式运行（确保有命令行窗口）
' --------------------------------------------------------------------------------
If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cscript.exe //nologo """ & WScript.ScriptFullName & """"
    Set objShell = Nothing
    WScript.Quit
End If

' --------------------------------------------------------------------------------
' 主入口
' --------------------------------------------------------------------------------
Call Main()

Sub Main()
    Dim fso, rootDir
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 脚本所在目录作为扫描根目录
    rootDir = fso.GetParentFolderName(WScript.ScriptFullName)
    
    ' 显示标题
    Call ShowBanner()
    
    ' 扫描文件 - 使用 Dictionary 模拟动态数组（兼容性最好）
    Dim xlsFiles
    Set xlsFiles = CreateObject("Scripting.Dictionary")
    
    WScript.Echo "[扫描目录] " & rootDir
    WScript.Echo "正在扫描 .xls 文件..."
    WScript.Echo ""
    
    Call ScanXlsFiles(fso, rootDir, xlsFiles)
    
    If xlsFiles.Count = 0 Then
        WScript.Echo "[结果] 未发现任何 .xls 文件，程序退出。"
        WScript.Echo ""
        Call PauseExit()
        Exit Sub
    End If
    
    ' 显示扫描报告
    Call ShowScanReport(fso, rootDir, xlsFiles)
    
    ' 安全确认流程
    Call ShowWarning()
    
    ' 备份确认
    Dim ans
    Do
        WScript.Echo ""
        WScript.Echo "是否已完成原始文件备份？"
        WScript.Echo "  [1] 已备份"
        WScript.Echo "  [Q] 未备份，退出程序"
        WScript.StdOut.Write "请输入: "
        ans = LCase(Trim(WScript.StdIn.ReadLine()))
        If ans = "1" Then Exit Do
        If ans = "q" Then
            WScript.Echo ""
            WScript.Echo "[退出] 请先备份原始文件后再运行。"
            Call PauseExit()
            Exit Sub
        End If
        WScript.Echo ">> 输入无效，请输入 1 或 Q"
    Loop
    
    ' 模式选择
    Dim deleteOriginal
    deleteOriginal = False
    Do
        WScript.Echo ""
        WScript.Echo "【操作选项】"
        WScript.Echo "  1. 转换后保留原 .xls 文件（安全，推荐）"
        WScript.Echo "  2. 转换后删除原 .xls 文件（危险，仅确认备份后使用）"
        WScript.StdOut.Write "请输入选项 (1 或 2): "
        ans = Trim(WScript.StdIn.ReadLine())
        If ans = "1" Then
            deleteOriginal = False
            Exit Do
        ElseIf ans = "2" Then
            deleteOriginal = True
            Exit Do
        End If
        WScript.Echo ">> 输入无效，请重新输入。"
    Loop
    
    ' 最终确认
    Dim modeStr
    If deleteOriginal Then
        modeStr = "【删除原文件】"
    Else
        modeStr = "【保留原文件】"
    End If
    
    Do
        WScript.Echo ""
        WScript.Echo "[执行确认] " & modeStr & " 即将处理 " & xlsFiles.Count & " 个 .xls 文件"
        WScript.Echo "是否开始转换？"
        WScript.Echo "  [1] 确认开始"
        WScript.Echo "  [2] 返回上一级"
        WScript.Echo "  [Q] 退出程序"
        WScript.StdOut.Write "请输入: "
        ans = LCase(Trim(WScript.StdIn.ReadLine()))
        If ans = "1" Then Exit Do
        If ans = "q" Then
            WScript.Echo ""
            WScript.Echo "[退出] 操作已取消。"
            Call PauseExit()
            Exit Sub
        End If
        ' ans = "2" 则重新选择模式
        Do
            WScript.Echo ""
            WScript.Echo "【操作选项】"
            WScript.Echo "  1. 转换后保留原 .xls 文件（安全，推荐）"
            WScript.Echo "  2. 转换后删除原 .xls 文件（危险，仅确认备份后使用）"
            WScript.StdOut.Write "请输入选项 (1 或 2): "
            ans = Trim(WScript.StdIn.ReadLine())
            If ans = "1" Then
                deleteOriginal = False
                Exit Do
            ElseIf ans = "2" Then
                deleteOriginal = True
                Exit Do
            End If
            WScript.Echo ">> 输入无效，请重新输入。"
        Loop
    Loop
    
    ' 启动 Excel
    WScript.Echo ""
    WScript.Echo "正在启动 Microsoft Excel..."
    
    Dim excel
    On Error Resume Next
    Set excel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        WScript.Echo "[错误] 无法启动 Microsoft Excel。请确认已安装 Excel。"
        WScript.Echo "详细错误: " & Err.Description
        Call PauseExit()
        Exit Sub
    End If
    On Error GoTo 0
    
    excel.DisplayAlerts = False
    excel.Visible = False
    excel.ScreenUpdating = False
    
    ' 执行转换
    WScript.Echo String(64, "-")
    WScript.Echo "开始转换..."
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
        
        ' 显示进度（相对路径）
        If Len(srcPath) > Len(rootDir) + 1 Then
            relPath = Mid(srcPath, Len(rootDir) + 2)
        Else
            relPath = fileName
        End If
        WScript.Echo "  [" & Right("   " & (i + 1), 3) & "/" & Right("   " & xlsFiles.Count, 3) & "] " & relPath
        
        ' 检查目标文件是否已存在
        If fso.FileExists(dstPath) Then
            WScript.Echo "      -> 跳过（目标文件已存在）"
            skipped = skipped + 1
        Else
            On Error Resume Next
            Set wb = excel.Workbooks.Open(srcPath)
            If Err.Number <> 0 Then
                failedList.Add failedList.Count, Array(srcPath, "打开文件失败: " & Err.Description)
                WScript.Echo "      -> 失败（打开文件错误）"
                Err.Clear
            Else
                wb.SaveAs dstPath, XL_FILE_FORMAT_XLSX
                If Err.Number <> 0 Then
                    failedList.Add failedList.Count, Array(srcPath, "保存文件失败: " & Err.Description)
                    WScript.Echo "      -> 失败（保存错误）"
                    Err.Clear
                Else
                    success = success + 1
                    WScript.Echo "      -> 成功"
                End If
                wb.Close SaveChanges=False
                Set wb = Nothing
            End If
            On Error GoTo 0
        End If
        
        ' 删除原文件（仅在转换成功且用户选择删除时）
        If deleteOriginal And fso.FileExists(dstPath) Then
            On Error Resume Next
            fso.DeleteFile srcPath, True
            If Err.Number <> 0 Then
                WScript.Echo "      -> 警告: 无法删除原文件 - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
    
    ' 关闭 Excel
    On Error Resume Next
    excel.Quit
    Set excel = Nothing
    On Error GoTo 0
    
    ' 显示汇总
    WScript.Echo ""
    Call ShowSummary(success, skipped, failedList)
    
    Call PauseExit()
End Sub

' --------------------------------------------------------------------------------
' 递归扫描 .xls 文件
' --------------------------------------------------------------------------------
Sub ScanXlsFiles(fso, folderPath, fileList)
    Dim folder, file, subFolder
    Set folder = fso.GetFolder(folderPath)
    
    ' 遍历当前目录的文件
    For Each file In folder.Files
        ' 严格匹配 .xls 后缀，排除 .xlsx
        If LCase(fso.GetExtensionName(file.Name)) = "xls" Then
            fileList.Add fileList.Count, file.Path
        End If
    Next
    
    ' 递归遍历子目录
    For Each subFolder In folder.SubFolders
        Call ScanXlsFiles(fso, subFolder.Path, fileList)
    Next
End Sub

' --------------------------------------------------------------------------------
' 格式化文件大小
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
' 显示标题
' --------------------------------------------------------------------------------
Sub ShowBanner()
    WScript.Echo ""
    WScript.Echo "================================================================================"
    WScript.Echo "        XLS -> XLSX 批量转换工具 (VBScript + Excel COM)"
    WScript.Echo "================================================================================"
    WScript.Echo "【技术特点】"
    WScript.Echo "  . 基于 Microsoft Excel COM 接口，保留 100% 原始格式"
    WScript.Echo "  . 只需安装 Excel，无需 Python 或其他依赖"
    WScript.Echo "  . 递归扫描目录，批量处理，自动跳过已存在文件"
    WScript.Echo "  . 完善的备份警告与确认机制，防止误操作"
    WScript.Echo "================================================================================"
End Sub

' --------------------------------------------------------------------------------
' 显示扫描报告
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
    WScript.Echo "[扫描结果]"
    WScript.Echo "  待转换文件 : " & fileList.Count & " 个"
    WScript.Echo "  涉及目录   : " & dirSet.Count & " 个"
    WScript.Echo "  总大小     : " & FormatSize(totalSize)
    
    ' 显示前 10 个文件
    Dim maxShow
    maxShow = 10
    If fileList.Count < maxShow Then maxShow = fileList.Count
    
    WScript.Echo ""
    WScript.Echo "[文件清单] 前 " & maxShow & " 个:"
    For i = 0 To maxShow - 1
        If Len(fileList(i)) > Len(rootDir) + 1 Then
            relPath = Mid(fileList(i), Len(rootDir) + 2)
        Else
            relPath = fso.GetFileName(fileList(i))
        End If
        WScript.Echo "  " & Right("  " & (i + 1), 2) & ". " & relPath
    Next
    If fileList.Count > 10 Then
        WScript.Echo "      ... 还有 " & (fileList.Count - 10) & " 个文件"
    End If
End Sub

' --------------------------------------------------------------------------------
' 显示警告
' --------------------------------------------------------------------------------
Sub ShowWarning()
    WScript.Echo ""
    WScript.Echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    WScript.Echo "  !   重 要 警 告"
    WScript.Echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    WScript.Echo "  1. 本工具通过 Excel COM 接口执行另存为操作，"
    WScript.Echo "     转换质量与手动在 Excel 中'另存为'完全一致。"
    WScript.Echo "  2. 转换过程会生成新的 .xlsx 文件，不会直接覆盖原 .xls。"
    WScript.Echo "  3. 但若选择'删除原文件'模式，原始 .xls 将被永久删除！"
    WScript.Echo "  4. 强烈建议在运行前对原始文件进行完整备份！"
    WScript.Echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
End Sub

' --------------------------------------------------------------------------------
' 显示汇总
' --------------------------------------------------------------------------------
Sub ShowSummary(success, skipped, failedList)
    WScript.Echo "================================================================================"
    WScript.Echo "                     处 理 结 果 汇 总"
    WScript.Echo "================================================================================"
    WScript.Echo "  . 成功转换 : " & success & " 个"
    WScript.Echo "  . 跳过(已存在) : " & skipped & " 个"
    WScript.Echo "  . 转换失败 : " & failedList.Count & " 个"
    
    If failedList.Count > 0 Then
        WScript.Echo "--------------------------------------------------------------------------------"
        WScript.Echo "【失败明细】"
        Dim i, item
        For i = 0 To failedList.Count - 1
            item = failedList(i)
            WScript.Echo "  . " & item(0)
            WScript.Echo "    原因: " & item(1)
        Next
    End If
    WScript.Echo "================================================================================"
End Sub

' --------------------------------------------------------------------------------
' 暂停退出
' --------------------------------------------------------------------------------
Sub PauseExit()
    WScript.Echo ""
    WScript.StdOut.Write "按回车键退出..."
    Dim dummy
    dummy = WScript.StdIn.ReadLine()
    WScript.Quit
End Sub
