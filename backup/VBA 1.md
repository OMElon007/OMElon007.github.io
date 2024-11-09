' ****************************************创建/打开文件夹****************************************
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' 创建String变量存储文件夹地址
    Dim folderPath As String
    ' 创建String变量存储打开文件夹地址
    Dim fullPath As String

    ' 确保只在点击 H 列单元格时执行
    If Not Intersect(Target, Me.Columns("H")) Is Nothing And Target.Cells.count = 1 Then
        ' 获取文件夹路径
        folderPath = Me.Range("H1").Value
        ' 构建完整路径，使用对应行的 B 列和 D 列的值
        fullPath = folderPath & Me.Cells(Target.Row, 2).Value & " " & Me.Cells(Target.Row, 4).Value
        
        ' 若无法打开文件夹则跳到CreateFolder标签
        On Error GoTo CreateFolder
        ' 尝试打开文件夹
        ' ActiveWorkbook指当前活动工作簿
        ActiveWorkbook.FollowHyperlink Address:=fullPath, NewWindow:=True
        ' 如果成功打开文件夹，则退出子程序，后续代码不再执行
        Exit Sub

CreateFolder:
        ' 如果出错则继续执行，而不是中断
        On Error Resume Next
        ' 创建文件夹，已有则会出错
        MkDir fullPath
        '用资源管理器打开新建文件夹
        Shell "explorer.exe " & fullPath, vbNormalFocus
    End If
End Sub
' ****************************************点击链接打开并复制****************************************
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    ' 检查是否是 E 列
    If Target.Range.Column = 5 Then
        ' 获取显示的文字
        Dim linkText As String
        ' Hyperlink.TextToDisplay为超链接中显示的文字
        linkText = Target.TextToDisplay

        ' 复制显示的文字到剪贴板
        ' 此处调用模块中的SetClipboardText过程
        SetClipboardText linkText
    End If
End Sub
' ****************************************分类高亮****************************************
' 更新并重新计算C列公式
Sub UpdateAndRecalculateColumnC()
    Dim lastRow As Long
    Dim i As Long
    
    ' 找到B列中的最后一行
    ' Rows.count为当前活动工作表的最大行数，1048576
    ' Cells(Rows.count, "B")为定位到B列的最后一个格子，即单元格B1048576
    ' End(xlUp):沿着B列最后一个格子，向上寻找第一次出现的非空单元格
    ' .Row取行号
    lastRow = Cells(Rows.count, "B").End(xlUp).Row
    
    ' 从第 2 行开始循环，更新每行 C 列的公式
    For i = 2 To lastRow
        If Cells(i, "A").Value <> "" And Cells(i, "B").Value <> "" Then
            ' 动态设置每行的C列公式
            Cells(i, "C").Formula = "=IF(B" & i - 1 & "=B" & i & ",C" & i - 1 & ",C" & i - 1 & "+1)"
        End If
    Next i
    
    ' 重新计算C列
    Range("C:C").Calculate
End Sub
' 当A列或B列发生变化时自动触发UpdateAndRecalculateColumnC更新并重新计算C、H列公式
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 检查是否在A列或B列中更改了单元格
    If Not Intersect(Target, Me.Columns("A")) Is Nothing Or Not Intersect(Target, Me.Columns("B")) Is Nothing Then
        Dim rowNum As Long
        rowNum = Target.Row
        
        ' 确保当前行A列和B列都有值，且不是第一行
        If rowNum > 1 And Cells(rowNum, "A").Value <> "" And Cells(rowNum, "B").Value <> "" Then
            ' 设置H列的公式
            Cells(rowNum, "H").Formula = "=HYPERLINK($H$1&B" & rowNum & "&"" ""&D" & rowNum & ", $D$1)"
            Cells(rowNum, "H").Font.Color = Cells(rowNum - 1, "H").Font.Color
            Cells(rowNum, "H").Font.Underline = xlUnderlineStyleNone ' 移除下划线
            
            ' 调用UpdateAndRecalculateColumnC，更新C列公式并重新计算
            UpdateAndRecalculateColumnC
        End If
    End If
End Sub