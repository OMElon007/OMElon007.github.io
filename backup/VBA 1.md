## 创建/打开文件夹
> Note
>
> #### 本段代码作用：
``` VBA
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
```