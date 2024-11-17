## 创建/打开文件夹
> #### 本段代码作用：
> 点击H列处的链接，若链接指向的文件夹存在，则打开文件夹，若不存在，则自动在指定存储文件夹处建立指定命名的文件夹并打开。

> #### 存储文件夹和自动建立的文件夹命名逻辑可自定义

``` VBA
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' 创建String变量存储文件夹地址
    Dim folderPath As String
    ' 创建String变量存储打开文件夹地址
    Dim fullPath As String

    ' 确保只在点击 H 列单元格时执行
    ' 可修改为自己想要的其他列
    If Not Intersect(Target, Me.Columns("H")) Is Nothing And Target.Cells.count = 1 Then
        ' 获取存储文件夹路径
        folderPath = Me.Range("H1").Value
        ' 构建完整路径，使用对应行的 B 列和 D 列的值
        ' 完整路径的命名逻辑是：存储文件夹路径+对应行B列值+空格+对应行D列值
        ' 可修改为自己需要的其他命名逻辑
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