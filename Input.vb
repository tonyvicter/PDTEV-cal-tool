Public fn_Input As Integer '已选文件个数
Public V_list As Integer

Private Sub HistoryClr_Click()
    
    Sheets("L-ok TEV").UsedRange.Clear
    Sheets("L-ok TEV").Cells(2, 2) = "time"
    
    Sheets("L-close TEV").UsedRange.Clear
    Sheets("L-close TEV").Cells(2, 2) = "time"
    
    Sheets("L-open TEV").UsedRange.Clear
    Sheets("L-open TEV").Cells(2, 2) = "time"
    
    Sheets("H-ok TEV").UsedRange.Clear
    Sheets("H-ok TEV").Cells(2, 2) = "time"
    
    Sheets("H-close TEV").UsedRange.Clear
    Sheets("H-close TEV").Cells(2, 2) = "time"
    
    Sheets("H-open TEV").UsedRange.Clear
    Sheets("H-open TEV").Cells(2, 2) = "time"
    
End Sub

Private Sub Worksheet_Activate()
    '打开工具之后的初始化操作：获得已选文件的个数fn_Input
    
    Set rg_Input = Sheets("Input").UsedRange
    Set c_Input = rg_Input.Find("编号", lookat:=xlWhole)
    
    If Not c_Input Is Nothing Then
        fn_Input = c_Input.Offset(-1, 2)
    Else
        MsgBox ("找不到内容为“编号”的单元格！")
        End
    End If
    
End Sub

Private Sub FilesAdd_Click()
    '选择并添加测量文件--获得文件路径
    
    Dim fd_Input As FileDialog
    Dim fn_Input_temp As Integer
    Dim i As Integer
    Dim a As Integer
    Dim text_Input As String
    Dim len_Input As Integer
    
    Set rg_Input = Sheets("Input").UsedRange
    Set c_Input = rg_Input.Find("编号", lookat:=xlWhole)
    
    Set fd_Input = Application.FileDialog(msoFileDialogFilePicker)
    With fd_Input
        .Title = "请选择测量文件"
        .Filters.Add "测量文件", "*.dat", 1
        .Filters.Add "所有文件(*.*)", "*.*", 2
        .FilterIndex = 1
        .InitialFileName = ActiveWorkbook.Path
        .AllowMultiSelect = True

    End With
    
    If fd_Input.Show = -1 Then
        fn_Input = fd_Input.SelectedItems.Count
        'c_Input.Offset(i - 1, 2) = fn_Input -- 用于记录已选文件的个数，目前使用表格自带的公式来实现，暂时不使用该段代码
    Else
        fn_Input = 0
    End If
    
    If fn_Input = 0 Then
        MsgBox ("请至少选择一个测量文件！")
        End
    Else
        If Not c_Input Is Nothing Then
            fn_Input_temp = c_Input.Offset(-1, 2) '读取之前选择的文件个数
            
            If fn_Input_temp > 0 Then
                Range(c_Input.Offset(1, 0), c_Input.Offset(fn_Input, 2)) = "" '清空之前选择的文件
            End If
            
            For i = 0 To fn_Input - 1
                c_Input.Offset(i + 1, 0) = i + 1
                
                text_Input = fd_Input.SelectedItems(i + 1)
                
                c_Input.Offset(i + 1, 1) = text_Input
                
                len_Input = Len(text_Input)
                a = InStrRev(text_Input, "\")
                c_Input.Offset(i + 1, 2) = Right(text_Input, len_Input - a)
            Next
        Else
            MsgBox ("找不到内容为“编号”的单元格！")
            End
        End If
    End If
    
    Range("I2").Calculate
    
End Sub

Private Sub FilesClr_Click()
    '清空当前已选择的文件
    
    Set rg_Input = Sheets("Input").UsedRange
    Set c_Input = rg_Input.Find("编号", lookat:=xlWhole)
    
    fn_Input = c_Input.Offset(-1, 2)
    
    If fn_Input = 0 Then
        MsgBox ("当前未选择任何文件，无需清空。")
    Else
        If Not c_Input Is Nothing Then
            Range(c_Input.Offset(1, 0), c_Input.Offset(fn_Input, 2)) = ""
        Else
            MsgBox ("找不到内容为“编号”的单元格！")
            End
        End If
    End If
    
    Range("I2").Calculate
    
End Sub

Private Sub ShowConf_Click()
    '打开窗口并初始化窗口信息
    
    Range("I2").Calculate
    
    Set rg_Input = Sheets("Input").UsedRange
    Set c_Input = rg_Input.Find("编号", lookat:=xlWhole)
    
    fn_Input = c_Input.Offset(-1, 2)
    
    '载入dat文件列表
    If fn_Input > 0 Then
        For i = 0 To fn_Input - 1
            'AnaConf.DatBox.AddItem c_Input.Offset(i + 1, 2)
            Files.RoadTestList.AddItem c_Input.Offset(i + 1, 2)
        Next i
        
        Files.RoadTestList.ListIndex = 0    '默认显示第一个文件
        Files.OptionL.Value = True          '默认为单管路配置
        
        Files.Show
    Else
        MsgBox ("请至少添加一个测量文件。")
    End If
    'dat文件列表载入完成
    
End Sub