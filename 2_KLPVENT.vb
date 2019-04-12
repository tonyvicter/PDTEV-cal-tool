'msoFileDialogFilePicker 允许用户选择一个文件
'msoFileDialogFolderPicker 允许用户选择一个文件夹
'msoFileDialogOpen 允许用户打开一个文件，选取多个文件
'msoFileDialogSaveAs 允许用户保存一个文件

Public fn As Integer                '文件个数
Public rg As Range, c As Range      '定位
Public ifcheck As Boolean           '是否是对比计算


Private Sub Calcu_run_Click()
    Dim i As Integer        '循环用
    Dim fadd As Integer     '程序正式运行之前统计的已有文件数
    
    Set rg = Sheets("KLPVENT").UsedRange
    Set c = rg.Find("路径", lookat:=xlWhole)
    
    fadd = 0
    ifcheck = False
    
    If Not c Is Nothing Then
        For i = 1 To 3
            If c.Offset(i, 0) <> "" Then
                fadd = fadd + 1
            End If
        Next i
    End If
    
    If fadd > 0 Then
        msg = MsgBox("好像已经选择了测量文件了，是否直接开始计算？", vbYesNoCancel)
        If msg = vbCancel Then
            End
        End If
        
        If msg = vbYes Then
            fn = fadd
            Call Calculate_pssp
        End If
        
        If msg = vbNo Then
            Set fd = Application.FileDialog(msoFileDialogFilePicker)
            With fd
                .AllowMultiSelect = True
                .Filters.Clear
                .Title = "选择测量文件"
                .Filters.Add "测量文件", "*.dat"
                .InitialFileName = ActiveWorkbook.Path
            End With
            
            If fd.Show = -1 Then
                fn = fd.SelectedItems.Count
            End If
            
            If fn = 0 Then
                MsgBox ("请至少选择一个测量文件！")
                End
            Else
                If Not c Is Nothing Then
                    For i = 0 To fn - 1
                        c.Offset(i + 1, 0) = ""
                        c.Offset(i + 1, 0) = fd.SelectedItems(i + 1)
                    Next
                    MsgBox ("文件载入完成，如果选择了多个文件，程序只载入前三个。")
                Else
                    MsgBox ("找不到单元格“路径”！请在“D2”单元格内写入“路径”然后重新运行程序。")
                    End
                End If
                Call Calculate_pssp
            End If
        End If
    Else
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .AllowMultiSelect = True
            .Filters.Clear
            .Title = "选择测量文件"
            .Filters.Add "测量文件", "*.dat"
            .InitialFileName = ActiveWorkbook.Path
        End With
        
        If fd.Show = -1 Then
            fn = fd.SelectedItems.Count
        End If
        
        If fn = 0 Then
            MsgBox ("请至少选择一个测量文件！")
            End
        Else
            If Not c Is Nothing Then
                For i = 0 To fn - 1
                    c.Offset(i + 1, 0) = ""
                    c.Offset(i + 1, 0) = fd.SelectedItems(i + 1)
                Next
                MsgBox ("文件载入完成，如果选择了多个文件，程序只载入前三个。")
            Else
                MsgBox ("找不到单元格“路径”！请在“D2”单元格内写入“路径”然后重新运行程序。")
                End
            End If
            Call Calculate_pssp
        End If
    End If
End Sub

Private Sub Check_KLPVENT_Click()
'
    Dim i As Integer        '循环用
    Dim fadd As Integer     '程序正式运行之前统计的已有文件数
    
    Set rg = Sheets("KLPVENT").UsedRange
    Set c = rg.Find("路径", lookat:=xlWhole)
    
    fadd = 0
    ifcheck = True
    
    If Not c Is Nothing Then
        For i = 1 To 3              '清空已选文件
            c.Offset(i, 0) = ""
        Next i
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Filters.Clear
        .Title = "选择测量文件"
        .Filters.Add "测量文件", "*.dat"
        .InitialFileName = ActiveWorkbook.Path
    End With
        
    If fd.Show = -1 Then
        fn = fd.SelectedItems.Count
    End If
        
    If fn = 0 Then
        MsgBox ("请选择一个测量文件！")
        End
    Else
        If Not c Is Nothing Then
            c.Offset(1, 0) = ""
            c.Offset(1, 0) = fd.SelectedItems(1)
        Else
            MsgBox ("找不到单元格“路径”！请在“D2”单元格内写入“路径”然后重新运行程序。")
            End
        End If
        Call Calculate_pssp
    End If
End Sub

Private Sub Load_clr_Click()
    '清空KLPVENT中的文件列表
    Set rg = Sheets("KLPVENT").UsedRange
    Set c = rg.Find("路径", lookat:=xlWhole)
    
    If Not c Is Nothing Then
        For i = 0 To 2
            c.Offset(i + 1, 0) = ""
        Next i
    End If
    
End Sub

Private Sub Readme_KLPVENT_Click()
    '说明窗口
    hlp.Show
    
End Sub
