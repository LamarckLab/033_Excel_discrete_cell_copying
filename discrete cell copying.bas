Attribute VB_Name = "模块1"
Sub PasteToVisibleCells()
    Dim SourceRange As Range
    Dim TargetRange As Range
    Dim VisibleCells As Range
    Dim Cell As Range
    Dim i As Long

    ' 获取剪贴板中实际复制的区域
    On Error Resume Next
    Set SourceRange = Application.InputBox("请先复制源数据，然后在这里点确定", Type:=8)
    On Error GoTo 0
    
    If SourceRange Is Nothing Then
        MsgBox "未选择源数据区域，操作已取消。", vbExclamation
        Exit Sub
    End If

    ' 获取当前选中的可见目标单元格区域
    On Error Resume Next
    Set VisibleCells = Selection.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If VisibleCells Is Nothing Then
        MsgBox "请先选中目标区域（筛选后的可见单元格）！", vbExclamation
        Exit Sub
    End If

    ' 校验数量
    If SourceRange.Cells.Count > VisibleCells.Cells.Count Then
        MsgBox "源数据行数多于可见目标单元格，请检查！", vbExclamation
        Exit Sub
    End If

    ' 粘贴数据
    i = 1
    For Each Cell In VisibleCells
        Cell.Value = SourceRange.Cells(i).Value
        i = i + 1
        If i > SourceRange.Cells.Count Then Exit For
    Next Cell

    MsgBox "粘贴完成！", vbInformation
End Sub

