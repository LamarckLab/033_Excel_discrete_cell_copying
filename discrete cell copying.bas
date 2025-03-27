Attribute VB_Name = "ģ��1"
Sub PasteToVisibleCells()
    Dim SourceRange As Range
    Dim TargetRange As Range
    Dim VisibleCells As Range
    Dim Cell As Range
    Dim i As Long

    ' ��ȡ��������ʵ�ʸ��Ƶ�����
    On Error Resume Next
    Set SourceRange = Application.InputBox("���ȸ���Դ���ݣ�Ȼ���������ȷ��", Type:=8)
    On Error GoTo 0
    
    If SourceRange Is Nothing Then
        MsgBox "δѡ��Դ�������򣬲�����ȡ����", vbExclamation
        Exit Sub
    End If

    ' ��ȡ��ǰѡ�еĿɼ�Ŀ�굥Ԫ������
    On Error Resume Next
    Set VisibleCells = Selection.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If VisibleCells Is Nothing Then
        MsgBox "����ѡ��Ŀ������ɸѡ��Ŀɼ���Ԫ�񣩣�", vbExclamation
        Exit Sub
    End If

    ' У������
    If SourceRange.Cells.Count > VisibleCells.Cells.Count Then
        MsgBox "Դ�����������ڿɼ�Ŀ�굥Ԫ�����飡", vbExclamation
        Exit Sub
    End If

    ' ճ������
    i = 1
    For Each Cell In VisibleCells
        Cell.Value = SourceRange.Cells(i).Value
        i = i + 1
        If i > SourceRange.Cells.Count Then Exit For
    Next Cell

    MsgBox "ճ����ɣ�", vbInformation
End Sub

