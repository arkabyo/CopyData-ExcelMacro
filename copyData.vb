Sub CopyData()
    ' Declare variables to hold the workbooks
    Dim HostWb As Workbook
    Dim DataWb As Workbook
    
    ' Set the HostWb variable to the workbook containing this code
    Set HostWb = ThisWorkbook
    
    ' Loop through all open workbooks to find the Data workbook
    For Each wb In Workbooks
        ' Check if the current workbook is not the Host workbook
        If wb.Name <> HostWb.Name Then
            ' Set the DataWb variable to the current workbook
            Set DataWb = wb
            ' Exit the loop
            Exit For
        End If
    Next wb
    
    ' Check if the DataWb variable was set
    If DataWb Is Nothing Then
        ' Display an error message if the Data workbook was not found
        MsgBox "Data workbook not found!", vbCritical
        ' Exit the macro
        Exit Sub
    End If
    
    ' Declare variables to hold the worksheets
    Dim HostWs As Worksheet
    Dim DataWs As Worksheet
    
    ' Set the HostWs variable to the first worksheet in the Host workbook
    Set HostWs = HostWb.Worksheets(1)
    
    ' Set the DataWs variable to the first worksheet in the Data workbook
    Set DataWs = DataWb.Worksheets(1)
    
    ' Find the last row and column numbers in both worksheets
    Dim HostLastRow As Long, HostLastCol As Long, DataLastRow As Long, DataLastCol As Long
    
    HostLastRow = HostWs.Cells(Rows.Count, 1).End(xlUp).Row
    HostLastCol = HostWs.Cells(1, Columns.Count).End(xlToLeft).Column
    DataLastRow = DataWs.Cells(Rows.Count, 1).End(xlUp).Row
    DataLastCol = DataWs.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Declare variables to hold column numbers in both worksheets
    Dim HostCol As Long, DataCol As Long
    
    ' Loop through all columns in both worksheets
    For HostCol = 1 To HostLastCol
        For DataCol = 1 To DataLastCol
            ' Check if the column headers match (case-insensitive)
            If StrComp(HostWs.Cells(1, HostCol).Value, DataWs.Cells(1, DataCol).Value, vbTextCompare) = 0 Then
                ' Resize the destination range in the Host worksheet to match the size of the source range in the Data worksheet
                Dim copyRange As Range
                Set copyRange = DataWs.Range(DataWs.Cells(2, DataCol), DataWs.Cells(DataLastRow, DataCol))
                HostWs.Cells(2, HostCol).Resize(copyRange.Rows.Count, copyRange.Columns.Count).Value = copyRange.Value
                ' Exit inner loop (no need to check remaining columns in Data worksheet)
                Exit For
            End If
        Next DataCol
    Next HostCol

    MsgBox "Data copied successfully!", vbInformation
End Sub
