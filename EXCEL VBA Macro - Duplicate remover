Sub RemoveDuplicatesFromColumnA()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim uniqueValues As New Collection
    Dim item As Variant
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there are values in column A
    If lastRow < 1 Then
        MsgBox "No data found in column A. Operation canceled.", vbExclamation
        Exit Sub
    End If
    
    ' Set the range to column A
    Set rng = ws.Range("A1:A" & lastRow)
    
    ' Loop through each cell in column A
    For Each cell In rng
        ' Add unique values to the collection
        If Not IsEmpty(cell.Value) Then
            On Error Resume Next
            uniqueValues.Add cell.Value, CStr(cell.Value)
            On Error GoTo 0
        End If
    Next cell
    
    ' Clear column A
    rng.ClearContents
    
    ' Write the unique values back to column A
    For Each item In uniqueValues
        ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = item
    Next item
    
    MsgBox "Duplicates removed successfully from column A!", vbInformation
End Sub
