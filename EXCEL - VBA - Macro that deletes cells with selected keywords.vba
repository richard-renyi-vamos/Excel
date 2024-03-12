Sub DeleteCellsByKeywords()
    Dim rng As Range
    Dim cell As Range
    Dim deleteKeywords() As String
    Dim keyword As Variant
    Dim cellValue As String
    
    ' Define your list of keywords to search for
    deleteKeywords = Array("keyword1", "keyword2", "keyword3") ' Add or remove keywords as needed
    
    ' Specify the range to search within
    Set rng = ActiveSheet.UsedRange ' Change to a specific range if needed
    
    ' Loop through each cell in the range
    For Each cell In rng
        cellValue = cell.Value
        ' Loop through each keyword
        For Each keyword In deleteKeywords
            ' Check if the keyword is found in the cell value
            If InStr(1, cellValue, keyword, vbTextCompare) > 0 Then
                ' Delete the cell if a keyword is found
                cell.ClearContents
                Exit For ' Exit the loop after deleting the cell
            End If
        Next keyword
    Next cell
End Sub
