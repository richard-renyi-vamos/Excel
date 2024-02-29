Sub SplitFirstTenColumns()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim newWorkbook As Workbook
    Dim i As Integer
    
    ' Set reference to the source worksheet
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your source sheet
    
    ' Create a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Loop through columns 1 to 10 and copy data to new workbook
    For i = 1 To 10
        ' Copy the column from source worksheet to new worksheet
        wsSource.Columns(i).Copy
        
        ' Paste the column into new workbook
        Set wsNew = newWorkbook.Sheets.Add(After:=newWorkbook.Sheets(newWorkbook.Sheets.Count))
        wsNew.Paste
        wsNew.Name = "Column_" & i
    Next i
    
    ' Delete the default sheet in the new workbook
    Application.DisplayAlerts = False ' Turn off display alerts
    newWorkbook.Sheets(1).Delete
    Application.DisplayAlerts = True ' Turn on display alerts
    
    ' Save the new workbook
    newWorkbook.SaveAs "C:\Path\To\Save\NewWorkbook.xlsx" ' Change the file path and name as needed
    newWorkbook.Close ' Close the new workbook
    
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' Inform the user
    MsgBox "First 10 columns have been split into a separate workbook.", vbInformation
End Sub
