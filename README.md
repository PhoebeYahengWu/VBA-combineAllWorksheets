# VBA-combineAllWorksheets




```
Sub combine()
    Sheets.Add.Name = "Combined_Data"
    
    'move the created sheet to be the first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    
    Set combined_sheet = Worksheets("Combined_Data")
    
    For Each ws In Worksheets
    
        'find the last row of the combined sheet after each paste
        'add 1 to get the first empty row
        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        'find the last row of each worksheet
        'subtract one to return the number of rows without header
        lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        
        'copy the contents of each state sheet into the combined sheet
        combined_sheet.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
        
    Next ws
    
    'copy the headers
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    combined_sheet.Columns("A:G").AutoFit
    
End Sub
```
