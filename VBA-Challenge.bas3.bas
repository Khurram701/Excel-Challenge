Attribute VB_Name = "Module1"
Sub CopyUniqueTickers()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickers As Range
    Dim uniqueTickers As Range
    
   
    Set ws = ThisWorkbook.Worksheets("2020")
    
  
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
   
    Set tickers = ws.Range("A1:A" & lastRow)
    
   
    Set uniqueTickers = tickers.Cells(1, 1).Resize(tickers.Cells.Count, 1). _
        AdvancedFilter(Action:=xlFilterCopy, CopyToRange:=ws.Range("I1"), Unique:=True)
        
           
End Sub


