Attribute VB_Name = "Module1"
Sub CalculateYearlyChange()
    
    ' Define variables
    Dim Ticker As String
    Dim lastRow As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim SummaryTableRowIndex As Integer
    Dim PercentChange As Double
   
    
    ' Set initial values
    SummaryTableRowIndex = 1
    OpenPrice = 0
    ClosePrice = 0
    
    ' Find the last row of data
    lastRow = Range("A1").End(xlDown).Row
    
    ' Loop through each row of data
    For i = 2 To lastRow
        
        ' Check if the ticker symbol has changed
        If Cells(i, 1).Value <> Ticker Then
            
            ' Calculate the yearly change and add it to the summary table
            If OpenPrice <> 0 Then
                YearlyChange = ClosePrice - OpenPrice
                Cells(SummaryTableRowIndex, 10).Value = YearlyChange
               
                
                PercentChange = YearlyChange / OpenPrice
                Cells(SummaryTableRowIndex, 11).Value = FormatPercent(PercentChange, 11)
            End If
            
            ' Set the new ticker symbol and open price
            Ticker = Cells(i, 1).Value
            OpenPrice = Cells(i, 3).Value
            SummaryTableRowIndex = SummaryTableRowIndex + 1
            
        End If
        
        ' Set the close price for the current row
        ClosePrice = Cells(i, 6).Value
        
    Next i
    
    ' Calculate the yearly change and add it to the summary table for the last ticker symbol
    YearlyChange = ClosePrice - OpenPrice
   PercentChange = YearlyChange / OpenPrice
   
   Cells(SummaryTableRowIndex, 10).Value = YearlyChange
   Cells(SummaryTableRowIndex, 11).Value = FormatPercent(PercentChange, 11)
   
   
   

   
   
    
End Sub


