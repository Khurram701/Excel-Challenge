Attribute VB_Name = "Module1"
Sub CalculateTotalVolume()
    Dim lastRow As Long
    Dim ticker As String
    Dim volume As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    
    ' Find the last row of data in column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set the initial values for the ticker and total volume
    ticker = Range("A2")
    totalVolume = Range("G2")
    outputRow = 2
    
    ' Loop through each row of data
    For i = 3 To lastRow
        ' Check if this is the same ticker as the previous row
        If Range("A" & i) = ticker Then
            ' If it is, add the volume to the total
            volume = Range("G" & i)
            totalVolume = totalVolume + volume
        Else
            ' If it isn't, write the total volume for the previous ticker and reset the total volume for the new ticker
            Range("L" & outputRow) = totalVolume
            outputRow = outputRow + 1
            ticker = Range("A" & i)
            totalVolume = Range("G" & i)
        End If
    Next i
    
    ' Write the total volume for the final ticker
    Range("L" & outputRow) = totalVolume
End Sub



