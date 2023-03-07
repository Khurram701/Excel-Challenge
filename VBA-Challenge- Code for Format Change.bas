Attribute VB_Name = "Module1"
Sub FormatYearlyChange()
    Dim LastRow As Long
    Dim cell As Range
    
  
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
   
    For Each cell In Range("J2:J" & LastRow)
        
        If cell.Value > 0 Then
           
            cell.Interior.Color = RGB(0, 255, 0)
        ElseIf cell.Value < 0 Then
           
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next cell
End Sub




