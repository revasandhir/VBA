Attribute VB_Name = "Module1"
Sub ListCustomersOwingMoreThan1000()
    Dim wsData As Worksheet, wsResults As Worksheet
    Dim lastRow As Long, resultRow As Long
    Dim customerID As Range, amountPurchased As Range, amountPaid As Range
    Dim amountOwed As Double

    ' Set references to the Data and Results sheets
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsResults = ThisWorkbook.Sheets("Results")
    
    ' Clear the Results sheet
    wsResults.Rows("2:" & wsResults.Rows.Count).ClearContents
    
    ' Find the last row with data in the Data sheet
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize the resultRow for the Results sheet (starting at row 2)
    resultRow = 2
    
    ' Loop through each row of the Data sheet
    For Each customerID In wsData.Range("A2:A" & lastRow)
        Set amountPurchased = customerID.Offset(0, 1)
        Set amountPaid = customerID.Offset(0, 2)
        
        ' Check if both "Amount purchased" and "Amount paid" are numeric
        If IsNumeric(amountPurchased.Value) And IsNumeric(amountPaid.Value) Then
            ' Calculate the amount owed
            amountOwed = amountPurchased.Value - amountPaid.Value
            
            ' If the amount owed is more than $1,000, add to Results sheet
            If amountOwed > 1000 Then
                wsResults.Cells(resultRow, 1).Value = customerID.Value
                wsResults.Cells(resultRow, 2).Value = amountOwed
                resultRow = resultRow + 1
            End If
        End If
    Next customerID
End Sub

