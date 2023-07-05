Sub ResetQuantity()
    Dim answer As Integer
    answer = MsgBox("Are you sure you want to reset the Quantity column?", vbYesNo + vbQuestion, "Reset Quantity")
    
    If answer = vbYes Then 'Resetting Quantity
        Dim lastRow As Integer
        lastRow = Range("A" & Rows.Count).End(xlUp).row
        
        'Reset the quantity to zero for each item
        Range("A2:A" & lastRow).value = 0
        
        MsgBox "The Quantity column has been reset to zero."
    End If
End Sub

