Private Sub Quantity_Textbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' Get the quantity from the Quantity_Textbox
        Dim Quantity As Integer
        Quantity = Val(Quantity_Textbox.Text)
        
        ' Update the Inventory table with the quantity
        Dim productID As String
        productID = Scanned_IN_Textbox.Text
        
        ' Find the corresponding row in the Inventory table
        Dim foundCell As Range
        Set foundCell = Sheets("Inventory").Range("B2:B100").Find(productID, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            ' Update the quantity in the Quantity column
            foundCell.Offset(0, -1).Value = foundCell.Offset(0, -1).Value + Quantity
        End If
        
        ' Display the product name and quantity in the INV_TextBox
        Dim productName As String
        productName = GetProductName(productID)

        Dim historyArray(1 To 15, 1 To 2) As String
        Dim historyCounter As Integer
        ' Update the history of last 15 items scanned
        historyArray(historyCounter Mod 15 + 1, 2) = productName & " - Quantity: " & Quantity
        ' Increment the history counter
        historyCounter = historyCounter + 1
        
        ' Update the INV_TextBox with the last 10 items from the history
        Dim history As String
        Dim startIndex As Integer
        If historyCounter <= 10 Then
            startIndex = 1
        Else
            startIndex = (historyCounter Mod 15) - 9
        End If
        For i = startIndex To (startIndex + 9)
            history = history & historyArray(i Mod 15, 2) & vbCrLf
        Next i
        
        INV_TextBox.Text = history
        
        ' Clear the Scanned_IN_Textbox and Quantity_Textbox
        Scanned_IN_Textbox.Text = ""
        Quantity_Textbox.Text = ""
        
        ' Set the focus back to the Scanned_IN_Textbox for the next item
        Scanned_IN_Textbox.SetFocus
    End If
End Sub
Private Sub Scanned_IN_Textbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' Get the product ID from the Scanned_IN_Textbox
        Dim productID As String
        productID = Scanned_IN_Textbox.Text
        
        ' Check if the product ID exists in the Inventory table
        Dim foundCell As Range
        Set foundCell = Sheets("Inventory").Range("B2:B100").Find(productID, LookIn:=xlValues, LookAt:=xlWhole)
        
        If foundCell Is Nothing Then
            MsgBox "Product ID not found!"
        Else
        
            Dim productName As String
            productName = foundCell.Offset(0, 2).Value
            Quantity_Textbox.SetFocus
        End If
    End If
End Sub
Private Function GetProductName(ByVal productID As String) As String
    Dim productName As String
    Dim inventorySheet As Worksheet
    Set inventorySheet = ThisWorkbook.Sheets("Inventory")
    
    Dim productRange As Range
    Set productRange = inventorySheet.Range("B1:B100") ' Adjust the range as per your data
    
    Dim productCell As Range
    Set productCell = productRange.Find(productID)
    
    If Not productCell Is Nothing Then
        Dim productNameRange As Range
        Set productNameRange = inventorySheet.Range("D" & productCell.row)
        productName = productNameRange.Value
    Else
        productName = "Product Not Found"
    End If
    
    GetProductName = productName
End Function

