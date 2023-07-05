Dim productID As Integer
Dim quantity As Integer
Dim historyCounter As Integer
Dim historyArray(1 To 15, 1 To 2) As String
    
Private Sub ScanIN_Click()
    CheckScannedItem
End Sub
Private Sub QuantityTextbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
    CheckScannedItem
    End If
End Sub
Private Function CheckMaxAmount(ByVal value As Integer) As Boolean
    Const MaxAmount As Integer = 200
        
    If value > MaxAmount Then
        CheckMaxAmount = False
    Else
        CheckMaxAmount = True
    End If
End Function
Private Function CheckScannedItem() As Boolean

    Const MaxAmount As Integer = 200 ' Maximum quantity and product ID
    productID = Val(Trim(productIDTextBox.Text))
    quantity = Val(Trim(QuantityTextBox.Text))
    
    If Not IsNumeric(Trim(QuantityTextBox.Text)) Or Not IsNumeric(Trim(productIDTextBox.Text)) Or productIDTextBox.Text = "" Or productIDTextBox.Text = "" Then
    MsgBox "Please enter a valid numeric quantity and product ID.", , "Invalid Input"
    ClearForm
    Exit Function
    End If

    Select Case True
        Case quantity <= 0 Or quantity > MaxAmount
            MsgBox "Invalid quantity. Please enter a value between 1 and " & MaxAmount & ".", , "Invalid Quantity"
            ClearForm
            CheckScannedItem = False
            Exit Function
            
        Case productID <= 0 Or productID > MaxAmount
            MsgBox "Invalid product ID. Please enter a value between 1 and " & MaxAmount & ".", , "Invalid Product ID"
            ClearForm
            CheckScannedItem = False
            Exit Function
            
        Case quantity = 0 And productID = 0
            MsgBox "Please scan an item and then hit the enter key OR you may use the button", , "FMG"
            ClearForm
            
        Case quantity > 0 And productID > 0
            ProcessScannedItem
            
        Case quantity <= 0 And productID <> 0
            MsgBox "So we are just trying to break me, is that it?", , "Just kinda rude at this point"
            ClearForm
            
        Case productID <> 0 And quantity = 0
            MsgBox "Make sure you scan the barcode first!", , "Oops"
            ClearForm
        
        Case Else
            ' Check if the product ID exists in the Inventory table
            Dim foundCell As Range
            Set foundCell = Sheets("Inventory").Range("B2:B100").Find(productID, LookIn:=xlValues, LookAt:=xlWhole)
            
            If foundCell Is Nothing Then
                MsgBox "Product ID not found!", , "Please check the number and dial again."
                ClearForm
                CheckScannedItem = False
            Else
                Dim productName As String
                productName = GetProductName(productID)
                CheckScannedItem = True
                ClearForm
            End If
    End Select
End Function

Private Function GetProductName(ByVal productID As String) As String
    Dim productName As String
    productID = Val(Trim(productIDTextBox.Text))
    Dim inventorySheet As Worksheet
    Set inventorySheet = ThisWorkbook.Sheets("Inventory")
    
    Dim productRange As Range
    Set productRange = inventorySheet.Range("B1:B100")
    
    Dim productCell As Range
    Set productCell = productRange.Find(productID)
    
    If Not productCell Is Nothing Then
        Dim productNameRange As Range
        Set productNameRange = inventorySheet.Range("D" & productCell.row)
        productName = productNameRange.value
    Else
    End If
    
    GetProductName = productName
End Function

Private Function ProcessScannedItem() As Boolean
        
    productID = Val(Trim(productIDTextBox.Text))
    quantity = Val(Trim(QuantityTextBox.Text))
 
    ' Update the Inventory table with the quantity
    
    ' Find the corresponding row in the Inventory table
    Dim foundCell As Range
    Set foundCell = Sheets("Inventory").Range("B2:B100").Find(productID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        ' Update the quantity in the Quantity column
        foundCell.Offset(0, -1).value = foundCell.Offset(0, -1).value + quantity
    Else
        MsgBox "We aint found shit!", , "What about you guys?"
        ClearForm
        ProcessScannedItem = False
        Exit Function
    End If
    
    ' This part below starts an array so that the history diplays the last 10
    Dim productName As String
    productName = GetProductName(productID)
    historyArray(historyCounter Mod 15 + 1, 2) = productName & " - Quantity: " & quantity ' Update the history of last 15 items scanned
    historyCounter = historyCounter + 1
    
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
    
    INVTextBox.Text = history
    ClearForm
    ProcessScannedItem = True
    End Function
    
    Sub ClearForm()
    productIDTextBox.Text = ""
    QuantityTextBox.Text = ""
    productIDTextBox.SetFocus
End Sub
    
Private Sub ExitButton_Click()
    Unload Me
 End Sub

