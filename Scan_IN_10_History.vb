Dim historyCounter As Integer
Dim historyArray(1 To 15, 1 To 2) As String

Private Sub Scan_IN_Click()
    CheckScannedItem
End Sub
Private Sub Quantity_Textbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
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
    Dim productID As String
    productID = Scanned_IN_Textbox.Text
    
    Dim numericProductID As Integer
    numericProductID = Val(productID)
    
    If Not CheckMaxAmount(numericProductID) Then
        MsgBox "Maximum amount exceeded. Please enter a value less than or equal to " & MaxAmount & "." ,, "No way thats true."
        CheckScannedItem = False
        Exit Function
    End If
    
    If Quantity_Textbox.Text = "" And Scanned_IN_Textbox.Text = "" Then
        MsgBox "Please Scan Item", , "FMG"
        clearForm
    ElseIf Quantity_Textbox.Text > "0" And Scanned_IN_Textbox.Text > "0" Then
        ProcessScannedItem
    ElseIf Quantity_Textbox.Text <= "0" And Scanned_IN_Textbox.Text <> "0" Then
        MsgBox "Do I need to explain this one?", , "You are bad at math"
        clearForm
    ElseIf Quantity_Textbox.Text <> "" And Scanned_IN_Textbox.Text = "" Then
        MsgBox "Make sure you scan the barcode first!", , "Oops"
        clearForm
        If Not IsNumeric(Quantity_Textbox.Text) Or Not IsValidProductID(Scanned_IN_Textbox.Text) Then
            MsgBox "Invalid input. Please enter a valid numeric quantity and product ID."
            clearForm
            Exit Function
        End If
    End If
    
    ' Check if the product ID exists in the Inventory table
    Dim foundCell As Range
    Set foundCell = Sheets("Inventory").Range("B2:B100").Find(productID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        MsgBox "Product ID not found!"
        clearForm
        CheckScannedItem = False
    Else
        Dim productName As String
        productName = GetProductName(productID)
        Scanned_IN_Textbox.SetFocus
        CheckScannedItem = True
    End If
End Function

Private Function GetProductName(ByVal productID As String) As String
    Dim productName As String
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
        Debug.Print productName
    Else
        MsgBox "These are not the droids you are looking for", , "Unknown Product" 'you got wayyyy too cocky here
        clearForm
    End If
    
    GetProductName = productName
End Function

Private Function ProcessScannedItem() As Boolean
    If Not IsNumeric(Quantity_Textbox.Text) Or Not IsValidProductID(Scanned_IN_Textbox.Text) Then
        MsgBox "Invalid input. Please enter a valid numeric quantity and product ID."
        clearForm
        ProcessScannedItem = False
        Exit Function
    End If
    Dim Quantity As Integer
    Quantity = Val(Quantity_Textbox.Text)
    If Not CheckMaxAmount(Quantity) Then
        KeyCode = 0
        clearForm
    End If
    ' Update the Inventory table with the quantity
    Dim productID As String
    productID = Scanned_IN_Textbox.Text
    
    ' Find the corresponding row in the Inventory table
    Dim foundCell As Range
    Set foundCell = Sheets("Inventory").Range("B2:B100").Find(productID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        ' Update the quantity in the Quantity column
        foundCell.Offset(0, -1).value = foundCell.Offset(0, -1).value + Quantity
    End If
    
    ' Display the product name and quantity in the INV_TextBox
    Dim productName As String
    productName = GetProductName(productID)
    ' Update the history of last 15 items scanned
    historyArray(historyCounter Mod 15 + 1, 2) = productName & " - Quantity: " & Quantity
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
    
    INV_TextBox.Text = history
    ProcessScannedItem = True
    clearForm
End Function
    
    Private Function IsValidProductID(ByVal productID As String) As Boolean
        Dim validCharacters As String
        validCharacters = "0123456789*"
        
        Dim i As Integer
        For i = 1 To Len(productID)
            If InStr(validCharacters, Mid(productID, i, 1)) = 0 Then
                IsValidProductID = False
                Exit Function
            End If
        Next i
        
        ' Check if the product ID exists in the Inventory table
        Dim foundCell As Range
        Set foundCell = Sheets("Inventory").Range("B2:B100").Find(productID, LookIn:=xlValues, LookAt:=xlWhole)
        
        IsValidProductID = Not foundCell Is Nothing
    End Function
  
Private Sub clearForm()
    ' Clear the Scanned_IN_Textbox and Quantity_Textbox
    Scanned_IN_Textbox.Text = ""
    Quantity_Textbox.Text = ""
    Scanned_IN_Textbox.SetFocus
End Sub

Private Sub Exit_Button_Click()
    Unload Me
 End Sub

'HELL YEAH I did it! I dont care that it took ChatGPT, StackOverFlow, 20mg Adderall, using a phone as a hotspot and 40 hours of tearing my hair out I DID THIS. (12:24pm) 6/14/2023
'you werent done dickhead, and now you jinxed us,TODO you fixed all the edge cases but fucked up the array(2:30pm) 6/14/2023
'yeah the array wasnt it you also messed up the actual inventory system...(2:36pm) 6/14/2023
'...Nobody...Touch...Anything....(6:12pm) 6/14/2023
'fuck, anywhoo..(8:02pm) 6/14/2023
'TODO EC Putting in Quantity above 200 and Scan in above 200 doesnt clearForm
'TODO EC if Scan in is correct but Quantity is above 200 NO ProductName appears in the array table
'TODO EC If both are set to letters 2 different error windows show up
'TODO EC if Scan in is correct but you hit the ScanIn button instead of hitting enter you get insulted, change to numbers <=0 you get insulted otherwise act as an enter key
'TODO EC putting zeros before the numbers causes the application to think of them as letters, no idea how to start that fix
'TODO Get cocky again and fuck things up... ya know... for funsies
