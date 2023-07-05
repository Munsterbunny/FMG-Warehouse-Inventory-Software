' Module: InventoryManagement' Maintain an array for keeping track of history for each technician
Public techHistory() As Variant

Function IsArrayInitialized(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayInitialized = IsArray(arr) And Not IsEmpty(LBound(arr))
    On Error GoTo 0
End Function

Function UpdateAndGetHistoryText(ByVal techName As String, ByVal productName As String, ByVal quantity As Integer) As String
    Dim j As Integer
    
    ' Check if techHistory is allocated, if not, initialize it with the first element
    If Not IsArrayInitialized(techHistory) Then
        ReDim techHistory(0 To 0)
        techHistory(0) = Array(productName, quantity)
    Else
        ' Otherwise, just add the new item to the array
        ReDim Preserve techHistory(LBound(techHistory) To UBound(techHistory) + 1)
        techHistory(UBound(techHistory)) = Array(productName, quantity)
    End If
    
    ' Only keep the last 10 items
    If UBound(techHistory) > 9 Then
        For j = LBound(techHistory) To UBound(techHistory) - 10
            techHistory(j) = Empty
        Next j
    End If
    
    ' Create the text to display in INVTextBox
    Dim displayText As String
    displayText = "History for " & techName & ":" & vbCrLf
    
    For j = LBound(techHistory) To UBound(techHistory)
        If Not IsEmpty(techHistory(j)) Then
            displayText = displayText & techHistory(j)(0) & " (Qty: " & techHistory(j)(1) & ")" & vbCrLf
        End If
    Next j
    
    UpdateAndGetHistoryText = displayText
End Function

' Function to get the Technician Name for a given Tech ID
Function GetTechnicianName(ByVal techID As String) As String
    Dim wsTech As Worksheet
    Set wsTech = ThisWorkbook.Sheets("Technicians")
    
    Dim lastRow As Long
    Dim i As Long
    lastRow = wsTech.Cells(wsTech.Rows.Count, "B").End(xlUp).row
    
    For i = 2 To lastRow
        If wsTech.Cells(i, 2).value = techID Then
            GetTechnicianName = wsTech.Cells(i, 1).value
            Exit Function
        End If
    Next i
    
    GetTechnicianName = "" ' Return empty if not found
End Function

' Function to get the Product Name for a given Product ID
Function GetProductName(ByVal productID As Integer) As String
    Dim wsInventory As Worksheet
    Set wsInventory = ThisWorkbook.Sheets("Inventory")
    
    Dim lastRow As Long
    Dim i As Long
    lastRow = wsInventory.Cells(wsInventory.Rows.Count, "B").End(xlUp).row
    
    For i = 2 To lastRow
        If wsInventory.Cells(i, 2).value = productID Then
            GetProductName = wsInventory.Cells(i, 4).value
            Exit Function
        End If
    Next i
    
    GetProductName = "" ' Return empty if not found
End Function

' Sub to deduct quantity from inventory
Sub DeductQuantityFromInventory(ByVal productID As Integer, ByVal quantity As Integer)
    Dim wsInventory As Worksheet
    Set wsInventory = ThisWorkbook.Sheets("Inventory")
    
    Dim lastRow As Long
    Dim i As Long
    lastRow = wsInventory.Cells(wsInventory.Rows.Count, "B").End(xlUp).row
    
    For i = 2 To lastRow
        If wsInventory.Cells(i, 2).value = productID Then
            wsInventory.Cells(i, 1).value = wsInventory.Cells(i, 1).value - quantity
            Exit Sub
        End If
    Next i
End Sub

Sub WriteScannedItemsToNewSheet(ByVal techName As String, ByVal productID As Integer, ByVal productName As String, ByVal quantity As Integer)
    Dim wsNew As Worksheet
    Dim sheetName As String
    sheetName = Format(Now, "mm-dd-yyyy")
    
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    Dim tbl As ListObject
    Dim tblExists As Boolean
    tblExists = False
    
    If wsNew Is Nothing Then
        Set wsNew = ThisWorkbook.Sheets.Add
        wsNew.Name = sheetName
        wsNew.Cells(1, 1).value = "Technician Name"
        wsNew.Cells(1, 2).value = "Product ID"
        wsNew.Cells(1, 3).value = "Product Name"
        wsNew.Cells(1, 4).value = "Quantity"
        Set tbl = wsNew.ListObjects.Add(xlSrcRange, wsNew.Range("A1:D1"), , xlYes)
        tbl.Name = "ScannedItemsTable"
    Else
        On Error Resume Next
        Set tbl = wsNew.ListObjects("ScannedItemsTable")
        On Error GoTo 0
        If Not tbl Is Nothing Then tblExists = True
    End If
    
    Dim nextRow As Long
    nextRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).row + 1
    
    wsNew.Cells(nextRow, 1).value = techName
    wsNew.Cells(nextRow, 2).value = productID
    wsNew.Cells(nextRow, 3).value = productName
    wsNew.Cells(nextRow, 4).value = quantity
    
    ' Dynamically expand table to fit new data
    If tblExists Then
        tbl.Resize wsNew.Range("A1:D" & nextRow)
    End If
End Sub

' Helper function to check if an array is allocated
Function IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And _
                       Not IsEmpty(arr) And _
                       LBound(arr, 1) <= UBound(arr, 1)
    On Error GoTo 0
End Function

Private Sub AssignTech(ByVal foundCell As Range, ByVal quantity As Integer)
    Dim techID As String
    Dim techName As String
    Dim targetSheet As Worksheet
    Dim currentDate As String
    Dim targetTable As ListObject
    
    ' Assuming techID and techName are in columns C and D respectively
    techID = foundCell.Offset(0, 1).value
    techName = foundCell.Offset(0, 2).value
    
    ' Get the current date
    currentDate = Format(Date, "yyyy-mm-dd")
    
    ' Check if sheet with today's date exists
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(currentDate)
    On Error GoTo 0
    
    ' If sheet doesn't exist, create a new sheet
    If targetSheet Is Nothing Then
        Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        targetSheet.Name = currentDate
        
        ' Create headers
        targetSheet.Cells(1, 1).value = "Technician ID"
        targetSheet.Cells(1, 2).value = "Technician Name"
        targetSheet.Cells(1, 3).value = "Quantity"
        
        ' Create table
        Set targetTable = targetSheet.ListObjects.Add(xlSrcRange, targetSheet.Range("A1:C1"), , xlYes)
        targetTable.Name = "ScannedItems"
        targetTable.TableStyle = "TableStyleMedium2"
    Else
        ' Get the existing table
        Set targetTable = targetSheet.ListObjects(1)
    End If
    
    ' Append the data to the table
    Dim newRow As ListRow
    Set newRow = targetTable.ListRows.Add
    newRow.Range(1, 1).value = techID
    newRow.Range(1, 2).value = techName
    newRow.Range(1, 3).value = quantity
End Sub
