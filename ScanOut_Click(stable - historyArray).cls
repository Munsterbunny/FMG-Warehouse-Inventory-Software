Private Sub ScanOut_Click()
MsgBox "I worked really hard for really long and with this new code you dont hit the button, you just press enter. I know, I know. Its not as fun. Look this thing is held together with prayers and typar tape. Well the tape ran out and I cant find the God of Lost Code in my heart. Just press the damn Enter Key.", , "o_o"
End Sub

Private Sub UserForm_Initialize()
    ' Set focus on techIDTextBox on UserForm Initialize
    techIDTextBox.SetFocus
End Sub

Private Sub techIDTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Move Focus to productIDTextBox on Enter Pressed in techIDTextBox
    If KeyCode = vbKeyReturn Then
        Dim techName As String
        techName = GetTechnicianName(techIDTextBox.value)
        If techName <> "" Then
            techIDTextBox.Enabled = False
            productIDTextBox.SetFocus
        Else
            MsgBox "Invalid Technician ID"
        End If
    End If
End Sub

Private Sub productIDTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Move Focus to quantityTextBox on Enter Pressed in productIDTextBox
    If KeyCode = vbKeyReturn Then
        Dim productName As String
        productName = GetProductName(CInt(productIDTextBox.value))
        If productName <> "" Then
            QuantityTextBox.SetFocus
        Else
            MsgBox "Invalid Product ID"
        End If
    End If
End Sub

Private Sub QuantityTextbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Update history array, inventory sheet, and reset input boxes on Enter Pressed in quantityTextBox
    If KeyCode = vbKeyReturn Then
        Dim techName As String
        techName = GetTechnicianName(techIDTextBox.value)
        Dim productName As String
        productName = GetProductName(CInt(productIDTextBox.value))
        Dim quantity As Integer
        quantity = CInt(QuantityTextBox.value)
        
        DeductQuantityFromInventory CInt(productIDTextBox.value), quantity
        WriteScannedItemsToNewSheet techName, CInt(productIDTextBox.value), productName, quantity
        Scanning_OUT.INVTextBox.Text = UpdateAndGetHistoryText(techName, productName, quantity)
        
        ' Reset productIDTextBox and quantityTextBox
        ClearForm
    End If
End Sub

Sub ClearForm()
    productIDTextBox.Text = ""
    QuantityTextBox.Text = ""
    productIDTextBox.SetFocus
End Sub

Private Sub ExitButton_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Save
    Erase techHistory
End Sub