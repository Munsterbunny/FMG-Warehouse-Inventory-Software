Sub OpenFMGForm_Click()
    FMG.Show
End Sub

Sub ExportToWord()
    Dim objWord As Object
    Dim objDoc As Object
    Dim tbl As Object
    Dim row As Long, col As Long
    
    ' Create a new instance of Word
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    
    ' Create a new Word document
    Set objDoc = objWord.Documents.Add
    
    ' Add a table to the document
    Set tbl = objDoc.Tables.Add(Range:=objDoc.Range,NumRows:=Range("A2:A100").Cells.SpecialCells(xlCellTypeConstants).Count + 1, NumColumns:=3)
    
    ' Set table headers
    tbl.cell(1, 1).Range.Text = "Item"
    tbl.cell(1, 2).Range.Text = "Quantity"
    tbl.cell(1, 3).Range.Text = "ID #"
    
    ' Populate the table with inventory data
    row = 2 ' Start from row 2 in the table
    For Each cell In Range("D2:D100").Cells ' Assuming items are in column D
        tbl.cell(row, 1).Range.Text = cell.Value
        tbl.cell(row, 2).Range.Text = cell.Offset(0, -3).Value
        tbl.cell(row, 3).Range.Text = cell.Offset(0, -1).Value
        row = row + 1
    Next cell
    
    ' Generate the filename with date
    Filename = "FMG Monthly Report " & Format(Now(), "YYYY-MM-DD") & ".docx"
    
    ' Save the document with the generated filename
    objDoc.SaveAs "D:" & Filename ' Update the path as needed
    
    ' Close Word
    objWord.Quit
    
    ' Clean up the objects
    Set tbl = Nothing
    Set objDoc = Nothing
    Set objWord = Nothing
End Sub

