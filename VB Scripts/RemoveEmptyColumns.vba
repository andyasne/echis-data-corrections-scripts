    Sub RemoveEmptyColumns()
        ' Call the method to update the "closed" column
        Call UpdateClosedColumn
        On Error Resume Next
        Dim ws As Worksheet
        Dim lastRow As Long, lastCol As Long, col As Long, row As Long
        Dim emptyCol As Boolean
        
            Dim deletedCount As Integer ' Variable to count the deleted columns
        ' Set the active sheet
        Set ws = ActiveSheet
        ' Find the last row and last column with data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        '----------------------
        
        Dim healthPostCol As Long
        Dim healthCenterCol As Long
        Dim hchpCol As Long
    
        healthPostCol = 0 ' Initialize health_post column index
        healthCenterCol = 0 ' Initialize health_center column index
        hchpCol = 0 ' Initialize hchp column index
      ' Find the column index of "health_post" and "health_center"
        
        
        
        '-------------------------------
    
     
    
        ' Replace "---" with an empty string in the entire worksheet
        ws.Cells.Replace What:="---", Replacement:="", LookAt:=xlPart, MatchCase:=False
    
        ' Initialize count of deleted columns
        deletedCount = 0
        ' Loop through each column
        For col = lastCol To 1 Step -1
            emptyCol = True ' Assume the column is empty initially
            ' Check each row in the column starting from the second row
            For row = 2 To lastRow ' Start from row 2 to skip the first row
                If Trim(ws.Cells(row, col).Value) <> "" Then
                    emptyCol = False ' If any cell in the column is not empty, mark as not empty
                    Exit For ' No need to check further in this column
                End If
            Next row
            ' If the column is entirely empty or matches the specified names, delete it
            If emptyCol Or ws.Cells(1, col).Value = "case_link" Or ws.Cells(1, col).Value = "number" Or ws.Cells(1, col).Value = "closed_date" Or ws.Cells(1, col).Value = "closed_by_username" Or ws.Cells(1, col).Value = "case_type" Or ws.Cells(1, col).Value = "closed_by_username" Or ws.Cells(1, col).Value = "owner_name" Then ws.Columns(col).Delete
                deletedCount = deletedCount + 1 ' Increment count of deleted columns
            End If
        Next col
        ' Show the number of columns deleted
        MsgBox deletedCount & " columns were deleted.", vbInformation, "Columns Deleted"
      


        ' Save the modified workbook
        Dim filePath As String
        Dim fileName As String
        Dim folderPath As String
    
        ' Get the current file's path and name
        filePath = ThisWorkbook.FullName
        fileName = ThisWorkbook.Name
    
        ' Create a folder named "importe" in the same directory if it doesn't exist
        folderPath = ThisWorkbook.Path & "\imported"
        If Dir(folderPath, vbDirectory) = "" Then
            MkDir folderPath
        End If
    
        ' Construct the new file path with the prefix and folder name
        Dim newFilePath As String
        newFilePath = folderPath & "\" & "Recovered Data - " & fileName
    
        ' Save the modified workbook with the prefix in the "importe" folder
       ThisWorkbook.SaveAs newFilePath
    
        MsgBox "File saved as 'Recovered Data " & fileName & "' in 'imported' folder.", vbInformation, "File Saved"
    
    
    
    End Sub
    
    
    
    
    
    
    Sub UpdateClosedColumn()
        Dim ws As Worksheet
        Dim lastRow As Long, lastCol As Long, col As Long, row As Long
        Dim closedColIndex As Long
        ' Set the active sheet
        Set ws = ActiveSheet
        ' Find the last row and last column with data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
        ' Find the "closed" column and rename it to "close"
        For col = 1 To lastCol
            If ws.Cells(1, col).Value = "closed" Then
                ws.Cells(1, col).Value = "close"
                closedColIndex = col ' Store the index of the "close" column
                Exit For
            End If
        Next col
    
        ' Loop through the "close" column and replace values
        If closedColIndex > 0 Then ' If "close" column was found
            For row = 2 To lastRow
                If UCase(Trim(ws.Cells(row, closedColIndex).Value)) = "TRUE" Then
                    ws.Cells(row, closedColIndex).Value = "Yes"
                Else
                    ws.Cells(row, closedColIndex).Value = "No"
                End If
            Next row
    ' Change background color of the "close" column to red
    ws.Columns(closedColIndex).Interior.Color = RGB(255, 0, 0) ' Red color
        Else
            MsgBox "Column 'closed' not found.", vbExclamation, "Column Not Found"
        End If
    End Sub





