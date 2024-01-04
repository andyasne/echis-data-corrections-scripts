    Sub FindTheRightLocationForaCase()
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
        For col = 1 To lastCol
            If ws.Cells(1, col).Value = "health_post" Then
                healthPostCol = col ' Store the index of the "health_post" column
            ElseIf ws.Cells(1, col).Value = "health_center" Then
                healthCenterCol = col ' Store the index of the "health_center" column
            End If
        Next col
    
        ' If both "health_post" and "health_center" columns are found, create "hchp" column
        If healthPostCol > 0 And healthCenterCol > 0 Then
            ' Insert a new column after the "health_center" column
            Dim max As Integer
            If healthCenterCol > healthPostCol Then
            max = healthCenterCol
            Else
            max = healthPostCol
            End If
            
            
            ws.Columns(max + 1).Insert Shift:=xlToRight
            ws.Cells(1, max + 1).Value = "hchp" ' Set header for the new column
    
            hchpCol = max + 1 ' Set index for the new "hchp" column
    
            ' Concatenate "health_center" and "health_post" for each row
            For row = 2 To lastRow
                Dim healthCenterValue As String
          
    
                healthCenterValue = Trim(ws.Cells(row, healthCenterCol).Value)
                healthPostValue = Trim(ws.Cells(row, healthPostCol).Value)
    
                ' Concatenate values and place in the new "hchp" column
                ws.Cells(row, hchpCol).Value = healthCenterValue & "" & healthPostValue
            Next row
        Else
            ' If either "health_post" or "health_center" is missing, skip creating the "hchp" column
            MsgBox "Either 'health_post' or 'health_center' column is missing. Skipping 'hchp' creation.", vbExclamation, "Column Not Found"
        End If
        
        
        '-------------------------------
    
    
     
        healthPostCol = 0 ' Initialize health_post column index
    
        ' Find the column index of "health_post"
        For col = 1 To lastCol
            If ws.Cells(1, col).Value = "health_post" Then
                healthPostCol = col ' Store the index of the "health_post" column
                Exit For
            End If
        Next col
    
        Dim owner_id As Long
    
        For col = 1 To lastCol
            If ws.Cells(1, col).Value = "owner_id" Then
                owner_id = col
                Exit For
            End If
        Next col
    
        If healthPostCol > 0 Then ' If "health_post" column was found
            Dim locationWorkbook As Workbook
               Dim locationFilePath As String
                Dim locationFilePath2 As String
                        Dim ls As Worksheet
    
                locationFilePath = ThisWorkbook.Path & "\locations.xlsx"
    
                locationFilePath2 = ThisWorkbook.Path & "\locations2.xlsx"
    
                    Set locationWorkbook = Workbooks.Open(locationFilePath)
    
            For row = 2 To lastRow
                 
                Dim ownerIdValue As String
                healthPostValue = Trim(ws.Cells(row, healthPostCol).Value) ' Get health_post value
                ownerIdValue = Trim(ws.Cells(row, owner_id).Value) ' Get owner_id value
    
                ' Get the location file path
    
                ' Check if the location file exists
                If Dir(locationFilePath) <> "" Then
    
    
                     Dim locationSheet As Worksheet
                    Dim locationCol As Long
                    locationCol = 0 ' Initialize location_id column index
    
                    ' Find the column index of "location_id" in the "health-post" sheet
                    For col = 1 To locationSheet.Cells(1, locationSheet.Columns.Count).End(xlToLeft).Column
                        If locationSheet.Cells(1, col).Value = "name" Then
                            locationCol = col ' Store the index of the "location_id" column
                            Exit For
                        End If
                    Next col
    
    
                    If locationCol > 0 Then ' If "location_id" column was found
                        Dim foundLocationId As Boolean
                        foundLocationId = False ' Flag to check if location_id was found
                         ' Loop through the "location_id" column in the "health-post" sheet
                         Dim foundOwner As Boolean
                        foundOwner = False
    
                        Dim locationRow As Long
                          locationRow = 2
    
    
                    Set locationSheet = locationWorkbook.Sheets("health-post")
    
    
                        For locationRow = 2 To locationSheet.Cells(locationSheet.Rows.Count, 1).End(xlUp).row
    
                        Dim v As String
                        v = locationSheet.Cells(locationRow, 1).Value
                   
                            If v = ownerIdValue Then
                                foundOwner = True ' Mark that owner_id was found
                                Exit For
                            End If
                        Next locationRow
    
    
                            Dim n As String
    
                            If foundOwner = False Then
                        ' Loop through the "location_id" column in the "health-post" sheet
                        locationRow = 2
        
        healthPostValue = Trim(ws.Cells(row, hchpCol).Value)
        
                    Dim locationWorkbookx As Workbook
    
    
                         Set locationWorkbookx = locationWorkbookx.Open(locationFilePath2)
    
    
                             Set ls = locationWorkbook.Sheets("health-postx")
    
    
                        For locationRow = 2 To ls.Cells(ls.Rows.Count, 1).End(xlUp).row
    
    
                        n = Trim(ls.Cells(locationRow, 3).Value)
    
                            If n = healthPostValue Then
                                ' If a matching health_post value is found in the location sheet
                                Dim currentLocationId As String
                                currentLocationId = Trim(ls.Cells(locationRow, 1).Value)  ' Get the current location_id
    
    
                                    ' Replace the location_id in the original sheet
                                    ws.Cells(row, owner_id).Value = currentLocationId
                                     ws.Cells(row, owner_id).Interior.Color = RGB(255, 255, 0)
                                    ' Output replaced values to the immediate window
                                    Debug.Print "Replaced: Row " & row & ", health_post: " & healthPostValue & ",Old location_name: " & ownerIdValue & " new location_id: " & currentLocationId
    
    
                                foundLocationId = True ' Mark that location_id was found
                                Exit For
                            End If
                        Next locationRow
                           End If
                        If Not foundLocationId Then
                           ' Debug.Print "No matching health_post found in location.xlsx for value: " & healthPostValue
                        End If
                    Else
                      '  Debug.Print "Column 'location_id' not found in 'health-post' sheet."
                    End If
    
                    ' Close the location workbook
                  '  locationWorkbook.Close SaveChanges:=False
                Else
                    Debug.Print "Location file 'location.xlsx' not found."
                End If
            Next row
        Else
            MsgBox "Column 'health_post' not found.", vbExclamation, "Column Not Found"
        End If



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




