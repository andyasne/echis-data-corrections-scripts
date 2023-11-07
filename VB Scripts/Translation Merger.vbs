Dim sourceSheet As Worksheet
Dim targetSheet As Worksheet
Dim lastRowSource As Long
Dim lastRowTarget As Long
Dim engColumnSource As Range
Dim engColumnTarget As Range
Dim somColumnSource As Range
Dim somColumnTarget As Range
Dim engValue As Variant
Dim colIndexSource As Long
Dim colIndexTarget As Long
Dim colIndexSom As Long
Dim colIndexEng As Long
Dim workbookB As Workbook
Dim reportWorkbook As Workbook
Dim reportWorksheet As Worksheet
Dim reportRow As Long
Dim translatedFilePath As String
Dim somHeader As String
Sub UpdateMissingLabelsFromTranslation()


' Fetch the folder path of the current Excel document
Dim currentFolder As String
currentFolder = ThisWorkbook.Path

' Construct the file path for Workbook B ("translatedFilePath")
translatedFilePath = currentFolder & "\Translated.xlsx"

' Check if Workbook B exists in the folder
If Dir(translatedFilePath) = "" Then
    MsgBox "Translated.xlsx not found in the same folder as the current workbook."
    Exit Sub
End If

' Ask the user for the column header for "default_som"
somHeader = InputBox("Enter the header. such as : 'default_som'")

' Ensure the user provided a header, or exit if canceled
If somHeader = "" Then
    Exit Sub
End If

' Open Workbook B
Set workbookB = Workbooks.Open(translatedFilePath)

' Create a new workbook for the report
Set reportWorkbook = Workbooks.Add
Set reportWorksheet = reportWorkbook.Sheets(1)

' Set the report headers

reportWorksheet.Cells(1, 1).Value = "Default_Eng Value"
reportWorksheet.Cells(1, 2).Value = somHeader & " Previous Value"
reportWorksheet.Cells(1, 3).Value = somHeader & " New Value"

' Initialize the report row
reportRow = 2
' Loop through each sheet in Workbook B
For Each targetSheet In workbookB.Sheets
Debug.Print "Processing target sheet: " & targetSheet.Name

' Check if the sheet exists in Workbook A
On Error Resume Next
Set sourceSheet = ThisWorkbook.Sheets(targetSheet.Name)
On Error GoTo 0

Debug.Print "Source sheet found: " & IIf(Not sourceSheet Is Nothing, "Yes", "No")

' Only proceed if the sheet names match
If Not sourceSheet Is Nothing Then
    ' Find the "default_eng" column in the target sheet
    On Error Resume Next
    Set engColumnTarget = targetSheet.UsedRange.Rows(1).Find("default_en", LookIn:=xlValues)
    Set engColumnSource = sourceSheet.UsedRange.Rows(1).Find("default_en", LookIn:=xlValues)
    colIndexEng = engColumnSource.Column
    targetSheet.Cells.EntireColumn.Hidden = False
    Set somColumnTarget = targetSheet.UsedRange.Rows(1).Find(somHeader, LookIn:=xlValues)
    colIndexSom = somColumnTarget.Column
    On Error GoTo 0
    
    Debug.Print "Processing sheet: " & targetSheet.Name
    Debug.Print "Eng Column found: " & IIf(Not engColumnTarget Is Nothing, "Yes", "No")
    
    If Not engColumnTarget Is Nothing Then
        colIndexTarget = engColumnTarget.Column
        lastRowTarget = targetSheet.Cells(targetSheet.Rows.Count, colIndexTarget).End(xlUp).Row
        
        ' Find the "default_som" column in the source sheet
        On Error Resume Next
        Set somColumnSource = sourceSheet.UsedRange.Rows(1).Find(somHeader, LookIn:=xlValues)
        On Error GoTo 0
        
        Debug.Print "Som Column found: " & IIf(Not somColumnSource Is Nothing, "Yes", "No")
        
        If Not somColumnSource Is Nothing Then
            colIndexSource = somColumnSource.Column
                   On Error Resume Next
            lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, colIndexSource).End(xlUp).Row
             On Error GoTo 0
            ' Initialize a flag to track if any cell meets the condition
            Dim changeFlag As Boolean
            changeFlag = False
            
            ' Loop through each row in the target sheet
            For i = 2 To lastRowTarget ' Starting from row 2 (assuming row 1 is headers)
              On Error Resume Next
                engValue = targetSheet.Cells(i, colIndexTarget).Value
                  On Error GoTo 0
                ' Search for the value in the source sheet and copy if found
                If Not IsEmpty(engValue) Then
                    Dim findRange As Range
                      On Error Resume Next
                    Set findRange = sourceSheet.Columns(colIndexEng).Find(engValue, LookIn:=xlValues, LookAt:=xlWhole)
                    On Error GoTo 0
                    
                    If Not findRange Is Nothing Then
                    If sourceSheet.Cells(findRange.Row, colIndexSource).Value <> targetSheet.Cells(i, colIndexSom).Value Then
                        ' Copy the "default_som" value from target to source
                        sourceSheet.Cells(findRange.Row, colIndexSource).Value = targetSheet.Cells(i, colIndexSom).Value
                        
                        ' Change the font color to red for the modified cell
                        sourceSheet.Cells(findRange.Row, colIndexSource).Font.Color = RGB(255, 0, 0)
                        sourceSheet.Tab.Color = RGB(255, 0, 0) ' Yellow color
                        
                        If InStr(sourceSheet.Cells(findRange.Row, colIndexSource).Value, "<") > 0 Or InStr(sourceSheet.Cells(findRange.Row, colIndexSource).Value, ">") > 0 Then
                            ' Change the background color of the cell to yellow
                            sourceSheet.Cells(findRange.Row, colIndexSource).Interior.Color = RGB(255, 255, 0) ' Yellow color
                            
                            ' Set the flag to indicate that a cell met the condition
                            changeFlag = True
                        End If
                        
                        ' Record the change in the report
              
                        reportWorksheet.Cells(reportRow, 1).Value = engValue
                        reportWorksheet.Cells(reportRow, 2).Value = sourceSheet.Cells(findRange.Row, colIndexSource).Value
                        reportWorksheet.Cells(reportRow, 3).Value = targetSheet.Cells(i, colIndexSom).Value
                        reportRow = reportRow + 1
                    End If
                       End If
                End If
            Next i
            
            ' Check if any cell met the condition and then change the tab color
            If changeFlag Then
                sourceSheet.Tab.Color = RGB(255, 255, 0) ' Yellow color
            End If
        End If
    End If
End If
Next targetSheet

' Save the report workbook in the same folder as the current workbook
Dim reportFilePath As String
reportFilePath = currentFolder & "\CompleteTranslated-" & somHeader & ".xlsx"

' Save and overwrite if it already exists
reportWorkbook.SaveAs reportFilePath
reportWorkbook.Close SaveChanges:=True

' Close Workbook B
workbookB.Close SaveChanges:=True


If MsgBox("Do you want to update all the missing untranslated Labels from the complete Translated list?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then
  Exit Sub ' Exit the subroutine if the user chooses not to run the code
 End If

Call UpdateMissingTranslationsFromCompleteList

End Sub


Sub UpdateMissingTranslationsFromCompleteList()

currentFolder = ThisWorkbook.Path
somHeader = InputBox("Enter the header. such as : 'default_som'")
'   somHeader = "default_aar"

      Dim reportFilePath As String
reportFilePath = currentFolder & "\CompleteTranslated-" & somHeader & ".xlsx"
    If Dir(reportFilePath) = "" Then
    MsgBox "CompleteTranslated-" & somHeader & ".xlsx" & " is not found in the same folder as the current workbook."
    Exit Sub
End If
Set reportWorkbook = Workbooks.Open(reportFilePath)
Dim ws As Worksheet
On Error Resume Next ' Enable error handling
For Each ws In ThisWorkbook.Sheets
On Error GoTo 0 ' Reset error handling
' Loop through each row in the sheet
Dim rowNum As Long
On Error Resume Next ' Enable error handling
For rowNum = 2 To ws.UsedRange.Rows.Count ' Assuming headers are in row 1
    On Error GoTo 0 ' Reset error handling
    
    Set engColumnSource = ws.UsedRange.Rows(1).Find("default_en", LookIn:=xlValues)
    If Not engColumnSource Is Nothing Then
        colIndexEng = engColumnSource.Column
        Dim defaultEngValue As Variant
        defaultEngValue = ws.Cells(rowNum, colIndexEng).Value ' Assuming colIndexEng is set earlier
        Debug.Print "Processing Worksheet: " & ws.Name & ", Row: " & rowNum

        Set somColumnTarget = ws.UsedRange.Rows(1).Find(somHeader, LookIn:=xlValues)
        If Not somColumnTarget Is Nothing Then
            colIndexSom = somColumnTarget.Column

            If Not IsEmpty(ws.Cells(rowNum, colIndexSom).Value) Then
                Debug.Print "Searching for default_eng value: " & defaultEngValue
                ' Search for the default_eng value in the reportWorkbook's first sheet
                Dim reportSheet As Worksheet
                Set reportSheet = reportWorkbook.Sheets(1) ' Assuming report data is in the first sheet
                Dim reportRange As Range
                Set reportRange = reportSheet.UsedRange
                If Not IsEmpty(defaultEngValue) And defaultEngValue <> "" Then
                ' Find the default_eng value in the 3rd column
                Dim foundCell As Range
                On Error Resume Next
                Set foundCell = reportRange.Find(What:=CStr(defaultEngValue), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                On Error GoTo 0 ' Reset error handling
                
                If Not foundCell Is Nothing Then
                    Debug.Print "Found default_eng value in reportSheet: " & foundCell.Value
                    Dim foundValue As Variant
                    foundValue = foundCell.Offset(0, 2).Value

                    If Not IsEmpty(foundValue) & foundValue <> ws.Cells(rowNum, colIndexSom).Value Then
                        Debug.Print "Copying value to somHeader column: " & foundValue
                        ' Copy the 5th column value from the reportSheet to the current row's somHeader column
                        ws.Cells(rowNum, colIndexSom).Value = foundValue
                        ws.Cells(rowNum, colIndexSom).Font.Color = RGB(0, 0, 255)
                        ws.Tab.Color = RGB(0, 0, 255)
                        ' Set the font color to blue (or any color you prefer)
                    End If
                End If
                    End If
            End If
        End If
    End If
Next rowNum
Next ws

End Sub






