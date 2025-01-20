Sub BuildReport()
    ' Declare variables
    Dim idRange As Range, titleRange As Range, descRange As Range
    Dim reportType As Integer, cell As Range
     Dim currentID As String
     Dim serialNumber As Integer

    'Prompt user to select ID, Title, and Description ranges
    On Error Resume Next
    Set idRange = Application.InputBox("Select the ID range", Type:=8)
    Set titleRange = Application.InputBox("Select the Title range", Type:=8)
    Set descRange = Application.InputBox("Select the Description range", Type:=8)
    On Error GoTo 0
Do
    'Check if ranges are selected
    If idRange Is Nothing Or titleRange Is Nothing Or descRange Is Nothing Then
        MsgBox "Please select all ID, Title, and Description ranges."
        Exit Sub
    End If

    'Prompt user to choose report type by number
    reportType = InputBox("Choose report type: 1 for datasource, 2 for web report, 3 for mobile report")
Dim outres As String
 


    'Check the selected report type
    Select Case reportType
        Case 1
            For Each cell In idRange
                
                currentID = Replace(Trim(cell.Value), ".", "_")
                If descRange.Cells(cell.Row - idRange.Row + 1, 1).Value <> "" Then
                   outres = outres + "{" & """type"": ""expression""," & """column_id"": """ & currentID & """," & """datatype"": ""integer""," & """display_name"": """ & titleRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """" & ","
                  outres = outres + """expression"": {""" & "datatype"": ""integer""," & """test"": {""" & "type"": ""and""," & """filters"": [{" & """operator"": ""eq""," & """type"": ""boolean_expression""," & """expression"": {""" & "datatype"": ""string""," & """type"": ""property_name""," & """property_name"": ""indicator_name""" & "}," & """comment"": null," & """property_value"": ""Xindicator_name""" & "},{" & """operator"": ""eq""," & """expression"": {""" & "type"": ""root_doc""," & """expression"": {""" & "datatype"": ""string""," & """type"": ""property_name""," & """property_name"": ""property_name""" & "}" & "}," & """type"": ""boolean_expression""," & """comment"": null," & """property_value"": ""xproperty_name""" & "}]" & "}," & """type"": ""conditional""," & """expression_if_true"": {""" & "type"": ""constant""," & """constant"": 1" & "}," & """expression_if_false"": {""" & "type"": ""constant""," & """constant"": 0" & "}" & "}," & """is_nullable"": true," & """is_primary_key"": false," & ""
                    outres = outres + """create_index"": false," & """transform"": {}," & """comment"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & " """ & "}" & ","
                
                End If
                
              
            Next cell
              CopyText (outres)
        Case 3
          serialNumber = 0
         For Each cell In idRange
        
                currentID = Replace(Trim(cell.Value), ".", "_")
                If descRange.Cells(cell.Row - idRange.Row + 1, 1).Value <> "" Then
                 serialNumber = serialNumber + 1
             outres = outres + "{" & """type"": ""detail-screen-config:Column""," & """contents"": {" & """hasAutocomplete"":false," & """useXpathExpression"":true," & """calc_xpath"":"".""," & """enum"":[],"
   outres = outres + """field"":""column[@id = '" & currentID & "']""," & """filter_xpath"":""""," & """format"":""plain""," & """graph_configuration"":null," & """header"":{""en"":""" & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """}," & """model"":""case""," & """date_format"":"""","
 outres = outres + """time_ago_interval"":365.25," & """horizontal_align"":""left""," & """vertical_align"":""start""," & """font_size"":""medium""," & """show_border"":false," & """show_shading"":false," & """late_flag"":50," & """case_tile_field"":""""," & """isTab"":false," & """hasNodeset"":false," & """nodeset"":"""","
 outres = outres + """nodesetCaseType"":""""," & """nodesetFilter"":""""," & """relevant"":""""," & """endpoint_action_id"":null," & """grid_x"":0," & """grid_y"":0," & """height"":1," & """width"":6}" & "}   " & vbCrLf


                End If
            Next cell
              CopyText (outres)
        Case 2
        Dim exp As String
        
        
              For Each cell In idRange
            
                  currentID = Replace(Trim(cell.Value), ".", "_")
                    If Right(currentID, 4) = "=sum" Then
                         exp = MethodX(cell, titleRange.Cells(cell.Row - idRange.Row + 1, 1), currentID, titleRange.Cells(cell.Row - idRange.Row + 1, 1).Value, descRange.Cells(cell.Row - idRange.Row + 1, 1).Value)
                                     
  outres = outres + "{" & """comment"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """," & """description"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """," & """format"": ""default""," & """css_class"": null," & """width"": null," & """column_id"": """ & currentID & """," & """visible"": true," & """transform"": {}," & """expression"": " & exp & "}," & """type"": ""expression""," & """display"":  """ & titleRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """}"
  outres = outres + "" & ","
                    Else
                        ' Current computation
                        If descRange.Cells(cell.Row - idRange.Row + 1, 1).Value <> "" Then
                            outres = outres + "{" & """comment"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """," & """field"": """ & LCase(currentID) & """," & """description"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """," & """format"": ""default""," & """css_class"": null," & """width"": null," & """aggregation"": ""sum""," & """column_id"": """ & currentID & """," & """visible"": true," & """transform"": {}," & """calculate_total"": true," & """type"": ""field""," & """display"":  """ & titleRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """}"
                              outres = outres + "" & ","
                        End If
                    End If
            Next cell
            
            
            
            
              CopyText (outres)
        Case Else
            MsgBox "Invalid report type selected."
    End Select
         If MsgBox("Do you want to run the macro again?", vbYesNo + vbQuestion) = vbNo Then Exit Do
    Loop While True
End Sub
Sub CopyText(Text As String)
  Debug.Print Text
  
  
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

 Function MethodX(cell As Range, titleCell As Range, currentID As String, title As String, description As String) As String
    Dim concatResults(1 To 50) As String  ' Array to hold up to 10 results
    Dim startNum As String
    Dim currentCell As Range
    Dim rowNum As Long
    Dim i As Integer
    Dim values() As String  ' Array to hold split values

    startNum = Trim(Split(title, "-")(0)) & "."
    rowNum = titleCell.Row  ' Initialize rowNum to the row of the provided cell
    i = 1  ' Start with the first index of the results

    ' Iterate over the rows below the current cell to find child cells
    Do
        rowNum = rowNum + 1
        Set currentCell = cell.Worksheet.Cells(rowNum, titleCell.Column)
        
        ' Check if the cell starts with the same number and a dot, e.g., "2.", "2.1", "2.2", etc.
        If Left(currentCell.Value, Len(startNum)) = startNum Then
            ' Split the cell value into individual entries
          concatResults(i) = currentCell.Offset(0, -3).Value
           i = i + 1
            
        Else
            ' Exit the loop if the cell value doesn't follow the child pattern
            Exit Do
        End If
    Loop Until IsEmpty(currentCell.Value) Or i > 50 Or currentCell.Value Like "#*" And Not currentCell.Value Like startNum & "#*"

    ' Format the output as a JSON-like string for expression
    Dim concatResult As String
    Dim statement As String

    ' Construct the context variables and the statement dynamically
    concatResult = """context_variables"": {"
    For j = 1 To i
        If concatResults(j) <> "" Then
            If j > 1 Then
                concatResult = concatResult & ","
                statement = statement & " + "
            End If
            concatResult = concatResult & vbCrLf & "        ""a" & j & """: {""type"": ""property_name"", ""property_name"": """ & concatResults(j) & """}"
            statement = statement & "a" & j
        End If
    Next j
    concatResult = concatResult & vbCrLf & "      }"

    ' Construct the full expression JSON
    Dim expression As String
    expression = "{""type"": ""evaluator"", ""statement"": """ & statement & """, " & concatResult '& "}"

    ' Return the formatted string
    MethodX = expression
End Function





