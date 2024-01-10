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

    'Check if ranges are selected
    If idRange Is Nothing Or titleRange Is Nothing Or descRange Is Nothing Then
        MsgBox "Please select all ID, Title, and Description ranges."
        Exit Sub
    End If

    'Prompt user to choose report type by number
    reportType = InputBox("Choose report type: 1 for datasource, 2 for mobile report, 3 for web report")

    'Check the selected report type
    Select Case reportType
        Case 1
            For Each cell In idRange
                
                currentID = Replace(Trim(cell.Value), ".", "_")
                If descRange.Cells(cell.Row - idRange.Row + 1, 1).Value <> "" Then
                    Debug.Print "{" & """type"": ""expression""," & """column_id"": """ & currentID & """," & """datatype"": ""integer""," & """display_name"": """ & titleRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """" & ","
                    Debug.Print """expression"": {""" & "datatype"": ""integer""," & """test"": {""" & "type"": ""and""," & """filters"": [{" & """operator"": ""eq""," & """type"": ""boolean_expression""," & """expression"": {""" & "datatype"": ""string""," & """type"": ""property_name""," & """property_name"": ""indicator_name""" & "}," & """comment"": null," & """property_value"": ""Xindicator_name""" & "},{" & """operator"": ""eq""," & """expression"": {""" & "type"": ""root_doc""," & """expression"": {""" & "datatype"": ""string""," & """type"": ""property_name""," & """property_name"": ""property_name""" & "}" & "}," & """type"": ""boolean_expression""," & """comment"": null," & """property_value"": ""xproperty_name """ & "}]" & "}," & """type"": ""conditional""," & """expression_if_true"": {""" & "type"": ""constant""," & """constant"": 1" & "}," & """expression_if_false"": {""" & "type"": ""constant""," & """constant"": 0" & "}" & "}," & """is_nullable"": true," & """is_primary_key"": false," & ""
                    Debug.Print """create_index"": false," & """transform"": {}," & """comment"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & " """ & "}" & ","
                End If
            Next cell
        Case 2
          serialNumber = -1
         For Each cell In idRange
        
                currentID = Replace(Trim(cell.Value), ".", "_")
                If descRange.Cells(cell.Row - idRange.Row + 1, 1).Value <> "" Then
                 serialNumber = serialNumber + 1
                Debug.Print "{" & """type"": ""detail-screen-config:Column""," & """contents"": {"
                Debug.Print """hasAutocomplete"":false," & """useXpathExpression"":true," & """calc_xpath"":"".""," & """enum"":[],"
                Debug.Print """field"":""column[@id = '" & currentID & "']""," & """filter_xpath"":""""," & """format"":""plain""," & """graph_configuration"":null," & """header"":{""en"":""" & serialNumber & " - " & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """}," & """model"":""case""," & """date_format"":"""","
                Debug.Print """time_ago_interval"":365.25," & """horizontal_align"":""left""," & """vertical_align"":""start""," & """font_size"":""medium""," & """show_border"":false,"
                Debug.Print """show_shading"":false," & """late_flag"":30," & """case_tile_field"":""""," & """isTab"":false," & """hasNodeset"":false," & """nodeset"":"""","
                Debug.Print """nodesetCaseType"":""""," & """nodesetFilter"":""""," & """relevant"":""""," & """endpoint_action_id"":null," & """grid_x"":0," & """grid_y"":0,"
                Debug.Print """height"":1," & """width"":6}" & "}"

                End If
            Next cell
        Case 3
              For Each cell In idRange
            
                currentID = Replace(Trim(cell.Value), ".", "_")
                    If descRange.Cells(cell.Row - idRange.Row + 1, 1).Value <> "" Then
                Debug.Print "{" & """comment"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """," & """field"": """ & currentID & """," & """description"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """," & """format"": ""default""," & """css_class"": null," & """width"": null," & """aggregation"": ""sum""," & """column_id"": """ & currentID & """," & """visible"": true," & """transform"": {}," & """calculate_total"": false," & """type"": ""field""," & """display"":  """ & titleRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """}"
                Debug.Print "" & ","
                      End If
            Next cell
        Case Else
            MsgBox "Invalid report type selected."
    End Select
End Sub

