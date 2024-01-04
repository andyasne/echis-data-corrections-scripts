Sub BuildReport()
    ' Declare variables
    Dim idRange As Range, titleRange As Range, descRange As Range
    Dim reportType As Integer, cell As Range

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
                Dim currentID As String
                currentID = Replace(Trim(cell.Value), ".", "_")
                If descRange.Cells(cell.Row - idRange.Row + 1, 1).Value <> "" Then
                    Debug.Print "{" & """type"": ""expression""," & """column_id"": """ & currentID & """," & """datatype"": ""integer""," & """display_name"": """ & titleRange.Cells(cell.Row - idRange.Row + 1, 1).Value & """" & ","
                    Debug.Print """expression"": {""" & "datatype"": ""integer""," & """test"": {""" & "type"": ""and""," & """filters"": [{" & """operator"": ""eq""," & """type"": ""boolean_expression""," & """expression"": {""" & "datatype"": ""string""," & """type"": ""property_name""," & """property_name"": ""indicator_name""" & "}," & """comment"": null," & """property_value"": ""Xindicator_name""" & "},{" & """operator"": ""eq""," & """expression"": {""" & "type"": ""root_doc""," & """expression"": {""" & "datatype"": ""string""," & """type"": ""property_name""," & """property_name"": ""property_name""" & "}" & "}," & """type"": ""boolean_expression""," & """comment"": null," & """property_value"": ""xproperty_name """ & "}]" & "}," & """type"": ""conditional""," & """expression_if_true"": {""" & "type"": ""constant""," & """constant"": 1" & "}," & """expression_if_false"": {""" & "type"": ""constant""," & """constant"": 0" & "}" & "}," & """is_nullable"": true," & """is_primary_key"": false," & ""
                    Debug.Print """create_index"": false," & """transform"": {}," & """comment"": """ & descRange.Cells(cell.Row - idRange.Row + 1, 1).Value & " """ & "}" & ","
                End If
            Next cell
        Case 2
            'Add your code for mobile report generation
        Case 3
            'Add your code for web report generation
        Case Else
            MsgBox "Invalid report type selected."
    End Select
End Sub

