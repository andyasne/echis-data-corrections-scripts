Function BuildReportText(rng As Range) As String
    Dim yellowCell As Range
        Dim idCell As Range
        Dim categoryCell As Range
    Dim blackCell As Range
    Dim i As Integer

    For i = rng.Row To 1 Step -1
        If Cells(i, rng.Column).Interior.Color = RGB(238, 238, 238) Or Cells(i, rng.Column).Interior.Color = RGB(204, 204, 204) Then
            Set blackCell = Cells(i, rng.Column)
            Exit For
        End If
    Next i

 
    Set yellowCell = Cells(rng.Row, rng.Column)
    Set idCell = Cells(rng.Row, rng.Column - 1)
       Set categoryCell = Cells(rng.Row, rng.Column + 1)

    If Not yellowCell Is Nothing And Not blackCell Is Nothing Then
        ConcatenateBlackYellow = blackCell.Value & " - " & yellowCell.Value & "(" & idCell.Value & "):" & categoryCell.Value
        
    Else
        ConcatenateBlackYellow = "CHECK"
    End If
End Function





