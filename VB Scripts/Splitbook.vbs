Sub Splitbook()
 
Dim xPath As String
xPath = Application.ActiveWorkbook.Path
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each xWs In ThisWorkbook.Sheets
    xWs.Copy
    Application.ActiveWorkbook.SaveAs Filename:=xPath & "\Training woreda -" & xWs.Name & ".xlsx"
    Application.ActiveWorkbook.Close False
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub CopySheets()
    Dim sh As Worksheet
    Dim wkbTarget As Workbook
    Dim i, arrNames
    Dim cell As Range, xVRg As Range
   Set xVRg = Application.InputBox("Please select the column that contains the Sheet NAME ", "Excel", "", Type:=8)
    Set wkbTarget = Workbooks.Add()
     Set sh = ThisWorkbook.Sheets("Menus_and_forms")
     sh.Copy After:=wkbTarget.Sheets(1)
    For Each cell In xVRg
      Set sh = Nothing
            On Error Resume Next
            Set sh = ThisWorkbook.Sheets(cell.Value)
            On Error GoTo 0
            If Not sh Is Nothing Then
                sh.Copy After:=wkbTarget.Sheets(2)
            End If
    Next cell
      wkbTarget.Sheets("Sheet1").Delete
End Sub