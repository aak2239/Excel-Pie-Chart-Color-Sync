Sub MatchAllPieSliceColors()
    ' Declare variables for worksheets and chart objects
    Dim ws As Worksheet
    Dim chrt As ChartObject
    Dim seriesFormula As String
    Dim valuesRangeAddress As String
    Dim valuesRange As Range
    Dim spaceUseSheet As Worksheet
    Dim i As Integer

    ' Set a reference to the 'Space Use' sheet.
    ' This is where the script will look for color references.
    ' Modify this according to your workbook's structure.
    Set spaceUseSheet = ThisWorkbook.Sheets("Space Use")

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through each chart object in the worksheet
        For Each chrt In ws.ChartObjects
            ' Check if the chart is a pie chart
            If chrt.Chart.ChartType <> xlPie Then
                ' Skip if not a pie chart and go to the next chart
                GoTo NextChart
            End If

            ' Attempt to get the series formula of the chart
            On Error Resume Next
            seriesFormula = chrt.Chart.SeriesCollection(1).Formula
            ' Handle potential errors in accessing the chart series
            If Err.Number <> 0 Then
                MsgBox "Error accessing series for chart on " & ws.Name
                Err.Clear
                GoTo NextChart
            End If
            On Error GoTo 0

            ' Extract the range address for the values from the series formula
            valuesRangeAddress = Split(Split(seriesFormula, "!")(1), ",")(0)

            ' Attempt to set the values range
            On Error Resume Next
            Set valuesRange = spaceUseSheet.Range(valuesRangeAddress)
            ' Handle potential errors in setting the values range
            If Err.Number <> 0 Then
                MsgBox "Error setting values range for chart on " & ws.Name
                Err.Clear
                GoTo NextChart
            End If
            On Error GoTo 0

            ' Apply the fill colors to the chart slices
            ' Loop through each cell in the values range
            For i = 1 To valuesRange.Cells.Count
                ' Set the color of the pie slice to match the cell's interior color
                chrt.Chart.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = _
                    valuesRange.Cells(i, 1).Interior.Color
            Next i

NextChart:
        ' Continue with the next chart
        Next chrt
    ' Continue with the next worksheet
    Next ws
End Sub
