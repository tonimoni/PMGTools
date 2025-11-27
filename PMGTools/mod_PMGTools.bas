Attribute VB_Name = "mod_PMGTools"

Sub ConvertChartDataToValues()
    ' A VBA macro to iterate through all charts in a PowerPoint presentation,
    ' access their underlying data, and paste the data as values to remove formulas.

    Dim sld As Slide
    Dim shp As Shape
    Dim cht As Chart
    Dim chtData As ChartData
    Dim wb As Object ' Using Object for late binding to Excel.Workbook
    Dim ws As Object ' Using Object for late binding to Excel.Worksheet
    Dim usedDataRange As Object ' Using Object for late binding to Excel.Range

    ' Loop through each slide in the active presentation
    For Each sld In ActivePresentation.Slides
        ' Loop through each shape on the slide
        For Each shp In sld.Shapes
            ' Check if the shape has a chart
            If shp.HasChart Then
                Set cht = shp.Chart
                Set chtData = cht.ChartData

                ' Activate the chart's data workbook to work with it
                On Error Resume Next ' In case the workbook is already active or other issues
                chtData.Activate
                Set wb = chtData.Workbook
                If wb Is Nothing Then
                    Err.Clear
                    On Error GoTo 0
                    GoTo NextShape ' Skip if workbook can't be accessed
                End If
                On Error GoTo 0

                ' Get the first worksheet in the workbook
                Set ws = wb.Worksheets(1)

                ' Determine the used range of the data
                Set usedDataRange = ws.usedRange

                ' Copy the data to the clipboard
                usedDataRange.Copy

                ' Paste the data back as values
                ' xlPasteValues is a constant in the Excel object model (value = -4163)
                ' Using the value directly avoids needing a direct reference to the Excel library
                usedDataRange.PasteSpecial Paste:=-4163, Operation:=-4142  ' xlNone

                ' Clear the clipboard
                wb.Application.CutCopyMode = False

                ' Close the workbook without saving changes, as the changes are in the chart
                wb.Close SaveChanges:=False

                ' Release the objects
                Set usedDataRange = Nothing
                Set ws = Nothing
                Set wb = Nothing
                Set chtData = Nothing
                Set cht = Nothing
            End If
NextShape:
        Next shp
    Next sld

    MsgBox "All chart data has been converted to values.", vbInformation

End Sub

Sub Print_messsage()
    MsgBox "All chart data has been converted to values.", vbInformation

End Sub

