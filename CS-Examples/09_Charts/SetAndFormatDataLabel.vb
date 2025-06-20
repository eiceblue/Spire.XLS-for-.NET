Imports Spire.Xls

Namespace SetAndFormatDataLabel

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Create an empty worksheet in the workbook
            workbook.CreateEmptySheets(1)

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the name of the worksheet to "Demo"
            sheet.Name = "Demo"

            ' Set cell values for the month and corresponding data
            sheet.Range("A1").Value = "Month"
            sheet.Range("A2").Value = "Jan"
            sheet.Range("A3").Value = "Feb"
            sheet.Range("A4").Value = "Mar"
            sheet.Range("A5").Value = "Apr"
            sheet.Range("A6").Value = "May"
            sheet.Range("A7").Value = "Jun"
            sheet.Range("B1").Value = "Peter"
            sheet.Range("B2").NumberValue = 25
            sheet.Range("B3").NumberValue = 18
            sheet.Range("B4").NumberValue = 8
            sheet.Range("B5").NumberValue = 13
            sheet.Range("B6").NumberValue = 22
            sheet.Range("B7").NumberValue = 28

            ' Add a line chart with markers to the worksheet
            Dim chart As Chart = sheet.Charts.Add(ExcelChartType.LineMarkers)

            ' Set the data range for the chart
            chart.DataRange = sheet.Range("B1:B7")

            ' Hide the plot area of the chart
            chart.PlotArea.Visible = False

            ' Set the series data source to not use the entire row or column as labels
            chart.SeriesDataFromRange = False

            ' Set the position of the chart within the worksheet
            chart.TopRow = 5
            chart.BottomRow = 26
            chart.LeftColumn = 2
            chart.RightColumn = 11

            ' Set the title of the chart
            chart.ChartTitle = "Data Labels Demo"

            ' Customize the appearance of the chart title
            chart.ChartTitleArea.IsBold = True
            chart.ChartTitleArea.Size = 12

            ' Get the first series in the chart
            Dim cs1 As Spire.Xls.Charts.ChartSerie = chart.Series(0)

            ' Set the category labels for the series
            cs1.CategoryLabels = sheet.Range("A2:A7")

            ' Enable data labels for the default data point in the series
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasValue = True

            ' Configure data label settings for the default data point
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = False
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = False
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = True
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = True
            cs1.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". "
            cs1.DataPoints.DefaultDataPoint.DataLabels.Size = 9
            cs1.DataPoints.DefaultDataPoint.DataLabels.Color = Color.Red
            cs1.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri"
            cs1.DataPoints.DefaultDataPoint.DataLabels.Position = DataLabelPositionType.Center

            ' Specify the output file name
            Dim output As String = "SetAndFormatDataLabel.xlsx"

            ' Save the modified workbook to the specified file with Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
            ExcelDocViewer(output)
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
